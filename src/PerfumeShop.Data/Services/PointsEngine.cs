using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// V18 Points and Rewards Engine - Full alignment with points_engine.asp (20.5KB)
/// Features: earn/redeem/exchange/expire/sign-in/rule cache/balance sync
/// V19 improvement: 3-table sync write (PointsLedger + UserPoints + Users.Points) wrapped in EF Core Transaction
/// Note: All DB operations use LINQ (no raw SQL) for InMemory Provider test compatibility
/// </summary>
public class PointsEngine : IPointsEngine
{
    private readonly PerfumeShopContext _db;
    private readonly IMemoryCache _cache;

    private const string PE_RuleCacheKey = "PE_RuleCache";
    private static readonly TimeSpan RuleCacheDuration = TimeSpan.FromMinutes(30);

    // V18 default rule fallback values - maps to PE_GetRule() Select Case defaults
    private static readonly Dictionary<string, decimal> DefaultRules = new(StringComparer.OrdinalIgnoreCase)
    {
        ["purchase_rate"] = 1,
        ["signin_points"] = 5,
        ["review_points"] = 20,
        ["review_with_photo"] = 10,
        ["share_points"] = 10,
        ["referral_points"] = 100,
        ["referral_purchase"] = 50,
        ["redeem_discount_rate"] = 100,
        ["max_redeem_pct"] = 30,
        ["points_expire_months"] = 12
    };

    public PointsEngine(PerfumeShopContext db, IMemoryCache cache)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
        _cache = cache ?? throw new ArgumentNullException(nameof(cache));
    }

    // ==================== Rule Cache ====================

    public async Task RefreshRuleCacheAsync(CancellationToken ct = default)
    {
        var rules = await _db.PointsRules
            .AsNoTracking()
            .Where(r => r.IsEnabled)
            .ToListAsync(ct);

        var cache = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        foreach (var rule in rules)
        {
            var code = rule.RuleCode?.ToLowerInvariant() ?? "";
            if (!string.IsNullOrEmpty(code))
                cache.TryAdd(code, rule.RuleValue);
        }

        _cache.Set(PE_RuleCacheKey, cache, RuleCacheDuration);
    }

    private async Task<Dictionary<string, decimal>> GetRuleCacheAsync(CancellationToken ct)
    {
        if (_cache.TryGetValue<Dictionary<string, decimal>>(PE_RuleCacheKey, out var cache) && cache != null)
            return cache;

        await RefreshRuleCacheAsync(ct);
        return _cache.TryGetValue<Dictionary<string, decimal>>(PE_RuleCacheKey, out var fresh) && fresh != null
            ? fresh
            : new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
    }

    public async Task<decimal> GetRuleAsync(string ruleCode, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(ruleCode)) return 0;

        var code = ruleCode.ToLowerInvariant();
        var cache = await GetRuleCacheAsync(ct);

        if (cache.TryGetValue(code, out var val))
            return val;

        return DefaultRules.TryGetValue(code, out var defaultVal) ? defaultVal : 0;
    }

    // ==================== Core: Earn Points ====================

    public async Task<bool> EarnAsync(int userId, int points, string source, string? description = null, int? referenceId = null, CancellationToken ct = default)
    {
        if (points <= 0) return false;

        var expireMonths = (int)await GetRuleAsync("points_expire_months", ct);
        DateTime? expiresAt = expireMonths > 0 ? DateTime.Now.AddMonths(expireMonths) : null;

        using var transaction = await _db.Database.BeginTransactionAsync(ct);
        try
        {
            // 1. Insert PointsLedger
            var ledger = new PointsLedger
            {
                UserId = userId,
                Points = points,
                PointType = "earn",
                Source = source,
                ReferenceId = referenceId,
                Description = description,
                ExpiresAt = expiresAt,
                IsExpired = false,
                CreatedAt = DateTime.Now
            };
            _db.PointsLedgers.Add(ledger);
            await _db.SaveChangesAsync(ct);

            // 2. Upsert UserPoints - LINQ instead of ExecuteSqlRaw
            var userPoint = await _db.UserPoints.FirstOrDefaultAsync(up => up.UserId == userId, ct);
            if (userPoint != null)
            {
                userPoint.AvailablePoints = (userPoint.AvailablePoints ?? 0) + points;
                userPoint.TotalPoints = (userPoint.TotalPoints ?? 0) + points;
                userPoint.LastUpdatedAt = DateTime.Now;
            }
            else
            {
                _db.UserPoints.Add(new UserPoint
                {
                    UserId = userId,
                    AvailablePoints = points,
                    TotalPoints = points,
                    UsedPoints = 0,
                    ExpiredPoints = 0,
                    LastUpdatedAt = DateTime.Now
                });
            }

            // 3. Sync Users.Points - LINQ instead of ExecuteSqlRaw
            var user = await _db.Users.FirstOrDefaultAsync(u => u.UserId == userId, ct);
            if (user != null)
            {
                user.Points = (user.Points ?? 0) + points;
            }

            await _db.SaveChangesAsync(ct);
            await transaction.CommitAsync(ct);
            return true;
        }
        catch
        {
            await transaction.RollbackAsync(ct);
            return false;
        }
    }

    // ==================== Core: Redeem Points ====================

    public async Task<bool> RedeemAsync(int userId, int points, string redemptionType, int? referenceId = null, CancellationToken ct = default)
    {
        if (points <= 0) return false;

        var available = await GetBalanceAsync(userId, ct);
        if (points > available)
        {
            points = available;
            if (points <= 0) return false;
        }

        var description = $"{redemptionType} redemption ({points} pts)";

        using var transaction = await _db.Database.BeginTransactionAsync(ct);
        try
        {
            var ledger = new PointsLedger
            {
                UserId = userId,
                Points = -points,
                PointType = "redeem",
                Source = redemptionType,
                ReferenceId = referenceId,
                Description = description,
                IsExpired = false,
                CreatedAt = DateTime.Now
            };
            _db.PointsLedgers.Add(ledger);
            await _db.SaveChangesAsync(ct);

            // Sync UserPoints - LINQ
            var userPoint = await _db.UserPoints.FirstOrDefaultAsync(up => up.UserId == userId, ct);
            if (userPoint != null)
            {
                userPoint.AvailablePoints = (userPoint.AvailablePoints ?? 0) - points;
                userPoint.UsedPoints = (userPoint.UsedPoints ?? 0) + points;
                userPoint.LastUpdatedAt = DateTime.Now;
            }

            // Sync Users.Points - LINQ
            var user = await _db.Users.FirstOrDefaultAsync(u => u.UserId == userId, ct);
            if (user != null)
            {
                user.Points = (user.Points ?? 0) - points;
            }

            await _db.SaveChangesAsync(ct);
            await transaction.CommitAsync(ct);
            return true;
        }
        catch
        {
            await transaction.RollbackAsync(ct);
            return false;
        }
    }

    // ==================== Get Balance ====================

    public async Task<int> GetBalanceAsync(int userId, CancellationToken ct = default)
    {
        await ApplyExpirationAsync(userId, ct);

        var pts = await _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId && !l.IsExpired)
            .SumAsync(l => (int?)l.Points, ct) ?? 0;

        if (pts == 0)
        {
            var up = await _db.UserPoints.FirstOrDefaultAsync(x => x.UserId == userId, ct);
            if (up != null)
                pts = up.AvailablePoints ?? 0;

            if (pts == 0)
            {
                pts = await _db.Users
                    .AsNoTracking()
                    .Where(u => u.UserId == userId)
                    .Select(u => u.Points ?? 0)
                    .FirstOrDefaultAsync(ct);
            }
        }

        return pts;
    }

    // ==================== Expiration ====================

    public async Task ApplyExpirationAsync(int userId, CancellationToken ct = default)
    {
        var expireMonths = await GetRuleAsync("points_expire_months", ct);
        if (expireMonths <= 0) return;

        var now = DateTime.Now;
        var expiredLedgers = await _db.PointsLedgers
            .Where(l => l.UserId == userId
                     && !l.IsExpired
                     && l.PointType == "earn"
                     && l.ExpiresAt != null
                     && l.ExpiresAt < now)
            .ToListAsync(ct);

        if (expiredLedgers.Count > 0)
        {
            foreach (var ledger in expiredLedgers)
                ledger.IsExpired = true;
            await _db.SaveChangesAsync(ct);
        }
    }

    // ==================== Balance Sync ====================

    public async Task UpdateBalanceAsync(int userId, CancellationToken ct = default)
    {
        await ApplyExpirationAsync(userId, ct);

        var totalEarned = await _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId && l.PointType == "earn" && !l.IsExpired)
            .SumAsync(l => (int?)l.Points, ct) ?? 0;

        var totalUsed = await _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId && l.PointType == "redeem")
            .SumAsync(l => (int?)Math.Abs(l.Points), ct) ?? 0;

        var totalExpired = await _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId && l.PointType == "earn" && l.IsExpired && l.Points > 0)
            .SumAsync(l => (int?)l.Points, ct) ?? 0;

        var available = totalEarned - totalUsed;
        if (available < 0) available = 0;

        using var transaction = await _db.Database.BeginTransactionAsync(ct);
        try
        {
            var userPoint = await _db.UserPoints.FirstOrDefaultAsync(up => up.UserId == userId, ct);
            if (userPoint != null)
            {
                userPoint.AvailablePoints = available;
                userPoint.UsedPoints = totalUsed;
                userPoint.ExpiredPoints = totalExpired;
                userPoint.TotalPoints = totalEarned;
                userPoint.LastUpdatedAt = DateTime.Now;
            }
            else
            {
                _db.UserPoints.Add(new UserPoint
                {
                    UserId = userId,
                    AvailablePoints = available,
                    TotalPoints = totalEarned,
                    UsedPoints = totalUsed,
                    ExpiredPoints = totalExpired,
                    LastUpdatedAt = DateTime.Now
                });
            }

            var user = await _db.Users.FirstOrDefaultAsync(u => u.UserId == userId, ct);
            if (user != null)
            {
                user.Points = available;
            }

            await _db.SaveChangesAsync(ct);
            await transaction.CommitAsync(ct);
        }
        catch
        {
            await transaction.RollbackAsync(ct);
        }
    }

    // ==================== Points Calculation ====================

    public async Task<decimal> CalcPointsValueAsync(int points, CancellationToken ct = default)
    {
        var rate = await GetRuleAsync("redeem_discount_rate", ct);
        if (rate <= 0) rate = 100;
        return (decimal)points / rate;
    }

    public async Task<int> CalcOrderPointsAsync(decimal orderAmount, CancellationToken ct = default)
    {
        var rate = await GetRuleAsync("purchase_rate", ct);
        if (rate <= 0) rate = 1;
        return (int)(orderAmount * rate);
    }

    public async Task<decimal> GetMaxRedeemPctAsync(CancellationToken ct = default)
    {
        return await GetRuleAsync("max_redeem_pct", ct);
    }

    public async Task<int> CalcMaxRedeemablePointsAsync(int userId, decimal orderAmount, CancellationToken ct = default)
    {
        var rate = await GetRuleAsync("redeem_discount_rate", ct);
        if (rate <= 0) rate = 100;

        var maxPct = (await GetRuleAsync("max_redeem_pct", ct)) / 100m;
        if (maxPct <= 0) maxPct = 0.3m;

        var maxValue = orderAmount * maxPct;
        return (int)(maxValue * rate);
    }

    // ==================== Sign-In ====================

    public async Task<bool> CheckSignInAsync(int userId, CancellationToken ct = default)
    {
        var today = DateTime.Today;
        var count = await _db.PointsLedgers
            .AsNoTracking()
            .CountAsync(l => l.UserId == userId
                          && l.Source == "signin"
                          && l.CreatedAt >= today
                          && l.CreatedAt < today.AddDays(1), ct);
        return count > 0;
    }

    public async Task<int> DoSignInAsync(int userId, CancellationToken ct = default)
    {
        if (await CheckSignInAsync(userId, ct))
            return -1;

        var points = (int)await GetRuleAsync("signin_points", ct);
        if (points <= 0) points = 5;

        var success = await EarnAsync(userId, points, "signin", "Daily sign-in", 0, ct);
        return success ? points : 0;
    }

    // ==================== Ledger and Summary ====================

    public async Task<PointsLedgerResult> GetLedgerAsync(int userId, int page = 1, int pageSize = 20, CancellationToken ct = default)
    {
        if (page < 1) page = 1;
        if (pageSize < 1) pageSize = 20;

        var query = _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId)
            .OrderByDescending(l => l.CreatedAt);

        var total = await query.CountAsync(ct);
        var items = await query
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .Select(l => new PointsLedgerItem
            {
                LedgerId = l.LedgerId,
                UserId = l.UserId,
                Points = l.Points,
                PointType = l.PointType,
                Source = l.Source,
                ReferenceId = l.ReferenceId,
                Description = l.Description,
                ExpiresAt = l.ExpiresAt,
                IsExpired = l.IsExpired,
                CreatedAt = l.CreatedAt
            })
            .ToListAsync(ct);

        return new PointsLedgerResult { Items = items, Total = total };
    }

    public async Task<PointsSummaryDto> GetPointsSummaryAsync(int userId, CancellationToken ct = default)
    {
        var available = await GetBalanceAsync(userId, ct);

        var totalEarned = await _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId && l.PointType == "earn" && !l.IsExpired)
            .SumAsync(l => (int?)l.Points, ct) ?? 0;

        var totalRedeemed = await _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId && l.PointType == "redeem")
            .SumAsync(l => (int?)Math.Abs(l.Points), ct) ?? 0;

        var today = DateTime.Today;
        var todayEarned = await _db.PointsLedgers
            .AsNoTracking()
            .Where(l => l.UserId == userId && l.PointType == "earn"
                     && l.CreatedAt >= today && l.CreatedAt < today.AddDays(1))
            .SumAsync(l => (int?)l.Points, ct) ?? 0;

        var expiringSoon = 0;
        var expireMonths = await GetRuleAsync("points_expire_months", ct);
        if (expireMonths > 0)
        {
            var deadline = DateTime.Now.AddDays(30);
            expiringSoon = await _db.PointsLedgers
                .AsNoTracking()
                .Where(l => l.UserId == userId && l.PointType == "earn" && !l.IsExpired
                         && l.ExpiresAt != null && l.ExpiresAt <= deadline)
                .SumAsync(l => (int?)l.Points, ct) ?? 0;
        }

        return new PointsSummaryDto
        {
            Available = available,
            TotalEarned = totalEarned,
            TotalRedeemed = totalRedeemed,
            TodayEarned = todayEarned,
            ExpiringSoon = expiringSoon
        };
    }

    public async Task<OrderPointsDto> GetOrderPointsAsync(int orderId, CancellationToken ct = default)
    {
        var result = new OrderPointsDto();

        var orderPoints = await _db.Orders
            .AsNoTracking()
            .Where(o => o.OrderId == orderId)
            .Select(o => new
            {
                o.PointsEarned,
                o.PointsRedeemed,
                o.PointsDiscount
            })
            .FirstOrDefaultAsync(ct);

        if (orderPoints != null)
        {
            result.Earned = orderPoints.PointsEarned ?? 0;
            result.Redeemed = orderPoints.PointsRedeemed ?? 0;
            result.Discount = orderPoints.PointsDiscount ?? 0;
        }

        if (result.Earned == 0)
        {
            result.Earned = await _db.PointsLedgers
                .AsNoTracking()
                .Where(l => l.UserId > 0 && l.ReferenceId == orderId && l.Source == "purchase")
                .SumAsync(l => (int?)l.Points, ct) ?? 0;
        }

        return result;
    }

    // ==================== Redemption Shop ====================

    public async Task<string> DoRedeemAsync(int userId, int redemptionId, CancellationToken ct = default)
    {
        var item = await _db.PointsRedemptions
            .FirstOrDefaultAsync(r => r.RedemptionId == redemptionId && r.IsEnabled, ct);

        if (item == null)
            return "Redemption item not found or sold out";

        if (item.Stock <= 0)
            return "Redemption item is out of stock";

        var available = await GetBalanceAsync(userId, ct);
        if (available < item.PointsCost)
            return $"Insufficient points (need {item.PointsCost}, have {available})";

        var success = await RedeemAsync(userId, item.PointsCost, item.ItemType, redemptionId, ct);
        if (!success)
            return "Points deduction failed, please try again";

        // Deduct stock - LINQ instead of ExecuteUpdateAsync
        item.Stock -= 1;
        await _db.SaveChangesAsync(ct);

        return ""; // success
    }
}
