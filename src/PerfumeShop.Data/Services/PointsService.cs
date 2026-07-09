using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

public class PointsService : IPointsService
{
    private readonly PerfumeShopContext _db;
    public PointsService(PerfumeShopContext db) => _db = db;

    public async Task<PointsBalanceDto> GetBalanceAsync(int userId)
    {
        // UserPoint is keyless — use FirstOrDefaultAsync
        var up = await _db.UserPoints.FirstOrDefaultAsync(x => x.UserId == userId);
        if (up == null)
        {
            return new PointsBalanceDto { UserId = userId };
        }
        return new PointsBalanceDto
        {
            UserId = userId,
            AvailablePoints = up.AvailablePoints ?? 0,
            TotalPoints = up.TotalPoints ?? 0,
            UsedPoints = up.UsedPoints ?? 0,
            ExpiredPoints = up.ExpiredPoints ?? 0,
            LastUpdatedAt = up.LastUpdatedAt
        };
    }

    public async Task<(List<PointsLedger> Items, int Total)> GetLedgerAsync(int userId, int page = 1, int pageSize = 20)
    {
        var query = _db.PointsLedgers
            .Where(l => l.UserId == userId)
            .OrderByDescending(l => l.CreatedAt);

        var total = await query.CountAsync();
        var items = await query.Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }

    public async Task<int> AwardPointsAsync(int userId, int points, string source, string? description = null, int? referenceId = null)
    {
        if (points <= 0) return 0;

        // 记录流水
        var ledger = new PointsLedger
        {
            UserId = userId,
            Points = points,
            PointType = "earn",
            Source = source,
            ReferenceId = referenceId,
            Description = description,
            IsExpired = false,
            CreatedAt = DateTime.Now
        };
        _db.PointsLedgers.Add(ledger);

        // 更新余额 (UserPoint is keyless — use raw SQL or find-or-create pattern)
        await _db.SaveChangesAsync();
        await UpdateUserPointsBalanceAsync(userId, points, 0);

        return points;
    }

    public async Task<bool> DeductPointsAsync(int userId, int points, string source, string? description = null, int? referenceId = null)
    {
        if (points <= 0) return false;

        var balance = await GetBalanceAsync(userId);
        if (balance.AvailablePoints < points) return false;

        var ledger = new PointsLedger
        {
            UserId = userId,
            Points = -points,
            PointType = "spend",
            Source = source,
            ReferenceId = referenceId,
            Description = description,
            IsExpired = false,
            CreatedAt = DateTime.Now
        };
        _db.PointsLedgers.Add(ledger);
        await _db.SaveChangesAsync();
        await UpdateUserPointsBalanceAsync(userId, 0, points);

        return true;
    }

    public async Task<List<PointsRedemption>> GetRedemptionItemsAsync()
    {
        return await _db.PointsRedemptions
            .Where(r => r.IsEnabled && r.Stock > 0)
            .OrderBy(r => r.SortOrder)
            .ToListAsync();
    }

    public async Task<RedeemResult> RedeemAsync(int userId, int redemptionId)
    {
        var item = await _db.PointsRedemptions.FirstOrDefaultAsync(r => r.RedemptionId == redemptionId && r.IsEnabled);
        if (item == null) return new() { Success = false, Message = "兑换商品不存在" };
        if (item.Stock <= 0) return new() { Success = false, Message = "库存不足" };

        var balance = await GetBalanceAsync(userId);
        if (balance.AvailablePoints < item.PointsCost)
            return new() { Success = false, Message = $"积分不足，需要{item.PointsCost}积分，当前可用{balance.AvailablePoints}积分" };

        var success = await DeductPointsAsync(userId, item.PointsCost, "redemption", $"兑换: {item.ItemName}", redemptionId);
        if (!success) return new() { Success = false, Message = "积分扣减失败" };

        item.Stock--;
        await _db.SaveChangesAsync();

        var newBalance = await GetBalanceAsync(userId);
        return new() { Success = true, Message = $"兑换成功: {item.ItemName}", PointsRemaining = newBalance.AvailablePoints };
    }

    public async Task<List<PointsRule>> GetRulesAsync()
    {
        return await _db.PointsRules
            .Where(r => r.IsEnabled)
            .OrderBy(r => r.SortOrder)
            .ToListAsync();
    }

    /// <summary>更新用户积分余额 (UserPoint is keyless, 需要特殊处理)</summary>
    private async Task UpdateUserPointsBalanceAsync(int userId, int addPoints, int deductPoints)
    {
        // UserPoint 是 keyless 实体，直接使用原始 SQL 更新
        // 如果表有 PointId 作为标识但标记为 keyless，使用存储过程或原始 SQL
        try
        {
            if (addPoints > 0)
            {
                await _db.Database.ExecuteSqlRawAsync(
                    "UPDATE UserPoints SET AvailablePoints = ISNULL(AvailablePoints,0) + {0}, TotalPoints = ISNULL(TotalPoints,0) + {0}, LastUpdatedAt = GETDATE() WHERE UserId = {1}",
                    addPoints, userId);
            }
            if (deductPoints > 0)
            {
                await _db.Database.ExecuteSqlRawAsync(
                    "UPDATE UserPoints SET AvailablePoints = ISNULL(AvailablePoints,0) - {0}, UsedPoints = ISNULL(UsedPoints,0) + {0}, LastUpdatedAt = GETDATE() WHERE UserId = {1}",
                    deductPoints, userId);
            }
        }
        catch
        {
            // UserPoints 表可能不存在该用户记录 — 忽略
        }
    }
}
