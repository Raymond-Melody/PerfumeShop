using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 促销引擎 — 优惠券验证、折扣计算、发放
/// 对应 ASP 中的 promotion_engine.asp
/// </summary>
public class PromotionEngine : IPromotionEngine
{
    private readonly PerfumeShopContext _db;
    private const decimal FreeShippingThreshold = 299m;
    private const decimal MaxDiscountRatio = 0.5m;

    public PromotionEngine(PerfumeShopContext db)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
    }

    // ==================== 优惠券验证 ====================

    public async Task<CouponValidateResult> ValidateCouponAsync(string code, int userId, decimal cartTotal, CancellationToken ct = default)
    {
        var result = new CouponValidateResult();

        if (string.IsNullOrEmpty(code))
        {
            result.Message = "请输入优惠码";
            return result;
        }

        var coupon = await _db.Coupons
            .AsNoTracking()
            .FirstOrDefaultAsync(c => c.CouponCode == code && c.IsActive == true, ct);

        if (coupon == null)
        {
            result.Message = "优惠码不存在或已失效";
            return result;
        }

        // 有效期检查
        var now = DateTime.Now;
        if (now < coupon.ValidFrom || now > coupon.ValidTo)
        {
            result.Message = "优惠码不在有效期内";
            return result;
        }

        // 库存检查
        if (coupon.TotalQty > 0 && coupon.UsedQty >= coupon.TotalQty)
        {
            result.Message = "优惠码已被领完";
            return result;
        }

        // 最低消费检查
        if (coupon.MinSpend > 0 && cartTotal < coupon.MinSpend)
        {
            result.Message = $"未达最低消费 ¥{coupon.MinSpend:F0}（当前 ¥{cartTotal:F2}）";
            return result;
        }

        // 首单检查
        if (coupon.FirstOrderOnly)
        {
            var orderCount = await _db.Orders.CountAsync(o => o.UserId == userId, ct);
            if (orderCount > 0)
            {
                result.Message = "此优惠券仅限首单使用";
                return result;
            }
        }

        // 检查用户是否已使用过此券
        var usedCount = await _db.UserCoupons
            .CountAsync(uc => uc.UserId == userId && uc.CouponCode == code && uc.Status == "used", ct);

        if (usedCount > 0)
        {
            result.Message = "该优惠码您已使用过";
            return result;
        }

        // 计算折扣金额
        var discount = CalculateCouponDiscount(coupon.CouponType, coupon.DiscountValue ?? 0, cartTotal, coupon.MaxDiscount);

        result.Valid = true;
        result.Message = "优惠码验证通过";
        result.Discount = discount;
        result.Type = coupon.CouponType;
        return result;
    }

    // ==================== 折扣计算 ====================

    public async Task<decimal> CalculateDiscountAsync(decimal cartTotal, int userId, CancellationToken ct = default)
    {
        decimal discount = 0;

        // 1. 满减折扣
        var thresholdSetting = await _db.SiteSettings
            .AsNoTracking()
            .FirstOrDefaultAsync(s => s.SettingKey == "Promotion_Threshold", ct);

        if (thresholdSetting != null && !string.IsNullOrEmpty(thresholdSetting.SettingValue))
        {
            var parts = thresholdSetting.SettingValue.Split('|');
            if (parts.Length >= 2 && decimal.TryParse(parts[0], out var threshold) &&
                decimal.TryParse(parts[1], out var thresholdDiscount))
            {
                if (cartTotal >= threshold)
                {
                    discount += thresholdDiscount;
                }
            }
        }

        // 2. 会员等级折扣
        // 简化实现：根据 CustomerTier 查询折扣率
        var user = await _db.Users
            .AsNoTracking()
            .Select(u => new { u.UserId, u.CustomerTier })
            .FirstOrDefaultAsync(u => u.UserId == userId, ct);

        if (user != null && !string.IsNullOrEmpty(user.CustomerTier))
        {
            var tier = await _db.MemberTiers
                .AsNoTracking()
                .FirstOrDefaultAsync(mt => mt.TierCode == user.CustomerTier, ct);

            if (tier != null && tier.DiscountRate > 0)
            {
                discount += cartTotal * (1m - tier.DiscountRate);
            }
        }

        // 3. 首单折扣
        var orderCount = await _db.Orders.CountAsync(o => o.UserId == userId, ct);
        if (orderCount <= 1)
        {
            var firstOrderSetting = await _db.SiteSettings
                .AsNoTracking()
                .FirstOrDefaultAsync(s => s.SettingKey == "Promotion_FirstOrder", ct);

            if (firstOrderSetting != null && decimal.TryParse(firstOrderSetting.SettingValue, out var firstDiscount))
            {
                if (firstDiscount > 0)
                {
                    discount += cartTotal * firstDiscount / 100m;
                }
            }
        }

        // 最高50%折扣上限
        if (discount > cartTotal * MaxDiscountRatio)
            discount = cartTotal * MaxDiscountRatio;

        return discount;
    }

    public bool CheckFreeShipping(decimal cartTotal, decimal freeShippingThreshold = FreeShippingThreshold)
    {
        return cartTotal >= freeShippingThreshold;
    }

    // ==================== 优惠券使用 ====================

    public async Task<bool> UseCouponAsync(string code, int userId, int orderId, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(code)) return false;

        try
        {
            // 标记用户券为已使用
            await _db.UserCoupons
                .Where(uc => uc.UserId == userId && uc.CouponCode == code && uc.Status == "available")
                .ExecuteUpdateAsync(setters => setters
                    .SetProperty(uc => uc.Status, "used")
                    .SetProperty(uc => uc.UsedAt, DateTime.Now)
                    .SetProperty(uc => uc.UsedOrderId, orderId), ct);

            // 增加券的已使用计数
            await _db.Coupons
                .Where(c => c.CouponCode == code)
                .ExecuteUpdateAsync(setters => setters
                    .SetProperty(c => c.UsedQty, c => c.UsedQty + 1), ct);

            return true;
        }
        catch
        {
            return false;
        }
    }

    // ==================== 优惠券发放 ====================

    public async Task<bool> IssueCouponAsync(int userId, string code, string source, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(code) || userId <= 0) return false;

        try
        {
            var coupon = await _db.Coupons
                .AsNoTracking()
                .FirstOrDefaultAsync(c => c.CouponCode == code && c.IsActive == true, ct);

            if (coupon == null) return false;

            // 库存检查
            if (coupon.TotalQty > 0 && coupon.UsedQty >= coupon.TotalQty) return false;

            // 检查用户是否已有此券且未使用
            var hasAvailable = await _db.UserCoupons
                .AnyAsync(uc => uc.UserId == userId && uc.CouponCode == code && uc.Status == "available", ct);

            if (hasAvailable) return false; // 已有可用券，不重复发放

            // 发放
            var userCoupon = new UserCoupon
            {
                UserId = userId,
                CouponId = coupon.CouponId,
                CouponCode = code,
                Source = source,
                Status = "available",
                ObtainedAt = DateTime.Now,
                ExpiresAt = coupon.ValidTo
            };

            _db.UserCoupons.Add(userCoupon);

            // 更新库存
            await _db.Coupons
                .Where(c => c.CouponId == coupon.CouponId)
                .ExecuteUpdateAsync(setters => setters
                    .SetProperty(c => c.UsedQty, c => c.UsedQty + 1), ct);

            await _db.SaveChangesAsync(ct);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<int> IssueWelcomeCouponsAsync(int userId, CancellationToken ct = default)
    {
        int count = 0;

        // 发放新人专属券
        if (await IssueCouponAsync(userId, "WELCOME10", "new_user", ct)) count++;
        if (await IssueCouponAsync(userId, "WELCOME20", "new_user", ct)) count++;

        // 检查是否有公开可领的免邮券
        var freeShipCoupons = await _db.Coupons
            .AsNoTracking()
            .Where(c => c.CouponType == "free_shipping" && c.IsActive == true && c.IsPublic == true &&
                        DateTime.Now >= c.ValidFrom && DateTime.Now <= c.ValidTo)
            .Select(c => c.CouponCode)
            .ToListAsync(ct);

        foreach (var freeCode in freeShipCoupons)
        {
            if (!string.IsNullOrEmpty(freeCode) && await IssueCouponAsync(userId, freeCode, "new_user", ct)) count++;
        }

        return count;
    }

    public async Task<bool> IssueTierUpgradeCouponAsync(int userId, string tierCode, CancellationToken ct = default)
    {
        return tierCode?.ToLower() switch
        {
            "gold" => await IssueCouponAsync(userId, "TIER_GOLD", "tier_upgrade", ct),
            "diamond" or "black" => await IssueCouponAsync(userId, "VIP5", "tier_upgrade", ct),
            _ => false
        };
    }

    // ==================== 查询 ====================

    public async Task<int> GetUserCouponCountAsync(int userId, CancellationToken ct = default)
    {
        return await _db.UserCoupons
            .CountAsync(uc => uc.UserId == userId && uc.Status == "available", ct);
    }

    // ==================== 辅助 ====================

    /// <summary>计算优惠券折扣金额</summary>
    private static decimal CalculateCouponDiscount(string couponType, decimal discountValue, decimal cartTotal, decimal maxDiscount)
    {
        return couponType?.ToLower() switch
        {
            "fixed" => discountValue,
            "percentage" => Math.Min(
                cartTotal * discountValue / 100m,
                maxDiscount > 0 ? maxDiscount : cartTotal),
            "free_shipping" => 0,
            "gift" => 0,
            _ => 0
        };
    }
}
