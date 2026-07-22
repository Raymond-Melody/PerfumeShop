namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 促销引擎接口 — 优惠券验证、折扣计算、优惠券发放、促销资格检查
/// 对应 ASP 中的 promotion_engine.asp
/// </summary>
public interface IPromotionEngine
{
    // ========== 促销资格检查 (V18: PE_CheckPromotionEligibility) ==========

    /// <summary>检查购物车是否满足促销条件，返回可用促销列表</summary>
    Task<PromotionEligibilityResult> CheckPromotionEligibilityAsync(decimal cartTotal, int userId, CancellationToken ct = default);

    // ========== 优惠券验证 (V18: PE_CouponValidate) ==========

    /// <summary>验证优惠券是否可用于当前购物车（完整校验链：首单限制、库存、有效期、品类限制、用户资格）</summary>
    Task<CouponValidateResult> ValidateCouponAsync(string code, int userId, decimal cartTotal, string? cartCategory = null, CancellationToken ct = default);

    // ========== 折扣计算 (V18: PE_CalculateDiscount) ==========

    /// <summary>计算折扣金额 (满减/折扣/会员等级/首单)</summary>
    Task<decimal> CalculateDiscountAsync(decimal cartTotal, int userId, CancellationToken ct = default);

    /// <summary>应用折扣到购物车 — 综合促销+优惠券，返回最终折扣</summary>
    Task<ApplyDiscountResult> ApplyDiscountAsync(decimal cartTotal, int userId, string? couponCode, string? cartCategory = null, CancellationToken ct = default);

    /// <summary>检查是否免运费 (V18: PE_CheckFreeShipping)</summary>
    bool CheckFreeShipping(decimal cartTotal, decimal freeShippingThreshold = 299);

    // ========== 优惠券使用 (V18: PE_CouponUse) ==========

    /// <summary>使用优惠券 (订单创建成功后调用)</summary>
    Task<bool> UseCouponAsync(string code, int userId, int orderId, CancellationToken ct = default);

    // ========== 优惠券发放 ==========

    /// <summary>给用户发放优惠券 (V18: PE_CouponIssue)</summary>
    Task<bool> IssueCouponAsync(int userId, string code, string source, CancellationToken ct = default);

    /// <summary>新人礼包发放 (V18: PE_CouponIssueWelcome)</summary>
    Task<int> IssueWelcomeCouponsAsync(int userId, CancellationToken ct = default);

    /// <summary>会员升级礼券 (V18: PE_CouponIssueTierUpgrade)</summary>
    Task<bool> IssueTierUpgradeCouponAsync(int userId, string tierCode, CancellationToken ct = default);

    // ========== 查询 ==========

    /// <summary>获取用户可用优惠券数量 (V18: PE_CouponGetUserCount)</summary>
    Task<int> GetUserCouponCountAsync(int userId, CancellationToken ct = default);

    /// <summary>获取用户促销使用历史</summary>
    Task<IEnumerable<PromotionHistoryRecord>> GetPromotionHistoryByUserAsync(int userId, CancellationToken ct = default);

    /// <summary>获取用户可用优惠券列表 (V18: PE_CouponGetUserCoupons)</summary>
    Task<IEnumerable<UserCouponDto>> GetUserCouponsAsync(int userId, string statusFilter = "available", CancellationToken ct = default);

    /// <summary>获取购物车可用的优惠券 (V18: PE_CouponGetApplicable)</summary>
    Task<IEnumerable<UserCouponDto>> GetApplicableCouponsAsync(int userId, decimal cartTotal, CancellationToken ct = default);

    /// <summary>获取优惠券使用统计 (V18: PE_CouponGetStats)</summary>
    Task<CouponStatsDto> GetCouponStatsAsync(int couponId, CancellationToken ct = default);
}

/// <summary>优惠券验证结果</summary>
public class CouponValidateResult
{
    public bool Valid { get; set; }
    public string Message { get; set; } = "";
    public decimal Discount { get; set; }
    public string Type { get; set; } = "";
}

/// <summary>促销资格检查结果</summary>
public class PromotionEligibilityResult
{
    public bool HasThreshold { get; set; }
    public string ThresholdName { get; set; } = "";
    public decimal ThresholdMinAmount { get; set; }
    public decimal ThresholdDiscount { get; set; }
    public bool IsEligible { get; set; }
    public decimal RemainingToThreshold { get; set; }
    public bool FreeShippingEligible { get; set; }
}

/// <summary>应用折扣结果</summary>
public class ApplyDiscountResult
{
    public decimal PromotionDiscount { get; set; }
    public decimal CouponDiscount { get; set; }
    public decimal MemberDiscount { get; set; }
    public decimal FirstOrderDiscount { get; set; }
    public decimal TotalDiscount { get; set; }
    public decimal FinalAmount { get; set; }
    public bool FreeShipping { get; set; }
    public string? AppliedCouponCode { get; set; }
    public string AppliedCouponType { get; set; } = "";
}

/// <summary>用户促销历史记录</summary>
public class PromotionHistoryRecord
{
    public int UserCouponId { get; set; }
    public string CouponCode { get; set; } = "";
    public string CouponName { get; set; } = "";
    public string CouponType { get; set; } = "";
    public decimal DiscountValue { get; set; }
    public string Source { get; set; } = "";
    public string Status { get; set; } = "";
    public DateTime ObtainedAt { get; set; }
    public DateTime? UsedAt { get; set; }
    public int? UsedOrderId { get; set; }
}

/// <summary>用户优惠券 DTO</summary>
public class UserCouponDto
{
    public int UserCouponId { get; set; }
    public string CouponCode { get; set; } = "";
    public string CouponName { get; set; } = "";
    public string CouponType { get; set; } = "";
    public decimal DiscountValue { get; set; }
    public decimal MinSpend { get; set; }
    public decimal MaxDiscount { get; set; }
    public string? Description { get; set; }
    public string? Terms { get; set; }
    public DateTime ValidFrom { get; set; }
    public DateTime ValidTo { get; set; }
    public string Status { get; set; } = "";
    public DateTime ObtainedAt { get; set; }
}

/// <summary>优惠券统计 DTO</summary>
public class CouponStatsDto
{
    public int TotalIssued { get; set; }
    public int TotalUsed { get; set; }
    public decimal TotalAmount { get; set; }
}
