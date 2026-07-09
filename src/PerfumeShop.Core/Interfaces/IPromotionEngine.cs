namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 促销引擎接口 — 优惠券验证、折扣计算、优惠券发放
/// 对应 ASP 中的 promotion_engine.asp
/// </summary>
public interface IPromotionEngine
{
    // ========== 优惠券验证 ==========

    /// <summary>验证优惠券是否可用于当前购物车</summary>
    Task<CouponValidateResult> ValidateCouponAsync(string code, int userId, decimal cartTotal, CancellationToken ct = default);

    /// <summary>计算折扣金额 (满减/折扣/会员等级/首单)</summary>
    Task<decimal> CalculateDiscountAsync(decimal cartTotal, int userId, CancellationToken ct = default);

    /// <summary>检查是否免运费</summary>
    bool CheckFreeShipping(decimal cartTotal, decimal freeShippingThreshold = 299);

    // ========== 优惠券使用 ==========

    /// <summary>使用优惠券 (订单创建成功后调用)</summary>
    Task<bool> UseCouponAsync(string code, int userId, int orderId, CancellationToken ct = default);

    // ========== 优惠券发放 ==========

    /// <summary>给用户发放优惠券</summary>
    Task<bool> IssueCouponAsync(int userId, string code, string source, CancellationToken ct = default);

    /// <summary>新人礼包发放</summary>
    Task<int> IssueWelcomeCouponsAsync(int userId, CancellationToken ct = default);

    /// <summary>会员升级礼券</summary>
    Task<bool> IssueTierUpgradeCouponAsync(int userId, string tierCode, CancellationToken ct = default);

    // ========== 查询 ==========

    /// <summary>获取用户可用优惠券数量</summary>
    Task<int> GetUserCouponCountAsync(int userId, CancellationToken ct = default);
}

/// <summary>优惠券验证结果</summary>
public class CouponValidateResult
{
    public bool Valid { get; set; }
    public string Message { get; set; } = "";
    public decimal Discount { get; set; }
    public string Type { get; set; } = "";
}
