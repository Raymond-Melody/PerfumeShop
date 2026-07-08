using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Coupon
{
    public string? CouponCode { get; set; }

    public int CouponId { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? DiscountType { get; set; }

    public decimal? DiscountValue { get; set; }

    public DateTime? EndDate { get; set; }

    public bool? IsActive { get; set; }

    public decimal? MinPurchase { get; set; }

    public DateTime? StartDate { get; set; }

    public int? UsageLimit { get; set; }

    public int? UsedCount { get; set; }

    public string CouponName { get; set; } = null!;

    public string CouponType { get; set; } = null!;

    public decimal MinSpend { get; set; }

    public decimal MaxDiscount { get; set; }

    public DateTime ValidFrom { get; set; }

    public DateTime ValidTo { get; set; }

    public int TotalQty { get; set; }

    public int UsedQty { get; set; }

    public bool FirstOrderOnly { get; set; }

    public string? ApplicableCategory { get; set; }

    public int? ApplicableProductId { get; set; }

    public string? Description { get; set; }

    public string? Terms { get; set; }

    public bool IsPublic { get; set; }

    public DateTime UpdatedAt { get; set; }

    public virtual ICollection<UserCoupon> UserCoupons { get; set; } = new List<UserCoupon>();
}
