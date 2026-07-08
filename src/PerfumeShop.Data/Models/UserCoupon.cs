using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class UserCoupon
{
    public int UserCouponId { get; set; }

    public int UserId { get; set; }

    public int CouponId { get; set; }

    public string CouponCode { get; set; } = null!;

    public string Source { get; set; } = null!;

    public string Status { get; set; } = null!;

    public DateTime? UsedAt { get; set; }

    public int? UsedOrderId { get; set; }

    public DateTime ObtainedAt { get; set; }

    public DateTime? ExpiresAt { get; set; }

    public virtual Coupon Coupon { get; set; } = null!;

    public virtual User User { get; set; } = null!;
}
