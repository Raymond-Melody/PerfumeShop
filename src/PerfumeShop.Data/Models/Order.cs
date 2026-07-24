using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Order
{
    public string? ChannelSource { get; set; }

    public decimal? CostAmount { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? DeliveredAt { get; set; }

    public string? Notes { get; set; }

    public int OrderId { get; set; }

    public string OrderNo { get; set; } = null!;

    public string? PaymentMethod { get; set; }

    public decimal? ProfitAmount { get; set; }

    public decimal? RefundAmount { get; set; }

    public DateTime? ShippedAt { get; set; }

    public string? ShippingAddress { get; set; }

    public string? ShippingCity { get; set; }

    public string? ShippingCompany { get; set; }

    public decimal? ShippingFee { get; set; }

    public string? ShippingName { get; set; }

    public string? ShippingNotes { get; set; }

    public string? ShippingPhone { get; set; }

    public string? ShippingPostalCode { get; set; }

    public string? ShippingStatus { get; set; }

    public string? Status { get; set; }

    public decimal TotalAmount { get; set; }

    public string? TrackingNumber { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public int UserId { get; set; }

    public int? PointsEarned { get; set; }

    public int? PointsRedeemed { get; set; }

    public decimal? PointsDiscount { get; set; }

    public string? CouponCode { get; set; }

    public decimal? CouponDiscount { get; set; }

    // V21: 费用分摊金额（运费/平台/推广分摊后回写，纳入利润计算）
    public decimal? ExpenseAmount { get; set; }
}
