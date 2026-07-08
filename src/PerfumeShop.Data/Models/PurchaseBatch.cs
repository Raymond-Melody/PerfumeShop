using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseBatch
{
    public int BatchId { get; set; }

    public int? PurchaseDetailId { get; set; }

    public int? PurchaseId { get; set; }

    public string? BatchNo { get; set; }

    public string? ItemType { get; set; }

    public string? ItemCode { get; set; }

    public string? ItemName { get; set; }

    public decimal? UnitPrice { get; set; }

    public double? Quantity { get; set; }

    public double? ReceivedQty { get; set; }

    public double? RemainingQty { get; set; }

    public DateTime? ReceivedDate { get; set; }

    public int? SupplierId { get; set; }

    public bool? CostAllocated { get; set; }

    public DateTime? CreatedAt { get; set; }
}
