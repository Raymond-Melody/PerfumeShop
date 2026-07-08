using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseOrderDetail
{
    public int DetailId { get; set; }

    public string? ItemCode { get; set; }

    public string? ItemName { get; set; }

    public int? PurchaseId { get; set; }

    public double? Quantity { get; set; }

    public double? ReceivedQty { get; set; }

    public string? Remarks { get; set; }

    public string? Specification { get; set; }

    public decimal? TotalPrice { get; set; }

    public string? Unit { get; set; }

    public decimal? UnitPrice { get; set; }
}
