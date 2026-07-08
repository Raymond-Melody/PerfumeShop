using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class SupplierPrice
{
    public DateTime? CreatedAt { get; set; }

    public DateTime? EffectiveDate { get; set; }

    public DateTime? ExpiryDate { get; set; }

    public bool? IsActive { get; set; }

    public string? ItemCode { get; set; }

    public string? ItemName { get; set; }

    public double? MinOrderQty { get; set; }

    public int PriceId { get; set; }

    public int? SupplierId { get; set; }

    public decimal? UnitPrice { get; set; }

    public string? PriceType { get; set; }

    public string? Unit { get; set; }
}
