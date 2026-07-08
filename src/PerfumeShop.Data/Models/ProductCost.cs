using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductCost
{
    public int CostId { get; set; }

    public string? CostName { get; set; }

    public string? CostType { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? CreatedBy { get; set; }

    public DateTime? EffectiveDate { get; set; }

    public DateTime? ExpiryDate { get; set; }

    public int ProductId { get; set; }

    public double? Quantity { get; set; }

    public decimal? TotalCost { get; set; }

    public decimal? UnitCost { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
