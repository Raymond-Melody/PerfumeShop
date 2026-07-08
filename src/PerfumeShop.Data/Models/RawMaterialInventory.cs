using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RawMaterialInventory
{
    public string? CategoryCode { get; set; }

    public string? ItemCode { get; set; }

    public string? ItemName { get; set; }

    public DateTime? LastPurchaseDate { get; set; }

    public int MaterialId { get; set; }

    public double? SafetyStock { get; set; }

    public double? StockQty { get; set; }

    public int? SupplierId { get; set; }

    public string? Unit { get; set; }

    public decimal? UnitPrice { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public decimal? AvgDailyUsage { get; set; }

    public int? LeadTimeDays { get; set; }

    public DateTime? LastReplenishDate { get; set; }

    public decimal? ReorderPoint { get; set; }

    public decimal? WeightedUnitCost { get; set; }
}
