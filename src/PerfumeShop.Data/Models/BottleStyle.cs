using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class BottleStyle
{
    public int BottleId { get; set; }

    public string BottleName { get; set; } = null!;

    public string? Description { get; set; }

    public string? ImageUrl { get; set; }

    public bool? IsActive { get; set; }

    public decimal? PriceAddition { get; set; }

    public decimal? StockQty { get; set; }

    public decimal? SafetyStock { get; set; }

    public decimal? AvgDailyUsage { get; set; }

    public int? LeadTimeDays { get; set; }

    public DateTime? LastReplenishDate { get; set; }

    public decimal? ReorderPoint { get; set; }

    public decimal? UnitPrice { get; set; }

    public decimal? UnitCost { get; set; }

    public string? BottleType { get; set; }

    public int? CapacityMl { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public DateTime? LastPurchaseDate { get; set; }
}
