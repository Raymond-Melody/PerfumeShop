using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PackagingInventory
{
    public int PackagingId { get; set; }

    public string? ItemName { get; set; }

    public string? ItemCode { get; set; }

    public decimal? StockQty { get; set; }

    public decimal? SafetyStock { get; set; }

    public string? Unit { get; set; }

    public decimal? UnitPrice { get; set; }

    public decimal? AvgDailyUsage { get; set; }

    public int? LeadTimeDays { get; set; }

    public DateTime? LastReplenishDate { get; set; }

    public decimal? ReorderPoint { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
