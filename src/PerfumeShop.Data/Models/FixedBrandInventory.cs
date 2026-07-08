using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FixedBrandInventory
{
    public int InventoryId { get; set; }

    public int? FixedProductId { get; set; }

    public string? ProductCode { get; set; }

    public string? ProductName { get; set; }

    public string? Specification { get; set; }

    public int? StockQty { get; set; }

    public int? SafetyStock { get; set; }

    public int? MinOrderQty { get; set; }

    public decimal? AvgUnitCost { get; set; }

    public decimal? LastPurchasePrice { get; set; }

    public DateTime? LastPurchaseDate { get; set; }

    public int? LastPurchaseId { get; set; }

    public int? TotalPurchased { get; set; }

    public int? TotalSold { get; set; }

    public string? ParamMode { get; set; }

    public decimal? DailySalesAvg { get; set; }

    public int? ConsecutiveDataMonths { get; set; }

    public DateTime? LastAutoCalcDate { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
