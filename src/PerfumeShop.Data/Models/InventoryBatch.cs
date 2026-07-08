using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class InventoryBatch
{
    public int BatchId { get; set; }

    public string? ItemType { get; set; }

    public int? ItemId { get; set; }

    public string? ItemCode { get; set; }

    public string? ItemName { get; set; }

    public string? BatchNo { get; set; }

    public decimal? UnitCost { get; set; }

    public double? StockQty { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public DateTime? CreatedAt { get; set; }
}
