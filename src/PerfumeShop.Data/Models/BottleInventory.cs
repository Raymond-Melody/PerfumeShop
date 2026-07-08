using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class BottleInventory
{
    public int BottleId { get; set; }

    public string? BottleName { get; set; }

    public decimal? StockQty { get; set; }

    public decimal? SafetyStock { get; set; }

    public decimal? UnitCost { get; set; }

    public bool? IsActive { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
