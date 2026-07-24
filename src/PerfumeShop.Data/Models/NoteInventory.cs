using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class NoteInventory
{
    public int InventoryId { get; set; }

    public DateTime? LastRestockDate { get; set; }

    public int? MinStockLevel { get; set; }

    public int NoteId { get; set; }

    public int? StockQuantity { get; set; }

    public DateTime? UpdatedAt { get; set; }

    // V21: 香调加权单位成本（Accord 生产完成时结转）
    public decimal? WeightedUnitCost { get; set; }
}
