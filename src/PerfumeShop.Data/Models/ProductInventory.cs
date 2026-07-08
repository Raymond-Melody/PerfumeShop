using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductInventory
{
    public int InventoryId { get; set; }

    public int? NoteId { get; set; }

    public int? ProductId { get; set; }

    public int? SafetyStock { get; set; }

    public int? StockQty { get; set; }

    public string? StockType { get; set; }

    public decimal? UnitCost { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
