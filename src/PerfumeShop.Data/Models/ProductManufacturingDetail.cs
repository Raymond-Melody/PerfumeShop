using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductManufacturingDetail
{
    public double? ActualQty { get; set; }

    public int DetailId { get; set; }

    public int? ManufacturingId { get; set; }

    public int? NoteId { get; set; }

    public string? NoteName { get; set; }

    public double? PlannedQty { get; set; }

    public decimal? TotalCost { get; set; }

    public decimal? UnitCost { get; set; }
}
