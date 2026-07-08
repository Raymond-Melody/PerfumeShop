using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AccordProductionDetail
{
    public double? ActualQty { get; set; }

    public int DetailId { get; set; }

    public int? MaterialId { get; set; }

    public string? MaterialName { get; set; }

    public double? PlannedQty { get; set; }

    public int? ProductionId { get; set; }

    public decimal? TotalCost { get; set; }

    public decimal? UnitCost { get; set; }
}
