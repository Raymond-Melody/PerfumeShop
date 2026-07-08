using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class MaterialOutboundDetail
{
    public double? ActualQty { get; set; }

    public int? MaterialId { get; set; }

    public int OutboundDetailId { get; set; }

    public int? OutboundId { get; set; }

    public int? ProductionOrderRef { get; set; }

    public double? RequestedQty { get; set; }

    public decimal? TotalAmount { get; set; }

    public decimal? UnitPrice { get; set; }
}
