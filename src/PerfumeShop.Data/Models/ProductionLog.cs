using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductionLog
{
    public DateTime? CreatedAt { get; set; }

    public string? CreatedBy { get; set; }

    public int LogId { get; set; }

    public string? Notes { get; set; }

    public int ProductionId { get; set; }

    public string? Status { get; set; }
}
