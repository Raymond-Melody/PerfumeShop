using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ReconciliationLog
{
    public DateTime? CreatedAt { get; set; }

    public decimal? Difference { get; set; }

    public int LogId { get; set; }

    public decimal? OrderAmount { get; set; }

    public int? OrderId { get; set; }

    public string? OrderNo { get; set; }

    public decimal? PaymentAmount { get; set; }

    public DateTime? ReconcileDate { get; set; }

    public string? Resolution { get; set; }

    public DateTime? ResolvedAt { get; set; }

    public string? ResolvedBy { get; set; }

    public string? Status { get; set; }
}
