using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseOrderStatusLog
{
    public int LogId { get; set; }

    public int PurchaseId { get; set; }

    public string? FromStatus { get; set; }

    public string ToStatus { get; set; } = null!;

    public string? ChangedBy { get; set; }

    public DateTime? ChangedAt { get; set; }

    public string? Remarks { get; set; }
}
