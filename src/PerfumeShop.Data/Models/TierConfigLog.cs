using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class TierConfigLog
{
    public int LogId { get; set; }

    public string TierCode { get; set; } = null!;

    public string FieldName { get; set; } = null!;

    public string? OldValue { get; set; }

    public string NewValue { get; set; } = null!;

    public string ChangedBy { get; set; } = null!;

    public DateTime ChangedAt { get; set; }
}
