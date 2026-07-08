using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PriceChangeLog
{
    public int LogId { get; set; }

    public int? PriceId { get; set; }

    public string? FieldChanged { get; set; }

    public string? OldValue { get; set; }

    public string? NewValue { get; set; }

    public string? ChangedBy { get; set; }

    public DateTime? ChangedAt { get; set; }
}
