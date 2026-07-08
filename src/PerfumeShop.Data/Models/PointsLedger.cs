using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PointsLedger
{
    public int LedgerId { get; set; }

    public int UserId { get; set; }

    public int Points { get; set; }

    public string PointType { get; set; } = null!;

    public string Source { get; set; } = null!;

    public int? ReferenceId { get; set; }

    public string? Description { get; set; }

    public DateTime? ExpiresAt { get; set; }

    public bool IsExpired { get; set; }

    public DateTime CreatedAt { get; set; }
}
