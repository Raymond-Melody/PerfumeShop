using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Ipblacklist
{
    public int Ipid { get; set; }

    public string Ipaddress { get; set; } = null!;

    public string? Reason { get; set; }

    public DateTime? BlockedAt { get; set; }

    public int? BlockedBy { get; set; }

    public bool? IsActive { get; set; }

    public DateTime? ExpiresAt { get; set; }

    public int? HitCount { get; set; }

    public DateTime? LastHitAt { get; set; }
}
