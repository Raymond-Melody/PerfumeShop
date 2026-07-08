using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AdminAuditLog
{
    public int LogId { get; set; }

    public int AdminId { get; set; }

    public string? AdminName { get; set; }

    public string ActionType { get; set; } = null!;

    public string? TargetType { get; set; }

    public int? TargetId { get; set; }

    public string? TargetName { get; set; }

    public string? Details { get; set; }

    public string? Ipaddress { get; set; }

    public string? UserAgent { get; set; }

    public DateTime? CreatedAt { get; set; }
}
