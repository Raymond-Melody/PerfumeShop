using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AppLog
{
    public long LogId { get; set; }

    public string LogLevel { get; set; } = null!;

    public string? LogMessage { get; set; }

    public string? LogSource { get; set; }

    public int? LineNumber { get; set; }

    public string? UserName { get; set; }

    public string? Ipaddress { get; set; }

    public string? PageUrl { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? LogType { get; set; }
}
