using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AdminLog
{
    public string? ActionType { get; set; }

    public int? AdminId { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int LogId { get; set; }

    public string? ModuleCode { get; set; }

    public string? Notes { get; set; }

    public string? RecordId { get; set; }

    public string? TableName { get; set; }

    public string? Ipaddress { get; set; }
}
