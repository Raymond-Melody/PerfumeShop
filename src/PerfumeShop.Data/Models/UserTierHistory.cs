using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class UserTierHistory
{
    public int HistoryId { get; set; }

    public int UserId { get; set; }

    public string? OldTierCode { get; set; }

    public string NewTierCode { get; set; } = null!;

    public decimal TotalSpent { get; set; }

    public string ChangeType { get; set; } = null!;

    public DateTime ChangedAt { get; set; }
}
