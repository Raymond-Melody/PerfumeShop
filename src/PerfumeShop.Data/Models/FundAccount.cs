using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FundAccount
{
    public int AccountId { get; set; }

    public string? AccountName { get; set; }

    public string? AccountType { get; set; }

    public decimal? AlertThreshold { get; set; }

    public decimal? AvailableBalance { get; set; }

    public DateTime? CreatedAt { get; set; }

    public decimal? FrozenAmount { get; set; }

    public bool? IsActive { get; set; }

    public DateTime? LastSyncAt { get; set; }

    public decimal? PendingSettlement { get; set; }

    public decimal? TotalBalance { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
