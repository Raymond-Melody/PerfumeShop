using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class MemberTier
{
    public int TierId { get; set; }

    public string TierCode { get; set; } = null!;

    public string TierName { get; set; } = null!;

    public string? TierNameEn { get; set; }

    public decimal MinSpent { get; set; }

    public decimal? MaxSpent { get; set; }

    public decimal DiscountRate { get; set; }

    public bool FreeShipping { get; set; }

    public bool PriorityShipping { get; set; }

    public bool BirthdayGift { get; set; }

    public bool DedicatedSupport { get; set; }

    public string? IconClass { get; set; }

    public string? Color { get; set; }

    public string? BadgeBg { get; set; }

    public int SortOrder { get; set; }

    public bool IsActive { get; set; }

    public DateTime CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
