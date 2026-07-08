using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ReferralToken
{
    public int TokenId { get; set; }

    public int ReferrerUserId { get; set; }

    public string? ReferrerType { get; set; }

    public string TokenHash { get; set; } = null!;

    public DateTime ExpiresAt { get; set; }

    public int? MaxUses { get; set; }

    public int? UsedCount { get; set; }

    public bool? IsActive { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? OriginalToken { get; set; }
}
