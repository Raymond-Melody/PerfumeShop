using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class MemberBenefit
{
    public int BenefitId { get; set; }

    public string TierCode { get; set; } = null!;

    public string BenefitName { get; set; } = null!;

    public string? BenefitDesc { get; set; }

    public string? BenefitIcon { get; set; }

    public int SortOrder { get; set; }

    public bool IsActive { get; set; }
}
