using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class SubscriptionPlan
{
    public int PlanId { get; set; }

    public string PlanName { get; set; } = null!;

    public string Period { get; set; } = null!;

    public decimal Price { get; set; }

    public int SampleCount { get; set; }

    public int FullSizeCount { get; set; }

    public bool FreeShipping { get; set; }

    public decimal CancellationFee { get; set; }

    public bool IsActive { get; set; }

    public int SortOrder { get; set; }

    public string? Description { get; set; }

    public string? FeaturedImage { get; set; }

    public DateTime CreatedAt { get; set; }
}
