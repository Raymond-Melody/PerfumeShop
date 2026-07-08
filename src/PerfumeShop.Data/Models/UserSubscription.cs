using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class UserSubscription
{
    public int SubscriptionId { get; set; }

    public int UserId { get; set; }

    public int PlanId { get; set; }

    public string Status { get; set; } = null!;

    public DateTime StartDate { get; set; }

    public DateTime? EndDate { get; set; }

    public bool AutoRenew { get; set; }

    public DateTime CreatedAt { get; set; }
}
