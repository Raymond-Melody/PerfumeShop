using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class GroupBuyPlan
{
    public int PlanId { get; set; }

    public int ProductId { get; set; }

    public int TeamSize { get; set; }

    public decimal GroupPrice { get; set; }

    public int MinUnit { get; set; }

    public int MaxUnit { get; set; }

    public DateTime StartTime { get; set; }

    public DateTime EndTime { get; set; }

    public int DurationHours { get; set; }

    public bool IsActive { get; set; }

    public int SortOrder { get; set; }

    public DateTime CreatedAt { get; set; }
}
