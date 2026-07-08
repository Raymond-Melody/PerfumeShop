using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class GroupBuyOrder
{
    public int GroupId { get; set; }

    public int PlanId { get; set; }

    public string GroupSn { get; set; } = null!;

    public int InitiatorId { get; set; }

    public int Status { get; set; }

    public int CurrentSize { get; set; }

    public int TargetSize { get; set; }

    public DateTime CreatedAt { get; set; }

    public DateTime? CompletedAt { get; set; }
}
