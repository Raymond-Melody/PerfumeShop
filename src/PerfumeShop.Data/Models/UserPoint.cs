using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class UserPoint
{
    public int? AvailablePoints { get; set; }

    public int? ExpiredPoints { get; set; }

    public DateTime? LastUpdatedAt { get; set; }

    public int PointId { get; set; }

    public int? TotalPoints { get; set; }

    public int? UsedPoints { get; set; }

    public int UserId { get; set; }
}
