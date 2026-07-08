using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PointsRedemption
{
    public int RedemptionId { get; set; }

    public string ItemName { get; set; } = null!;

    public string ItemType { get; set; } = null!;

    public int PointsCost { get; set; }

    public int Stock { get; set; }

    public string? ImageUrl { get; set; }

    public decimal RedemptionValue { get; set; }

    public int MinUserLevel { get; set; }

    public bool IsEnabled { get; set; }

    public int SortOrder { get; set; }

    public string? Description { get; set; }

    public string? Terms { get; set; }

    public DateTime CreatedAt { get; set; }

    public DateTime UpdatedAt { get; set; }
}
