using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PointsRule
{
    public int RuleId { get; set; }

    public string RuleCode { get; set; } = null!;

    public string RuleName { get; set; } = null!;

    public decimal RuleValue { get; set; }

    public string RuleUnit { get; set; } = null!;

    public bool IsEnabled { get; set; }

    public int SortOrder { get; set; }

    public string? Description { get; set; }

    public DateTime CreatedAt { get; set; }

    public DateTime UpdatedAt { get; set; }
}
