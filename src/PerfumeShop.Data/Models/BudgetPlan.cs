using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class BudgetPlan
{
    public decimal? ActualAmount { get; set; }

    public double? AlertPercent { get; set; }

    public double? AlertRoi { get; set; }

    public decimal? BudgetAmount { get; set; }

    public int BudgetId { get; set; }

    public string? BudgetName { get; set; }

    public string? Category { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? CreatedBy { get; set; }

    public decimal? Gmvamount { get; set; }

    public string? Period { get; set; }

    public double? Roi { get; set; }

    public string? Status { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
