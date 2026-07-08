using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class CostCenter
{
    public int CenterId { get; set; }

    public string CenterCode { get; set; } = null!;

    public string CenterName { get; set; } = null!;

    public string? CenterType { get; set; }

    public int? ParentId { get; set; }

    public decimal? BudgetAmount { get; set; }

    public bool? IsActive { get; set; }

    public string? Notes { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
