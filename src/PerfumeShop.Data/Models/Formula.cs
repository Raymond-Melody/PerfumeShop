using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Formula
{
    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public int FormulaId { get; set; }

    public string FormulaName { get; set; } = null!;

    public bool? IsActive { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
