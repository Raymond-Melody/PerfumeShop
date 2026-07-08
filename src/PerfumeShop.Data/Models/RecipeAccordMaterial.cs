using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipeAccordMaterial
{
    public int? AccordRecipeId { get; set; }

    public int DetailId { get; set; }

    public int? MaterialId { get; set; }

    public string? MaterialName { get; set; }

    public string? Notes { get; set; }

    public double? Percentage { get; set; }

    public double? PlannedQty { get; set; }
}
