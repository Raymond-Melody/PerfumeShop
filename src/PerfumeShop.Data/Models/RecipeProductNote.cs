using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipeProductNote
{
    public int DetailId { get; set; }

    public int? NoteId { get; set; }

    public string? NoteName { get; set; }

    public string? Notes { get; set; }

    public double? Percentage { get; set; }

    public double? PlannedQty { get; set; }

    public int? ProductRecipeId { get; set; }
}
