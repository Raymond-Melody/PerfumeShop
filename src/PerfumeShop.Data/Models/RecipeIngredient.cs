using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipeIngredient
{
    public int Id { get; set; }

    public string? IngredientName { get; set; }

    public int? NoteId { get; set; }

    public double? Percentage { get; set; }

    public int? RecipeId { get; set; }
}
