using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipeAccord
{
    public int AccordRecipeId { get; set; }

    public double? BatchSize { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int? NoteId { get; set; }

    public DateTime? PublishedAt { get; set; }

    public string? PublishedBy { get; set; }

    public int? RecipeId { get; set; }

    public string? RecipeName { get; set; }

    public string? Status { get; set; }
}
