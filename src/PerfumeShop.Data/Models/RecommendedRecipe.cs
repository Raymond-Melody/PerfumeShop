using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecommendedRecipe
{
    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public bool? IsActive { get; set; }

    public int? ProductId { get; set; }

    public int RecipeId { get; set; }

    public string? RecipeName { get; set; }

    public int? SortOrder { get; set; }
}
