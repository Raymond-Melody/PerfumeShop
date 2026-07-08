using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipeProduct
{
    public double? BatchSize { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int? ProductId { get; set; }

    public int ProductRecipeId { get; set; }

    public DateTime? PublishedAt { get; set; }

    public string? PublishedBy { get; set; }

    public int? RecipeId { get; set; }

    public string? Status { get; set; }
}
