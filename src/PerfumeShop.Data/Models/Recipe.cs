using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Recipe
{
    public DateTime? CreatedAt { get; set; }

    public string? CreatedBy { get; set; }

    public string? Description { get; set; }

    public bool? IsActive { get; set; }

    public string? ProductType { get; set; }

    public string? RecipeCode { get; set; }

    public int RecipeId { get; set; }

    public string? RecipeName { get; set; }

    public string? ReviewStatus { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
