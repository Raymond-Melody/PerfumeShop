using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Ingredient
{
    public string? Casnumber { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public int IngredientId { get; set; }

    public string IngredientName { get; set; } = null!;

    public bool? IsActive { get; set; }
}
