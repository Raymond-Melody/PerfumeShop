using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Product
{
    public string? BaseIngredients { get; set; }

    public decimal BasePrice { get; set; }

    public decimal? Bomcost { get; set; }

    public string? Category { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public bool? Engravable { get; set; }

    public decimal? EngravingPrice { get; set; }

    public string? ImageUrl { get; set; }

    public bool? IsActive { get; set; }

    public int? Kolid { get; set; }

    public int ProductId { get; set; }

    public string ProductName { get; set; } = null!;

    public string? ProductType { get; set; }

    public int? RecipeId { get; set; }

    public string? ReviewStatus { get; set; }

    public decimal? UnitCost { get; set; }
}
