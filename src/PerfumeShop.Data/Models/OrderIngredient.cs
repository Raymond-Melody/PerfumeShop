using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class OrderIngredient
{
    public DateTime? CreatedAt { get; set; }

    public int? DetailId { get; set; }

    public int IngredientId { get; set; }

    public string IngredientName { get; set; } = null!;

    public int OrderId { get; set; }
}
