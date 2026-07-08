using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipePopularity
{
    public int? FavoriteCount { get; set; }

    public DateTime? LastCalculatedAt { get; set; }

    public int PopularityId { get; set; }

    public int ProductId { get; set; }

    public int? PurchaseCount { get; set; }

    public int? ViewCount { get; set; }
}
