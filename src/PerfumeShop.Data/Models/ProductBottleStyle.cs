using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductBottleStyle
{
    public int BottleId { get; set; }

    public decimal? CustomPrice { get; set; }

    public int Id { get; set; }

    public int ProductId { get; set; }
}
