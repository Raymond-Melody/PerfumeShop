using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductVolumePrice
{
    public decimal Price { get; set; }

    public int ProductId { get; set; }

    public int PvpriceId { get; set; }

    public int VolumeId { get; set; }
}
