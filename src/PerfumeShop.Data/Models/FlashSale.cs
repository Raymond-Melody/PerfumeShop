using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FlashSale
{
    public int FlashSaleId { get; set; }

    public int ProductId { get; set; }

    public decimal FlashPrice { get; set; }

    public int Stock { get; set; }

    public int SoldCount { get; set; }

    public int LimitPerUser { get; set; }

    public DateTime StartTime { get; set; }

    public DateTime EndTime { get; set; }

    public bool IsActive { get; set; }

    public int SortOrder { get; set; }

    public DateTime CreatedAt { get; set; }
}
