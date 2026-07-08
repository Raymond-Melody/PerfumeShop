using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseHistoryStat
{
    public int StatId { get; set; }

    public string ItemType { get; set; } = null!;

    public string? ItemCode { get; set; }

    public string? ItemName { get; set; }

    public decimal? Avg30DayUsage { get; set; }

    public decimal? Avg90DayUsage { get; set; }

    public DateTime? LastOrderDate { get; set; }

    public int? TotalOrders90Days { get; set; }

    public int? PreferredSupplierId { get; set; }

    public decimal? PreferredUnitPrice { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
