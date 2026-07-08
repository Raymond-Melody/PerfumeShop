using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FixedBrandCostAllocation
{
    public int AllocationId { get; set; }

    public int? OrderId { get; set; }

    public string? OrderNo { get; set; }

    public int? PurchaseId { get; set; }

    public string? PurchaseNo { get; set; }

    public int? FixedProductId { get; set; }

    public string? ProductName { get; set; }

    public decimal? CostPerUnit { get; set; }

    public int? Quantity { get; set; }

    public decimal? TotalCost { get; set; }

    public decimal? SalePrice { get; set; }

    public decimal? ProfitAmount { get; set; }

    public decimal? ProfitRate { get; set; }

    public DateTime? AllocatedAt { get; set; }
}
