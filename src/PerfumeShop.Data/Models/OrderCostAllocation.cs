using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class OrderCostAllocation
{
    public int AllocationId { get; set; }

    public int? OrderId { get; set; }

    public string? OrderNo { get; set; }

    public string? CostType { get; set; }

    public string? ItemCode { get; set; }

    public string? ItemName { get; set; }

    public int? BatchId { get; set; }

    public int? InvBatchId { get; set; }

    public decimal? UnitCost { get; set; }

    public double? Quantity { get; set; }

    public decimal? TotalCost { get; set; }

    public DateTime? AllocatedAt { get; set; }

    public DateTime? CreatedAt { get; set; }
}
