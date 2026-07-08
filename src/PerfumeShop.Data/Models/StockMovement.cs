using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class StockMovement
{
    public int MovementId { get; set; }

    public string? ItemType { get; set; }

    public int? ItemId { get; set; }

    public string? ItemName { get; set; }

    public string? ItemCode { get; set; }

    public string? MovementType { get; set; }

    public decimal? Quantity { get; set; }

    public decimal? BeforeQty { get; set; }

    public decimal? AfterQty { get; set; }

    public string? Unit { get; set; }

    public string? ReferenceNo { get; set; }

    public string? Notes { get; set; }

    public string? CreatedBy { get; set; }

    public DateTime? CreatedAt { get; set; }
}
