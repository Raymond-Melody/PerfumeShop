using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AfterSale
{
    public int AfterSalesId { get; set; }

    public int OrderId { get; set; }

    public int UserId { get; set; }

    public string RequestType { get; set; } = null!;

    public string? Reason { get; set; }

    public string? Status { get; set; }

    public decimal? RefundAmount { get; set; }

    public string? AdminNotes { get; set; }

    public int? ProcessedBy { get; set; }

    public DateTime? ProcessedAt { get; set; }

    public DateTime? CreatedAt { get; set; }
}
