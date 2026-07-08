using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseCostReview
{
    public string? CostAllocation { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int? PurchaseId { get; set; }

    public decimal? ReviewAmount { get; set; }

    public string? ReviewComments { get; set; }

    public DateTime? ReviewedAt { get; set; }

    public int? ReviewerId { get; set; }

    public int ReviewId { get; set; }

    public string? ReviewStatus { get; set; }
}
