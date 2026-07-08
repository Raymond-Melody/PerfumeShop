using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class SupplierEvaluation
{
    public int EvaluationId { get; set; }

    public int SupplierId { get; set; }

    public string? EvaluatedBy { get; set; }

    public DateTime? EvaluationDate { get; set; }

    public int? QualityScore { get; set; }

    public int? DeliveryScore { get; set; }

    public int? PriceScore { get; set; }

    public int? ServiceScore { get; set; }

    public int? OverallScore { get; set; }

    public string? Rating { get; set; }

    public string? Comments { get; set; }

    public string? Recommendations { get; set; }

    public string? Period { get; set; }

    public DateTime? CreatedAt { get; set; }
}
