using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RefundRecord
{
    public DateTime? ApprovedAt { get; set; }

    public string? ApprovedBy { get; set; }

    public DateTime? CompletedAt { get; set; }

    public bool? CostWriteBack { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int OrderId { get; set; }

    public string? OrderNo { get; set; }

    public decimal RefundAmount { get; set; }

    public int RefundId { get; set; }

    public string? RefundNo { get; set; }

    public string? RefundReason { get; set; }

    public string? Status { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
