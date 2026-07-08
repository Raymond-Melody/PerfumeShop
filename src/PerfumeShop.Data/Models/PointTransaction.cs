using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PointTransaction
{
    public DateTime? CreatedAt { get; set; }

    public string? CreatedBy { get; set; }

    public string? Description { get; set; }

    public int? OrderId { get; set; }

    public int Points { get; set; }

    public int? PointsChange { get; set; }

    public string? Reason { get; set; }

    public int TransactionId { get; set; }

    public string? TransactionType { get; set; }

    public int UserId { get; set; }
}
