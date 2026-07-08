using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ExpenseRecord
{
    public string? AllocationMethod { get; set; }

    public double? AllocationRatio { get; set; }

    public decimal? Amount { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int ExpenseId { get; set; }

    public string? ExpenseName { get; set; }

    public string? ExpenseType { get; set; }

    public int? OrderId { get; set; }

    public string? Period { get; set; }

    public int? ProductId { get; set; }

    public int? SourceOrderId { get; set; }

    public int? CenterId { get; set; }
}
