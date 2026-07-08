using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AccountsPayable
{
    public int PayableId { get; set; }

    public string? PayableNo { get; set; }

    public string? SupplierName { get; set; }

    public decimal? Amount { get; set; }

    public decimal? PaidAmount { get; set; }

    public string? Status { get; set; }

    public DateTime? DueDate { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
