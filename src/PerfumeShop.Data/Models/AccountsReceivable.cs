using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AccountsReceivable
{
    public int ReceivableId { get; set; }

    public string? ReceivableNo { get; set; }

    public string? CustomerName { get; set; }

    public decimal? Amount { get; set; }

    public decimal? ReceivedAmount { get; set; }

    public string? Status { get; set; }

    public DateTime? DueDate { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
