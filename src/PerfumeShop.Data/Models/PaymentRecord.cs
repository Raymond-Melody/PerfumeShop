using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PaymentRecord
{
    public decimal? Amount { get; set; }

    public string? Category { get; set; }

    public DateTime? CreatedAt { get; set; }

    public decimal? Fee { get; set; }

    public decimal? NetAmount { get; set; }

    public int? OrderId { get; set; }

    public string? OrderNo { get; set; }

    public string? PaymentMethod { get; set; }

    public string? ReconcileStatus { get; set; }

    public int RecordId { get; set; }

    public string? Remark { get; set; }

    public string? Status { get; set; }

    public string? TransactionNo { get; set; }

    public string? TransactionType { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public string? PaymentType { get; set; }

    public string? VoucherNo { get; set; }

    public int? CenterId { get; set; }

    public int? PayableId { get; set; }

    public int? ReceivableId { get; set; }
}
