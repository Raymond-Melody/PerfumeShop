using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseReceipt
{
    public DateTime? CreatedAt { get; set; }

    public string? Notes { get; set; }

    public int? PurchaseId { get; set; }

    public DateTime? ReceiptDate { get; set; }

    public int ReceiptId { get; set; }

    public string? ReceiptNo { get; set; }

    public string? ReceivedBy { get; set; }

    public string? Status { get; set; }

    public int? SupplierId { get; set; }

    public double? TotalReceivedQty { get; set; }
}
