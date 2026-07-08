using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FixedBrandReceipt
{
    public int ReceiptId { get; set; }

    public int? PurchaseId { get; set; }

    public string? ReceiptNo { get; set; }

    public int? SupplierId { get; set; }

    public string? ReceivedBy { get; set; }

    public DateTime? ReceiptDate { get; set; }

    public int? TotalReceivedQty { get; set; }

    public string? Notes { get; set; }

    public DateTime? CreatedAt { get; set; }
}
