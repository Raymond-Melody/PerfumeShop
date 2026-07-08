using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseReceiptDetail
{
    public double? AcceptedQty { get; set; }

    public int? MaterialId { get; set; }

    public int? PurchaseDetailId { get; set; }

    public int ReceiptDetailId { get; set; }

    public int? ReceiptId { get; set; }

    public double? ReceivedQty { get; set; }

    public double? RejectedQty { get; set; }

    public string? RejectReason { get; set; }

    public decimal? UnitPrice { get; set; }
}
