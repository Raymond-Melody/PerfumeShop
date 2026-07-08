using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FixedBrandReceiptDetail
{
    public int ReceiptDetailId { get; set; }

    public int? ReceiptId { get; set; }

    public int? DetailId { get; set; }

    public int? FixedProductId { get; set; }

    public int? AcceptedQty { get; set; }

    public int? RejectedQty { get; set; }

    public string? RejectReason { get; set; }

    public decimal? UnitPrice { get; set; }
}
