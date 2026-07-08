using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FixedBrandPurchaseDetail
{
    public int DetailId { get; set; }

    public int? PurchaseId { get; set; }

    public int? FixedProductId { get; set; }

    public string? ProductName { get; set; }

    public string? Specification { get; set; }

    public int? Quantity { get; set; }

    public int? ReceivedQty { get; set; }

    public decimal? UnitPrice { get; set; }

    public decimal? SubTotal { get; set; }

    public DateTime? ExpectedDate { get; set; }

    public string? Remarks { get; set; }
}
