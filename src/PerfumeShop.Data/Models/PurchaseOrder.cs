using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseOrder
{
    public DateTime? ApprovedAt { get; set; }

    public int? ApprovedBy { get; set; }

    public string? CategoryCode { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int? CreatedBy { get; set; }

    public DateTime? ExpectedDate { get; set; }

    public DateTime? OrderDate { get; set; }

    public int PurchaseId { get; set; }

    public string? PurchaseNo { get; set; }

    public string? Remarks { get; set; }

    public string? Status { get; set; }

    public int? SupplierId { get; set; }

    public decimal? TotalAmount { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public string? OrderType { get; set; }

    public DateTime? ExpectedDeliveryDate { get; set; }
}
