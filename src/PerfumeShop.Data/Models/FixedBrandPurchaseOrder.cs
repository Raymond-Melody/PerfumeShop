using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FixedBrandPurchaseOrder
{
    public int PurchaseId { get; set; }

    public string? PurchaseNo { get; set; }

    public int? SupplierId { get; set; }

    public string? SupplierName { get; set; }

    public decimal? TotalAmount { get; set; }

    public string? Status { get; set; }

    public DateTime? OrderDate { get; set; }

    public DateTime? ExpectedDate { get; set; }

    public string? ApprovedBy { get; set; }

    public DateTime? ApprovedAt { get; set; }

    public string? Remarks { get; set; }

    public string? CreatedBy { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
