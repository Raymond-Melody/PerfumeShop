using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class SupplierContract
{
    public int ContractId { get; set; }

    public int SupplierId { get; set; }

    public string? ContractNo { get; set; }

    public string? ContractName { get; set; }

    public string? ContractType { get; set; }

    public DateTime? StartDate { get; set; }

    public DateTime? EndDate { get; set; }

    public decimal? TotalAmount { get; set; }

    public string? PaymentTerms { get; set; }

    public string? TermsSummary { get; set; }

    public string? AttachmentUrl { get; set; }

    public string? Status { get; set; }

    public DateTime? SignedAt { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
