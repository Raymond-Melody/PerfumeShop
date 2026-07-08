using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FixedBrandProduct
{
    public int FixedProductId { get; set; }

    public int? ProductId { get; set; }

    public string? ProductCode { get; set; }

    public string ProductName { get; set; } = null!;

    public string? Specification { get; set; }

    public decimal? UnitPrice { get; set; }

    public decimal? SalePrice { get; set; }

    public int? SupplierId { get; set; }

    public string? SupplierName { get; set; }

    public int? MinOrderQty { get; set; }

    public int? LeadTimeDays { get; set; }

    public string? ImageUrl { get; set; }

    public string? Status { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public int? SafetyStockManual { get; set; }

    public int? LeadTimeDaysManual { get; set; }
}
