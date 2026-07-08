using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductManufacturing
{
    public double? ActualQty { get; set; }

    public string? BatchNo { get; set; }

    public DateTime? CompletedAt { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int ManufacturingId { get; set; }

    public string? Notes { get; set; }

    public double? PlannedQty { get; set; }

    public int? ProductId { get; set; }

    public string? ProductName { get; set; }

    public int? ProductRecipeId { get; set; }

    public DateTime? StartedAt { get; set; }

    public string? Status { get; set; }

    public int? TransferRequestId { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public string? WorkCenter { get; set; }
}
