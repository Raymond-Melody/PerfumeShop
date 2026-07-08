using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductionOrder
{
    public string? AssignedTo { get; set; }

    public int? BottleIndex { get; set; }

    public DateTime? CompletedAt { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int? DetailId { get; set; }

    public DateTime? EstimatedDate { get; set; }

    public string? Notes { get; set; }

    public int OrderId { get; set; }

    public int? Priority { get; set; }

    public string? PriorityText { get; set; }

    public int ProductionId { get; set; }

    public string? Qcnotes { get; set; }

    public DateTime? QcpassedAt { get; set; }

    public int? RecipeId { get; set; }

    public string? RecipeName { get; set; }

    public DateTime? ShippedOutAt { get; set; }

    public DateTime? StartedAt { get; set; }

    public string? Status { get; set; }

    public int? TotalBottles { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public DateTime? WarehouseInAt { get; set; }

    public string? WorkOrderNo { get; set; }

    public int? PlannedQty { get; set; }

    public string? BatchNo { get; set; }
}
