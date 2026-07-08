using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AccordProduction
{
    public int? AccordRecipeId { get; set; }

    public double? ActualQty { get; set; }

    public string? ApprovedBy { get; set; }

    public string? BatchNo { get; set; }

    public DateTime? CompletedAt { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int? NoteId { get; set; }

    public string? NoteName { get; set; }

    public string? Notes { get; set; }

    public double? PlannedQty { get; set; }

    public int ProductionId { get; set; }

    public DateTime? StartedAt { get; set; }

    public string? Status { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public string? WorkCenter { get; set; }
}
