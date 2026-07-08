using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class WorkshopTransfer
{
    public DateTime? CreatedAt { get; set; }

    public string? FromWorkshop { get; set; }

    public DateTime? FulfilledAt { get; set; }

    public int? NoteId { get; set; }

    public string? Notes { get; set; }

    public DateTime? RequestedAt { get; set; }

    public string? RequestedBy { get; set; }

    public double? RequestQty { get; set; }

    public string? Status { get; set; }

    public string? ToWorkshop { get; set; }

    public int TransferId { get; set; }

    public string? TransferNo { get; set; }
}
