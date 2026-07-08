using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class InventoryTransaction
{
    public DateTime? CreatedAt { get; set; }

    public string? CreatedBy { get; set; }

    public int? MaterialId { get; set; }

    public int NoteId { get; set; }

    public string? Notes { get; set; }

    public int? ProductId { get; set; }

    public int Quantity { get; set; }

    public int? ReferenceOrderId { get; set; }

    public string? ReferenceType { get; set; }

    public string? TransactionDirection { get; set; }

    public int TransactionId { get; set; }

    public string? TransactionType { get; set; }

    public decimal? UnitCost { get; set; }
}
