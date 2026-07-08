using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Cart
{
    public int? BaseNoteId { get; set; }

    public int? BottleId { get; set; }

    public int CartId { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? CustomLabel { get; set; }

    public int? MiddleNoteId { get; set; }

    public int ProductId { get; set; }

    public int? Quantity { get; set; }

    public string? SessionId { get; set; }

    public int? TopNoteId { get; set; }

    public decimal UnitPrice { get; set; }

    public int? UserId { get; set; }

    public int? VolumeId { get; set; }
}
