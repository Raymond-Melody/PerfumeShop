using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class BaseNote
{
    public int BaseNoteId { get; set; }

    public string BaseNoteName { get; set; } = null!;

    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public string? Ingredients { get; set; }

    public bool? IsActive { get; set; }

    public decimal? UnitPrice { get; set; }
}
