using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FragranceNote
{
    public int? BaseNoteId { get; set; }

    public string? Description { get; set; }

    public string? ImageUrl { get; set; }

    public string? Ingredients { get; set; }

    public bool? IsActive { get; set; }

    public int? IsBaseNote { get; set; }

    public int NoteId { get; set; }

    public string NoteName { get; set; } = null!;

    public string NoteType { get; set; } = null!;

    public decimal? PriceAddition { get; set; }

    public int? RecommendedPercentage { get; set; }
}
