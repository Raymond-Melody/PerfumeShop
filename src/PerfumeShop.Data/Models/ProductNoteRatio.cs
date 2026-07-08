using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductNoteRatio
{
    public int NoteId { get; set; }

    public string? NoteType { get; set; }

    public int Percentage { get; set; }

    public int ProductId { get; set; }

    public int RatioId { get; set; }
}
