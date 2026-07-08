using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class OrderDetailNoteSelection
{
    public int DetailId { get; set; }

    public int NoteId { get; set; }

    public string? NoteType { get; set; }

    public int Percentage { get; set; }

    public int SelectionId { get; set; }
}
