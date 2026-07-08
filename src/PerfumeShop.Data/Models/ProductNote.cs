using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductNote
{
    public int NoteId { get; set; }

    public int ProductId { get; set; }

    public int ProductNoteId { get; set; }
}
