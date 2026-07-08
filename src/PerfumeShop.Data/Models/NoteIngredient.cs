using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class NoteIngredient
{
    public int BaseNoteId { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int Id { get; set; }

    public int NoteId { get; set; }

    public double? Percentage { get; set; }
}
