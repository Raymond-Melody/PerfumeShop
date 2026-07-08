using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipeNote
{
    public int Id { get; set; }

    public int? NoteId { get; set; }

    public string? NoteType { get; set; }

    public int? Percentage { get; set; }

    public int? RecipeId { get; set; }
}
