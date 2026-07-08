using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FormulaNote
{
    public int FormulaId { get; set; }

    public int Id { get; set; }

    public int NoteId { get; set; }

    public int? Percentage { get; set; }
}
