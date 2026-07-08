using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class FragranceIngredient
{
    public DateTime? CreatedAt { get; set; }

    public int FragranceIngredientId { get; set; }

    public int IngredientId { get; set; }

    public int NoteId { get; set; }

    public float Percentage { get; set; }

    public int? SortOrder { get; set; }
}
