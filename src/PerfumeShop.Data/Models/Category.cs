using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Category
{
    public int CategoryId { get; set; }

    public string CategoryName { get; set; } = null!;

    public bool? IsActive { get; set; }

    public int? SortOrder { get; set; }
}
