using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PurchaseCategory
{
    public string? CategoryCode { get; set; }

    public int CategoryId { get; set; }

    public string? CategoryName { get; set; }

    public string? Description { get; set; }

    public int? DisplayOrder { get; set; }

    public bool? IsActive { get; set; }
}
