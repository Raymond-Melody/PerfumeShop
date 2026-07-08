using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductTypeConfig
{
    public int ConfigId { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public string? DisplayName { get; set; }

    public int? DisplayOrder { get; set; }

    public string? Icon { get; set; }

    public bool? IsActive { get; set; }

    public string? NavName { get; set; }

    public bool? RequiresRatio { get; set; }

    public bool? RequiresReview { get; set; }

    public string TypeCode { get; set; } = null!;
}
