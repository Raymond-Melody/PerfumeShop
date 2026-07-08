using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductImage
{
    public int ImageId { get; set; }

    public int? ProductId { get; set; }

    public string? ImageUrl { get; set; }

    public int? ImageSize { get; set; }

    public int? SortOrder { get; set; }

    public bool? IsPrimary { get; set; }

    public DateTime? CreatedAt { get; set; }
}
