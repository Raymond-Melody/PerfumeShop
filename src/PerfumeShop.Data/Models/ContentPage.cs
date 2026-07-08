using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ContentPage
{
    public int PageId { get; set; }

    public string Title { get; set; } = null!;

    public string? Slug { get; set; }

    public string? Content { get; set; }

    public string? MetaDescription { get; set; }

    public bool? IsPublished { get; set; }

    public int? SortOrder { get; set; }

    public int? UpdatedBy { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public DateTime? CreatedAt { get; set; }
}
