using System;

namespace PerfumeShop.Data.Models;

public partial class ReviewImage
{
    public int ImageId { get; set; }

    public int ReviewId { get; set; }

    public string ImageUrl { get; set; } = string.Empty;

    public int SortOrder { get; set; }

    public DateTime CreatedAt { get; set; }

    public ProductReview? Review { get; set; }
}
