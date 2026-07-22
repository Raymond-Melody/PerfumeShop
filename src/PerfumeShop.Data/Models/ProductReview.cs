using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ProductReview
{
    public string? Comment { get; set; }

    public DateTime? CreatedAt { get; set; }

    public int OrderId { get; set; }

    public int? ProductId { get; set; }

    public int? Rating { get; set; }

    public int ReviewId { get; set; }

    public string? Status { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public int UserId { get; set; }

    public string? Title { get; set; }

    public bool IsVerifiedPurchase { get; set; }

    public string? AIFeelingSummary { get; set; }

    public int LikeCount { get; set; }

    public ICollection<ReviewImage> Images { get; set; } = new List<ReviewImage>();
}
