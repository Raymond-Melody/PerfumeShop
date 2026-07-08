using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class CommunityPost
{
    public int PostId { get; set; }

    public int UserId { get; set; }

    public string Title { get; set; } = null!;

    public string Content { get; set; } = null!;

    public string PostType { get; set; } = null!;

    public string? FragranceNotes { get; set; }

    public string? Tags { get; set; }

    public bool IsPublic { get; set; }

    public int LikeCount { get; set; }

    public int CommentCount { get; set; }

    public int ViewCount { get; set; }

    public bool IsPinned { get; set; }

    public bool IsActive { get; set; }

    public DateTime CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
