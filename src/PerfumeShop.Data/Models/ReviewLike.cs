using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ReviewLike
{
    public int LikeId { get; set; }

    public int ReviewId { get; set; }

    public int UserId { get; set; }

    public DateTime CreatedAt { get; set; }
}
