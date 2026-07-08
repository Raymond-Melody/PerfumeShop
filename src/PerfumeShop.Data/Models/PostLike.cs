using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class PostLike
{
    public int LikeId { get; set; }

    public int PostId { get; set; }

    public int UserId { get; set; }

    public DateTime CreatedAt { get; set; }
}
