using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class UserFavorite
{
    public DateTime? CreatedTime { get; set; }

    public int FavoriteId { get; set; }

    public int ProductId { get; set; }

    public int UserId { get; set; }
}
