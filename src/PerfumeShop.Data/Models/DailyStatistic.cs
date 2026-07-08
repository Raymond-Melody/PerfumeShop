using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class DailyStatistic
{
    public DateTime? CreatedAt { get; set; }

    public string? DataJson { get; set; }

    public int? NewUsers { get; set; }

    public DateTime StatDate { get; set; }

    public int StatId { get; set; }

    public int? TopNoteId { get; set; }

    public int? TopProductId { get; set; }

    public int? TotalOrders { get; set; }

    public decimal? TotalRevenue { get; set; }

    public int? TotalUsers { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
