using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class MarketingCampaign
{
    public int CampaignId { get; set; }

    public string? CampaignName { get; set; }

    public string? CampaignType { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public decimal? DiscountValue { get; set; }

    public DateTime? EndDate { get; set; }

    public bool? IsActive { get; set; }

    public decimal? MinPurchase { get; set; }

    public int? ParticipantCount { get; set; }

    public DateTime? StartDate { get; set; }

    public decimal? TotalSales { get; set; }
}
