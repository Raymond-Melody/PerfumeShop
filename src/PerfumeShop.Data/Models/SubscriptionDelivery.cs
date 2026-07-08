using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class SubscriptionDelivery
{
    public int DeliveryId { get; set; }

    public int SubscriptionId { get; set; }

    public DateTime DeliveryDate { get; set; }

    public string Status { get; set; } = null!;

    public string? TrackingNo { get; set; }

    public DateTime CreatedAt { get; set; }
}
