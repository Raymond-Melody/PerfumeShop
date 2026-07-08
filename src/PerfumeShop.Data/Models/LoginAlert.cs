using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class LoginAlert
{
    public int AlertId { get; set; }

    public string? AlertType { get; set; }

    public string? AlertLevel { get; set; }

    public string? AlertMessage { get; set; }

    public string? Ipaddress { get; set; }

    public int? AdminId { get; set; }

    public bool? IsRead { get; set; }

    public DateTime? CreatedAt { get; set; }
}
