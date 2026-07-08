using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ShippingCompany
{
    public int CompanyId { get; set; }

    public string CompanyName { get; set; } = null!;

    public string? ContactPerson { get; set; }

    public string? ContactPhone { get; set; }

    public string? Website { get; set; }

    public bool? IsActive { get; set; }

    public string? Notes { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }
}
