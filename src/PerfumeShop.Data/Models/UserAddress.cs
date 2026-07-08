using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class UserAddress
{
    public string Address { get; set; } = null!;

    public int AddressId { get; set; }

    public string? City { get; set; }

    public string Consignee { get; set; } = null!;

    public DateTime? CreatedAt { get; set; }

    public string? District { get; set; }

    public bool? IsDefault { get; set; }

    public string Phone { get; set; } = null!;

    public string? Province { get; set; }

    public int UserId { get; set; }
}
