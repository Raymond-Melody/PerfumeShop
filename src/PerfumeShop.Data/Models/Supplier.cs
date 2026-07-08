using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Supplier
{
    public string? Address { get; set; }

    public string? Category { get; set; }

    public string? ContactPerson { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Email { get; set; }

    public bool? IsActive { get; set; }

    public string? Notes { get; set; }

    public string? Phone { get; set; }

    public int SupplierId { get; set; }

    public string SupplierName { get; set; } = null!;
}
