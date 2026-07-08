using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class OrderDetail
{
    public string? BaseNoteName { get; set; }

    public string? BottleName { get; set; }

    public string? CustomLabel { get; set; }

    public int DetailId { get; set; }

    public string? MiddleNoteName { get; set; }

    public int OrderId { get; set; }

    public int ProductId { get; set; }

    public string? ProductName { get; set; }

    public int Quantity { get; set; }

    public decimal Subtotal { get; set; }

    public string? TopNoteName { get; set; }

    public decimal UnitPrice { get; set; }

    public int? VolumeMl { get; set; }

    public string? VolumeName { get; set; }
}
