using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class MaterialOutbound
{
    public string? ApprovedBy { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Notes { get; set; }

    public DateTime? OutboundDate { get; set; }

    public int OutboundId { get; set; }

    public string? OutboundNo { get; set; }

    public string? OutboundType { get; set; }

    public int? ReferenceId { get; set; }

    public string? ReferenceType { get; set; }

    public string? RequestedBy { get; set; }

    public string? Status { get; set; }
}
