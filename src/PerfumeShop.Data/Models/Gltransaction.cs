using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Gltransaction
{
    public int Glid { get; set; }

    public string? Glno { get; set; }

    public DateTime? TransactionDate { get; set; }

    public string? AccountCode { get; set; }

    public string? AccountName { get; set; }

    public decimal? DebitAmount { get; set; }

    public decimal? CreditAmount { get; set; }

    public int? CenterId { get; set; }

    public string? RefType { get; set; }

    public int? RefId { get; set; }

    public string? RefNo { get; set; }

    public string? Description { get; set; }

    public string? CreatedBy { get; set; }

    public DateTime? CreatedAt { get; set; }
}
