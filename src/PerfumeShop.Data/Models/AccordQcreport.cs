using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AccordQcreport
{
    public string? BatchNo { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Notes { get; set; }

    public int? ProductionId { get; set; }

    public int QcreportId { get; set; }

    public string? Qcresult { get; set; }

    public DateTime? TestDate { get; set; }

    public int? TesterId { get; set; }

    public string? TesterName { get; set; }
}
