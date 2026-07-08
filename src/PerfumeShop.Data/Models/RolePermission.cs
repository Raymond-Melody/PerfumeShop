using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RolePermission
{
    public int PermId { get; set; }

    public int RoleId { get; set; }

    public string ModuleCode { get; set; } = null!;

    public bool? CanView { get; set; }

    public bool? CanCreate { get; set; }

    public bool? CanEdit { get; set; }

    public bool? CanDelete { get; set; }

    public bool? CanExport { get; set; }

    public bool? CanApprove { get; set; }
}
