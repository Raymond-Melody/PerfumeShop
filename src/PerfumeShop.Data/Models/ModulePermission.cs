using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ModulePermission
{
    public bool? IsActive { get; set; }

    public string ModuleCode { get; set; } = null!;

    public string ModuleName { get; set; } = null!;

    public string? ParentModule { get; set; }

    public int PermissionId { get; set; }

    public int? PermissionLevel { get; set; }

    public string? RequiredRole { get; set; }

    public string? Urlpattern { get; set; }
}
