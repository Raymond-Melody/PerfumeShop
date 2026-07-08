using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AdminRole
{
    public DateTime? CreatedAt { get; set; }

    public string? Description { get; set; }

    public string? Permissions { get; set; }

    public string RoleCode { get; set; } = null!;

    public int RoleId { get; set; }

    public string RoleName { get; set; } = null!;

    public DateTime? UpdatedAt { get; set; }

    public string? ModuleAccess { get; set; }
}
