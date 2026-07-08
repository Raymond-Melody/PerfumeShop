using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class AdminUser
{
    public int AdminId { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string? Department { get; set; }

    public string Email { get; set; } = null!;

    public string? FullName { get; set; }

    public bool? IsActive { get; set; }

    public bool? IsLocked { get; set; }

    public DateTime? LastLogin { get; set; }

    public string? PasswordHash { get; set; }

    public string? ResetToken { get; set; }

    public DateTime? ResetTokenExpiry { get; set; }

    public int? RoleId { get; set; }

    public string Username { get; set; } = null!;
}
