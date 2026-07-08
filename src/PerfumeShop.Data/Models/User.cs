using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class User
{
    public string? Address { get; set; }

    public string? City { get; set; }

    public DateTime? CreatedAt { get; set; }

    public string Email { get; set; } = null!;

    public string? FullName { get; set; }

    public bool? IsActive { get; set; }

    public bool? IsVip { get; set; }

    public string Password { get; set; } = null!;

    public string? Phone { get; set; }

    public int? Points { get; set; }

    public string? PostalCode { get; set; }

    public int UserId { get; set; }

    public string Username { get; set; } = null!;

    public string? UserRole { get; set; }

    public decimal? TotalSpent { get; set; }

    public int? OrderCount { get; set; }

    public DateTime? LastOrderDate { get; set; }

    public string? PreferredNote { get; set; }

    public string? CustomerTier { get; set; }

    public string? FavoriteCategory { get; set; }

    public int? ReferrerUserId { get; set; }

    public string? DeviceFingerprint { get; set; }

    public virtual ICollection<UserCoupon> UserCoupons { get; set; } = new List<UserCoupon>();
}
