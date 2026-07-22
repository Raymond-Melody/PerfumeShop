using Microsoft.EntityFrameworkCore;

namespace PerfumeShop.Data.Models;

/// <summary>
/// PerfumeShopContext 扩展 — M3-A 新增实体
/// </summary>
public partial class PerfumeShopContext
{
    public virtual DbSet<PasswordResetToken> PasswordResetTokens { get; set; }
}
