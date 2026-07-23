using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Admin.Services;

/// <summary>操作粒度 — 对齐 V18 RolePermissions 表的 CanView/Create/Edit/Delete/Export/Approve。</summary>
public enum PermissionAction { View, Create, Edit, Delete, Export, Approve }

/// <summary>
/// 数据驱动 RBAC 服务 — 对齐 V18 role_auth.asp：
/// 模块访问基于 AdminRole.ModuleAccess（逗号分隔），操作粒度基于 RolePermissions。
/// SUPER_ADMIN 拥有全部权限。roleCode 由调用方从 AuthenticationState / ClaimsPrincipal 提供，
/// 因此可同时用于路由授权处理器（HTTP 与 Blazor 交互态）与组件按钮级控制。
/// </summary>
public interface IPermissionService
{
    Task<bool> CanAccessAsync(string? roleCode, string moduleCode);
    Task<bool> CanAsync(string? roleCode, string moduleCode, PermissionAction action);
    Task<HashSet<string>> GetAccessibleModulesAsync(string? roleCode);
}

public class PermissionService : IPermissionService
{
    private readonly DbContextOptions<PerfumeShopContext> _options;

    public PermissionService(DbContextOptions<PerfumeShopContext> options) => _options = options;

    public async Task<bool> CanAccessAsync(string? roleCode, string moduleCode)
    {
        if (IsSuper(roleCode)) return true;
        if (string.IsNullOrEmpty(roleCode) || string.IsNullOrEmpty(moduleCode)) return false;

        await using var db = new PerfumeShopContext(_options);
        var role = await db.AdminRoles.AsNoTracking().FirstOrDefaultAsync(r => r.RoleCode == roleCode);
        if (role == null) return false;

        var access = role.ModuleAccess ?? role.Permissions ?? "";
        return SplitModules(access).Contains(moduleCode, StringComparer.OrdinalIgnoreCase);
    }

    public async Task<bool> CanAsync(string? roleCode, string moduleCode, PermissionAction action)
    {
        if (IsSuper(roleCode)) return true;
        if (!await CanAccessAsync(roleCode, moduleCode)) return false;

        await using var db = new PerfumeShopContext(_options);
        var role = await db.AdminRoles.AsNoTracking().FirstOrDefaultAsync(r => r.RoleCode == roleCode);
        if (role == null) return false;

        var perm = await db.RolePermissions.AsNoTracking()
            .FirstOrDefaultAsync(p => p.RoleId == role.RoleId && p.ModuleCode == moduleCode);

        // 未配置操作级权限时，回退到"有模块访问即可 CRUD"（向后兼容当前行为）
        if (perm == null) return true;

        return action switch
        {
            PermissionAction.View => perm.CanView == true,
            PermissionAction.Create => perm.CanCreate == true,
            PermissionAction.Edit => perm.CanEdit == true,
            PermissionAction.Delete => perm.CanDelete == true,
            PermissionAction.Export => perm.CanExport == true,
            PermissionAction.Approve => perm.CanApprove == true,
            _ => false
        };
    }

    public async Task<HashSet<string>> GetAccessibleModulesAsync(string? roleCode)
    {
        if (IsSuper(roleCode))
            return new HashSet<string>(AllModules, StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(roleCode)) return new(StringComparer.OrdinalIgnoreCase);

        await using var db = new PerfumeShopContext(_options);
        var role = await db.AdminRoles.AsNoTracking().FirstOrDefaultAsync(r => r.RoleCode == roleCode);
        var access = role?.ModuleAccess ?? role?.Permissions ?? "";
        return new HashSet<string>(SplitModules(access), StringComparer.OrdinalIgnoreCase);
    }

    private static bool IsSuper(string? roleCode) => roleCode == "SUPER_ADMIN";

    private static readonly string[] AllModules =
    {
        "operation", "finance", "purchase", "techcenter", "system",
        "prodcenter", "semifinished", "logistics", "inventory", "analytics"
    };

    private static IEnumerable<string> SplitModules(string csv) =>
        (csv ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}
