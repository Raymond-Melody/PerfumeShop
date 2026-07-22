using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class SystemRepository
{
    private readonly PerfumeShopContext _context;
    public SystemRepository(PerfumeShopContext context) => _context = context;

    // Users
    public async Task<(List<AdminUser> Items, int Total)> GetAdminUsersAsync(int page, int pageSize, string? search = null)
    {
        var q = _context.AdminUsers.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(search))
            q = q.Where(u => u.Username.Contains(search) || u.Email.Contains(search));
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(u => u.AdminId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AdminUser?> GetAdminUserAsync(int id) => await _context.AdminUsers.FindAsync(id);
    public async Task SaveAdminUserAsync(AdminUser user)
    {
        if (user.AdminId == 0) _context.AdminUsers.Add(user);
        else _context.AdminUsers.Update(user);
        await _context.SaveChangesAsync();
    }
    public async Task DeleteAdminUserAsync(int id)
    {
        var u = await _context.AdminUsers.FindAsync(id);
        if (u != null) { _context.AdminUsers.Remove(u); await _context.SaveChangesAsync(); }
    }

    // Roles
    public async Task<List<AdminRole>> GetRolesAsync() => await _context.AdminRoles.AsNoTracking().ToListAsync();
    public async Task<AdminRole?> GetRoleAsync(int id) => await _context.AdminRoles.FindAsync(id);
    public async Task SaveRoleAsync(AdminRole role)
    {
        if (role.RoleId == 0) _context.AdminRoles.Add(role);
        else _context.AdminRoles.Update(role);
        await _context.SaveChangesAsync();
    }
    public async Task DeleteRoleAsync(int id)
    {
        var r = await _context.AdminRoles.FindAsync(id);
        if (r != null) { _context.AdminRoles.Remove(r); await _context.SaveChangesAsync(); }
    }

    // AuditLogs
    public async Task<(List<AdminAuditLog> Items, int Total)> GetAuditLogsAsync(int page, int pageSize, string? actionType = null)
    {
        var q = _context.AdminAuditLogs.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(actionType)) q = q.Where(l => l.ActionType == actionType);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(l => l.LogId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AdminAuditLog?> GetAuditLogAsync(int id) => await _context.AdminAuditLogs.FindAsync(id);

    // SystemConfig
    public async Task<List<SiteSetting>> GetSettingsAsync() => await _context.SiteSettings.AsNoTracking().ToListAsync();
    public async Task SaveSettingAsync(SiteSetting setting)
    {
        _context.SiteSettings.Update(setting);
        await _context.SaveChangesAsync();
    }

    // AppLogs
    public async Task<(List<AppLog> Items, int Total)> GetAppLogsAsync(int page, int pageSize, string? logLevel = null, string? logType = null)
    {
        var q = _context.AppLogs.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(logLevel)) q = q.Where(l => l.LogLevel == logLevel);
        if (!string.IsNullOrWhiteSpace(logType)) q = q.Where(l => l.LogType == logType);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(l => l.LogId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AppLog?> GetAppLogAsync(long id) => await _context.AppLogs.FindAsync(id);

    // LoginAlerts
    public async Task<(List<LoginAlert> Items, int Total)> GetLoginAlertsAsync(int page, int pageSize)
    {
        var total = await _context.LoginAlerts.CountAsync();
        var items = await _context.LoginAlerts.AsNoTracking().OrderByDescending(a => a.AlertId)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }

    // ModulePermissions
    public async Task<List<ModulePermission>> GetModulePermissionsAsync() => await _context.ModulePermissions.AsNoTracking().ToListAsync();
    public async Task<List<RolePermission>> GetRolePermissionsAsync(int roleId) =>
        await _context.RolePermissions.AsNoTracking().Where(p => p.RoleId == roleId).ToListAsync();
    public async Task SaveRolePermissionAsync(RolePermission perm)
    {
        if (perm.PermId == 0) _context.RolePermissions.Add(perm);
        else _context.RolePermissions.Update(perm);
        await _context.SaveChangesAsync();
    }
}
