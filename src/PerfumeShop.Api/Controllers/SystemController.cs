using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class SystemController : ControllerBase
{
    private readonly SystemRepository _repo;
    public SystemController(SystemRepository repo) => _repo = repo;

    [HttpGet("admin-users")]
    public async Task<IActionResult> GetAdminUsers(int page = 1, int pageSize = 20, string? search = null)
    {
        var (items, total) = await _repo.GetAdminUsersAsync(page, pageSize, search);
        return Ok(new { items, total });
    }

    [HttpGet("admin-users/{id}")]
    public async Task<IActionResult> GetAdminUser(int id) { var u = await _repo.GetAdminUserAsync(id); return u == null ? NotFound() : Ok(u); }

    [HttpPost("admin-users")]
    public async Task<IActionResult> CreateAdminUser([FromBody] AdminUser user) { await _repo.SaveAdminUserAsync(user); return Ok(user); }

    [HttpPut("admin-users/{id}")]
    public async Task<IActionResult> UpdateAdminUser(int id, [FromBody] AdminUser user) { user.AdminId = id; await _repo.SaveAdminUserAsync(user); return Ok(); }

    [HttpDelete("admin-users/{id}")]
    public async Task<IActionResult> DeleteAdminUser(int id) { await _repo.DeleteAdminUserAsync(id); return Ok(); }

    [HttpGet("roles")]
    public async Task<IActionResult> GetRoles() => Ok(await _repo.GetRolesAsync());

    [HttpGet("roles/{id}")]
    public async Task<IActionResult> GetRole(int id) { var r = await _repo.GetRoleAsync(id); return r == null ? NotFound() : Ok(r); }

    [HttpPost("roles")]
    public async Task<IActionResult> CreateRole([FromBody] AdminRole role) { await _repo.SaveRoleAsync(role); return Ok(role); }

    [HttpPut("roles/{id}")]
    public async Task<IActionResult> UpdateRole(int id, [FromBody] AdminRole role) { role.RoleId = id; await _repo.SaveRoleAsync(role); return Ok(); }

    [HttpDelete("roles/{id}")]
    public async Task<IActionResult> DeleteRole(int id) { await _repo.DeleteRoleAsync(id); return Ok(); }

    [HttpGet("audit-logs")]
    public async Task<IActionResult> GetAuditLogs(int page = 1, int pageSize = 20, string? actionType = null)
    {
        var (items, total) = await _repo.GetAuditLogsAsync(page, pageSize, actionType);
        return Ok(new { items, total });
    }

    [HttpGet("settings")]
    public async Task<IActionResult> GetSettings() => Ok(await _repo.GetSettingsAsync());

    [HttpGet("app-logs")]
    public async Task<IActionResult> GetAppLogs(int page = 1, int pageSize = 20, string? logLevel = null, string? logType = null)
    {
        var (items, total) = await _repo.GetAppLogsAsync(page, pageSize, logLevel, logType);
        return Ok(new { items, total });
    }
}
