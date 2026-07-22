const fs = require('fs');
const path = require('path');
const base = String.raw`f:\网站制作\网站\网站二`;
function w(rp, c) { const fp = path.join(base, rp); fs.mkdirSync(path.dirname(fp), { recursive: true }); fs.writeFileSync(fp, c, 'utf-8'); }
let n = 0; function W(rp, c) { w(rp, c); n++; }

// Controllers
W("src/PerfumeShop.Api/Controllers/SystemController.cs", `using Microsoft.AspNetCore.Mvc;
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
`);

W("src/PerfumeShop.Api/Controllers/InventoryController.cs", `using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class InventoryController : ControllerBase
{
    private readonly InventoryRepository _repo;
    public InventoryController(InventoryRepository repo) => _repo = repo;

    [HttpGet("products")]
    public async Task<IActionResult> GetProductInventories() => Ok(await _repo.GetProductInventoriesAsync());

    [HttpGet("materials")]
    public async Task<IActionResult> GetRawMaterials() => Ok(await _repo.GetRawMaterialInventoriesAsync());

    [HttpGet("bottles")]
    public async Task<IActionResult> GetBottles() => Ok(await _repo.GetBottleInventoriesAsync());

    [HttpGet("packaging")]
    public async Task<IActionResult> GetPackaging() => Ok(await _repo.GetPackagingInventoriesAsync());

    [HttpGet("movements")]
    public async Task<IActionResult> GetMovements(int page = 1, int pageSize = 20, string? itemType = null)
    {
        var (items, total) = await _repo.GetStockMovementsAsync(page, pageSize, itemType);
        return Ok(new { items, total });
    }

    [HttpGet("alerts/products")]
    public async Task<IActionResult> GetLowStockProducts(int threshold = 10) => Ok(await _repo.GetLowStockProductsAsync(threshold));

    [HttpGet("alerts/materials")]
    public async Task<IActionResult> GetLowStockMaterials(decimal threshold = 50) => Ok(await _repo.GetLowStockMaterialsAsync(threshold));
}
`);

W("src/PerfumeShop.Api/Controllers/LogisticsController.cs", `using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class LogisticsController : ControllerBase
{
    private readonly LogisticsRepository _repo;
    public LogisticsController(LogisticsRepository repo) => _repo = repo;

    [HttpGet("shipping")]
    public async Task<IActionResult> GetShippingOrders(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetShippingOrdersAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpGet("shipping/{id}")]
    public async Task<IActionResult> GetOrder(int id) { var o = await _repo.GetOrderAsync(id); return o == null ? NotFound() : Ok(o); }

    [HttpPut("shipping/{id}/status")]
    public async Task<IActionResult> UpdateStatus(int id, [FromBody] string status) { await _repo.UpdateOrderStatusAsync(id, status); return Ok(); }

    [HttpGet("returns")]
    public async Task<IActionResult> GetReturns(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetReturnsAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpGet("carriers")]
    public async Task<IActionResult> GetCarriers() => Ok(await _repo.GetCarriersAsync());

    [HttpPost("carriers")]
    public async Task<IActionResult> CreateCarrier([FromBody] ShippingCompany carrier) { await _repo.SaveCarrierAsync(carrier); return Ok(carrier); }

    [HttpPut("carriers/{id}")]
    public async Task<IActionResult> UpdateCarrier(int id, [FromBody] ShippingCompany carrier) { carrier.CompanyId = id; await _repo.SaveCarrierAsync(carrier); return Ok(); }

    [HttpDelete("carriers/{id}")]
    public async Task<IActionResult> DeleteCarrier(int id) { await _repo.DeleteCarrierAsync(id); return Ok(); }
}
`);

W("src/PerfumeShop.Api/Controllers/FinanceController.cs", `using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class FinanceController : ControllerBase
{
    private readonly FinanceRepository _repo;
    public FinanceController(FinanceRepository repo) => _repo = repo;

    [HttpGet("payables")]
    public async Task<IActionResult> GetPayables(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetPayablesAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpGet("payables/{id}")]
    public async Task<IActionResult> GetPayable(int id) { var p = await _repo.GetPayableAsync(id); return p == null ? NotFound() : Ok(p); }

    [HttpPost("payables")]
    public async Task<IActionResult> CreatePayable([FromBody] AccountsPayable p) { await _repo.SavePayableAsync(p); return Ok(p); }

    [HttpGet("receivables")]
    public async Task<IActionResult> GetReceivables(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetReceivablesAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpGet("budgets")]
    public async Task<IActionResult> GetBudgets(int page = 1, int pageSize = 20)
    {
        var (items, total) = await _repo.GetBudgetsAsync(page, pageSize);
        return Ok(new { items, total });
    }

    [HttpPost("budgets")]
    public async Task<IActionResult> CreateBudget([FromBody] BudgetPlan b) { await _repo.SaveBudgetAsync(b); return Ok(b); }

    [HttpGet("reconciliations")]
    public async Task<IActionResult> GetReconciliations(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetReconciliationsAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpGet("refunds")]
    public async Task<IActionResult> GetRefunds(int page = 1, int pageSize = 20, string? status = null)
    {
        var (items, total) = await _repo.GetRefundsAsync(page, pageSize, status);
        return Ok(new { items, total });
    }

    [HttpPut("refunds/{id}/approve")]
    public async Task<IActionResult> ApproveRefund(int id) { await _repo.ApproveRefundAsync(id); return Ok(); }

    [HttpGet("payment-records")]
    public async Task<IActionResult> GetPaymentRecords(int page = 1, int pageSize = 20)
    {
        var (items, total) = await _repo.GetPaymentRecordsAsync(page, pageSize);
        return Ok(new { items, total });
    }

    [HttpGet("expenses")]
    public async Task<IActionResult> GetExpenses(string? period = null) => Ok(await _repo.GetExpenseRecordsAsync(period));

    [HttpGet("cost-centers")]
    public async Task<IActionResult> GetCostCenters() => Ok(await _repo.GetCostCentersAsync());
}
`);

console.log(`${n} controllers created`);
