$ErrorActionPreference = "Stop"
$base = "f:\网站制作\网站\网站二"

# ============ REPOSITORIES ============

$systemRepo = @'
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

    // AppLogs (ErrorLog / OperationLog / LoginHistory)
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
'@

$inventoryRepo = @'
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class InventoryRepository
{
    private readonly PerfumeShopContext _context;
    public InventoryRepository(PerfumeShopContext context) => _context = context;

    // Product Inventory
    public async Task<List<ProductInventory>> GetProductInventoriesAsync() =>
        await _context.ProductInventories.AsNoTracking().ToListAsync();

    // Raw Material Inventory
    public async Task<List<RawMaterialInventory>> GetRawMaterialInventoriesAsync() =>
        await _context.RawMaterialInventories.AsNoTracking().ToListAsync();

    // Bottle Inventory
    public async Task<List<BottleInventory>> GetBottleInventoriesAsync() =>
        await _context.BottleInventories.AsNoTracking().ToListAsync();

    // Packaging Inventory
    public async Task<List<PackagingInventory>> GetPackagingInventoriesAsync() =>
        await _context.PackagingInventories.AsNoTracking().ToListAsync();

    // Stock Movements
    public async Task<(List<StockMovement> Items, int Total)> GetStockMovementsAsync(int page, int pageSize, string? itemType = null)
    {
        var q = _context.StockMovements.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(itemType)) q = q.Where(m => m.ItemType == itemType);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(m => m.MovementId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }

    // Inventory Batches
    public async Task<List<InventoryBatch>> GetInventoryBatchesAsync() =>
        await _context.InventoryBatches.AsNoTracking().ToListAsync();

    // Inventory Transactions
    public async Task<(List<InventoryTransaction> Items, int Total)> GetInventoryTransactionsAsync(int page, int pageSize)
    {
        var total = await _context.InventoryTransactions.CountAsync();
        var items = await _context.InventoryTransactions.AsNoTracking()
            .OrderByDescending(t => t.TransactionId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }

    // Alerts - low stock items
    public async Task<List<ProductInventory>> GetLowStockProductsAsync(int threshold = 10) =>
        await _context.ProductInventories.AsNoTracking()
            .Where(p => p.CurrentStock <= threshold).ToListAsync();

    public async Task<List<RawMaterialInventory>> GetLowStockMaterialsAsync(decimal threshold = 50) =>
        await _context.RawMaterialInventories.AsNoTracking()
            .Where(m => m.CurrentQuantity <= threshold).ToListAsync();
}
'@

$logisticsRepo = @'
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class LogisticsRepository
{
    private readonly PerfumeShopContext _context;
    public LogisticsRepository(PerfumeShopContext context) => _context = context;

    // Shipping (Orders with shipping status)
    public async Task<(List<Order> Items, int Total)> GetShippingOrdersAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.Orders.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(o => o.Status == status);
        else q = q.Where(o => o.Status == "shipped" || o.Status == "paid" || o.Status == "confirmed");
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(o => o.OrderId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<Order?> GetOrderAsync(int id) => await _context.Orders.FindAsync(id);
    public async Task UpdateOrderStatusAsync(int orderId, string status)
    {
        await _context.Orders.Where(o => o.OrderId == orderId)
            .ExecuteUpdateAsync(s => s.SetProperty(o => o.Status, status));
    }

    // Returns
    public async Task<(List<AfterSale> Items, int Total)> GetReturnsAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.AfterSales.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(a => a.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(a => a.AfterSaleId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AfterSale?> GetReturnAsync(int id) => await _context.AfterSales.FindAsync(id);

    // Carriers
    public async Task<List<ShippingCompany>> GetCarriersAsync() =>
        await _context.ShippingCompanies.AsNoTracking().ToListAsync();
    public async Task<ShippingCompany?> GetCarrierAsync(int id) => await _context.ShippingCompanies.FindAsync(id);
    public async Task SaveCarrierAsync(ShippingCompany carrier)
    {
        if (carrier.CompanyId == 0) _context.ShippingCompanies.Add(carrier);
        else _context.ShippingCompanies.Update(carrier);
        await _context.SaveChangesAsync();
    }
    public async Task DeleteCarrierAsync(int id)
    {
        var c = await _context.ShippingCompanies.FindAsync(id);
        if (c != null) { _context.ShippingCompanies.Remove(c); await _context.SaveChangesAsync(); }
    }

    // Report stats
    public async Task<object> GetLogisticsStatsAsync()
    {
        var totalShipped = await _context.Orders.CountAsync(o => o.Status == "shipped");
        var totalDelivered = await _context.Orders.CountAsync(o => o.Status == "delivered");
        var totalReturns = await _context.AfterSales.CountAsync();
        var pendingReturns = await _context.AfterSales.CountAsync(a => a.Status == "pending");
        return new { totalShipped, totalDelivered, totalReturns, pendingReturns };
    }
}
'@

$financeRepo = @'
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class FinanceRepository
{
    private readonly PerfumeShopContext _context;
    public FinanceRepository(PerfumeShopContext context) => _context = context;

    // Payables
    public async Task<(List<AccountsPayable> Items, int Total)> GetPayablesAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.AccountsPayables.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(p => p.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(p => p.PayableId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AccountsPayable?> GetPayableAsync(int id) => await _context.AccountsPayables.FindAsync(id);
    public async Task SavePayableAsync(AccountsPayable p)
    {
        if (p.PayableId == 0) _context.AccountsPayables.Add(p);
        else _context.AccountsPayables.Update(p);
        await _context.SaveChangesAsync();
    }

    // Receivables
    public async Task<(List<AccountsReceivable> Items, int Total)> GetReceivablesAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.AccountsReceivables.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(r => r.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(r => r.ReceivableId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AccountsReceivable?> GetReceivableAsync(int id) => await _context.AccountsReceivables.FindAsync(id);

    // Budgets
    public async Task<(List<BudgetPlan> Items, int Total)> GetBudgetsAsync(int page, int pageSize)
    {
        var total = await _context.BudgetPlans.CountAsync();
        var items = await _context.BudgetPlans.AsNoTracking().OrderByDescending(b => b.BudgetId)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<BudgetPlan?> GetBudgetAsync(int id) => await _context.BudgetPlans.FindAsync(id);
    public async Task SaveBudgetAsync(BudgetPlan b)
    {
        if (b.BudgetId == 0) _context.BudgetPlans.Add(b);
        else _context.BudgetPlans.Update(b);
        await _context.SaveChangesAsync();
    }
    public async Task DeleteBudgetAsync(int id)
    {
        var b = await _context.BudgetPlans.FindAsync(id);
        if (b != null) { _context.BudgetPlans.Remove(b); await _context.SaveChangesAsync(); }
    }

    // Reconciliations
    public async Task<(List<ReconciliationLog> Items, int Total)> GetReconciliationsAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.ReconciliationLogs.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(r => r.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(r => r.LogId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<ReconciliationLog?> GetReconciliationAsync(int id) => await _context.ReconciliationLogs.FindAsync(id);

    // Refunds
    public async Task<(List<RefundRecord> Items, int Total)> GetRefundsAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.RefundRecords.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(r => r.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(r => r.RefundId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<RefundRecord?> GetRefundAsync(int id) => await _context.RefundRecords.FindAsync(id);
    public async Task ApproveRefundAsync(int id)
    {
        await _context.RefundRecords.Where(r => r.RefundId == id)
            .ExecuteUpdateAsync(s => s.SetProperty(r => r.Status, "approved"));
    }

    // Payment Records (for invoices)
    public async Task<(List<PaymentRecord> Items, int Total)> GetPaymentRecordsAsync(int page, int pageSize)
    {
        var total = await _context.PaymentRecords.CountAsync();
        var items = await _context.PaymentRecords.AsNoTracking().OrderByDescending(p => p.RecordId)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<PaymentRecord?> GetPaymentRecordAsync(int id) => await _context.PaymentRecords.FindAsync(id);

    // GL Transactions
    public async Task<(List<Gltransaction> Items, int Total)> GetGlTransactionsAsync(int page, int pageSize)
    {
        var total = await _context.Gltransactions.CountAsync();
        var items = await _context.Gltransactions.AsNoTracking().OrderByDescending(g => g.Glid)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }

    // Expense Records (cost analysis)
    public async Task<List<ExpenseRecord>> GetExpenseRecordsAsync(string? period = null)
    {
        var q = _context.ExpenseRecords.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(period)) q = q.Where(e => e.Period == period);
        return await q.OrderByDescending(e => e.ExpenseId).ToListAsync();
    }

    // Cost Centers
    public async Task<List<CostCenter>> GetCostCentersAsync() =>
        await _context.CostCenters.AsNoTracking().ToListAsync();

    // Dashboard stats
    public async Task<object> GetDashboardStatsAsync()
    {
        var totalRevenue = await _context.Orders.Where(o => o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        var totalPayable = await _context.AccountsPayables.SumAsync(p => (decimal?)p.Amount) ?? 0;
        var totalReceivable = await _context.AccountsReceivables.SumAsync(r => (decimal?)r.Amount) ?? 0;
        var totalRefunds = await _context.RefundRecords.SumAsync(r => (decimal?)r.RefundAmount) ?? 0;
        var pendingReconciliations = await _context.ReconciliationLogs.CountAsync(r => r.Status == "pending");
        return new { totalRevenue, totalPayable, totalReceivable, totalRefunds, pendingReconciliations };
    }
}
'@

[System.IO.File]::WriteAllText("$base\src\PerfumeShop.Data\Repositories\SystemRepository.cs", $systemRepo, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText("$base\src\PerfumeShop.Data\Repositories\InventoryRepository.cs", $inventoryRepo, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText("$base\src\PerfumeShop.Data\Repositories\LogisticsRepository.cs", $logisticsRepo, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText("$base\src\PerfumeShop.Data\Repositories\FinanceRepository.cs", $financeRepo, [System.Text.Encoding]::UTF8)

Write-Host "4 repositories created"
