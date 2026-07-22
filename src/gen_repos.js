const fs = require('fs');
const path = require('path');

const base = String.raw`f:\网站制作\网站\网站二`;

function write(relpath, content) {
  const fullpath = path.join(base, relpath);
  fs.mkdirSync(path.dirname(fullpath), { recursive: true });
  fs.writeFileSync(fullpath, content, 'utf-8');
}

let count = 0;
function w(relpath, content) { write(relpath, content); count++; }

// ===== REPOSITORIES =====
w("src/PerfumeShop.Data/Repositories/InventoryRepository.cs", `using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class InventoryRepository
{
    private readonly PerfumeShopContext _context;
    public InventoryRepository(PerfumeShopContext context) => _context = context;

    public async Task<List<ProductInventory>> GetProductInventoriesAsync() =>
        await _context.ProductInventories.AsNoTracking().ToListAsync();
    public async Task<List<RawMaterialInventory>> GetRawMaterialInventoriesAsync() =>
        await _context.RawMaterialInventories.AsNoTracking().ToListAsync();
    public async Task<List<BottleInventory>> GetBottleInventoriesAsync() =>
        await _context.BottleInventories.AsNoTracking().ToListAsync();
    public async Task<List<PackagingInventory>> GetPackagingInventoriesAsync() =>
        await _context.PackagingInventories.AsNoTracking().ToListAsync();
    public async Task<(List<StockMovement> Items, int Total)> GetStockMovementsAsync(int page, int pageSize, string? itemType = null)
    {
        var q = _context.StockMovements.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(itemType)) q = q.Where(m => m.ItemType == itemType);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(m => m.MovementId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<List<InventoryBatch>> GetInventoryBatchesAsync() =>
        await _context.InventoryBatches.AsNoTracking().ToListAsync();
    public async Task<(List<InventoryTransaction> Items, int Total)> GetInventoryTransactionsAsync(int page, int pageSize)
    {
        var total = await _context.InventoryTransactions.CountAsync();
        var items = await _context.InventoryTransactions.AsNoTracking()
            .OrderByDescending(t => t.TransactionId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<List<ProductInventory>> GetLowStockProductsAsync(int threshold = 10) =>
        await _context.ProductInventories.AsNoTracking().Where(p => p.CurrentStock <= threshold).ToListAsync();
    public async Task<List<RawMaterialInventory>> GetLowStockMaterialsAsync(decimal threshold = 50) =>
        await _context.RawMaterialInventories.AsNoTracking().Where(m => m.CurrentQuantity <= threshold).ToListAsync();
}
`);

w("src/PerfumeShop.Data/Repositories/LogisticsRepository.cs", `using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class LogisticsRepository
{
    private readonly PerfumeShopContext _context;
    public LogisticsRepository(PerfumeShopContext context) => _context = context;

    public async Task<(List<Order> Items, int Total)> GetShippingOrdersAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.Orders.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(o => o.Status == status);
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
    public async Task<(List<AfterSale> Items, int Total)> GetReturnsAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.AfterSales.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(a => a.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(a => a.AfterSaleId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AfterSale?> GetReturnAsync(int id) => await _context.AfterSales.FindAsync(id);
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
}
`);

w("src/PerfumeShop.Data/Repositories/FinanceRepository.cs", `using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class FinanceRepository
{
    private readonly PerfumeShopContext _context;
    public FinanceRepository(PerfumeShopContext context) => _context = context;

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
        if (p.PayableId == 0) _context.AccountsPayables.Add(p); else _context.AccountsPayables.Update(p);
        await _context.SaveChangesAsync();
    }
    public async Task<(List<AccountsReceivable> Items, int Total)> GetReceivablesAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.AccountsReceivables.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(r => r.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(r => r.ReceivableId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AccountsReceivable?> GetReceivableAsync(int id) => await _context.AccountsReceivables.FindAsync(id);
    public async Task<(List<BudgetPlan> Items, int Total)> GetBudgetsAsync(int page, int pageSize)
    {
        var total = await _context.BudgetPlans.CountAsync();
        var items = await _context.BudgetPlans.AsNoTracking().OrderByDescending(b => b.BudgetId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<BudgetPlan?> GetBudgetAsync(int id) => await _context.BudgetPlans.FindAsync(id);
    public async Task SaveBudgetAsync(BudgetPlan b)
    {
        if (b.BudgetId == 0) _context.BudgetPlans.Add(b); else _context.BudgetPlans.Update(b);
        await _context.SaveChangesAsync();
    }
    public async Task DeleteBudgetAsync(int id)
    {
        var b = await _context.BudgetPlans.FindAsync(id);
        if (b != null) { _context.BudgetPlans.Remove(b); await _context.SaveChangesAsync(); }
    }
    public async Task<(List<ReconciliationLog> Items, int Total)> GetReconciliationsAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.ReconciliationLogs.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(r => r.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(r => r.LogId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<ReconciliationLog?> GetReconciliationAsync(int id) => await _context.ReconciliationLogs.FindAsync(id);
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
        await _context.RefundRecords.Where(r => r.RefundId == id).ExecuteUpdateAsync(s => s.SetProperty(r => r.Status, "approved"));
    }
    public async Task<(List<PaymentRecord> Items, int Total)> GetPaymentRecordsAsync(int page, int pageSize)
    {
        var total = await _context.PaymentRecords.CountAsync();
        var items = await _context.PaymentRecords.AsNoTracking().OrderByDescending(p => p.RecordId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<PaymentRecord?> GetPaymentRecordAsync(int id) => await _context.PaymentRecords.FindAsync(id);
    public async Task<(List<Gltransaction> Items, int Total)> GetGlTransactionsAsync(int page, int pageSize)
    {
        var total = await _context.Gltransactions.CountAsync();
        var items = await _context.Gltransactions.AsNoTracking().OrderByDescending(g => g.Glid).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<List<ExpenseRecord>> GetExpenseRecordsAsync(string? period = null)
    {
        var q = _context.ExpenseRecords.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(period)) q = q.Where(e => e.Period == period);
        return await q.OrderByDescending(e => e.ExpenseId).ToListAsync();
    }
    public async Task<List<CostCenter>> GetCostCentersAsync() => await _context.CostCenters.AsNoTracking().ToListAsync();
}
`);

console.log(`Generated ${count} repository files`);
