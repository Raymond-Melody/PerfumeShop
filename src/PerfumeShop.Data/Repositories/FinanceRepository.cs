using Microsoft.EntityFrameworkCore;
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

    // ===== V19.4 报表查询扩展（对齐 dal_finance.asp + index.asp）=====

    public async Task<decimal> GetTotalRevenueAsync(string? startDate = null, string? endDate = null)
    {
        var q = _context.Orders.AsNoTracking()
            .Where(o => o.Status != null && (new[] { "Paid", "Processing", "Shipped", "Completed" }).Contains(o.Status));
        if (!string.IsNullOrEmpty(startDate) && DateTime.TryParse(startDate, out var sd))
            q = q.Where(o => o.CreatedAt >= sd);
        if (!string.IsNullOrEmpty(endDate) && DateTime.TryParse(endDate, out var ed))
            q = q.Where(o => o.CreatedAt <= ed);
        return await q.SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
    }

    public async Task<decimal> GetTotalCostAsync()
    {
        return await _context.Orders.AsNoTracking()
            .Where(o => o.Status != null && (new[] { "Paid", "Processing", "Shipped", "Completed" }).Contains(o.Status))
            .SumAsync(o => (decimal?)o.CostAmount) ?? 0;
    }

    public async Task<decimal> GetTotalRefundAsync()
    {
        return await _context.RefundRecords.AsNoTracking()
            .Where(r => r.Status == "Completed")
            .SumAsync(r => (decimal?)r.RefundAmount) ?? 0;
    }

    public async Task<int> GetPendingReconCountAsync()
    {
        return await _context.ReconciliationLogs.AsNoTracking()
            .CountAsync(r => r.Status != null && !(new[] { "Matched", "Resolved" }).Contains(r.Status));
    }

    public async Task<int> GetReconAlertCountAsync()
    {
        return await _context.ReconciliationLogs.AsNoTracking()
            .CountAsync(r => r.Status == "Exception" || r.Status == "Pending");
    }

    public async Task<int> GetBudgetAlertCountAsync()
    {
        return await _context.BudgetPlans.AsNoTracking()
            .CountAsync(b => b.Status == "Active" && b.BudgetAmount > 0 && (b.ActualAmount / b.BudgetAmount) * 100 > 90);
    }

    public async Task<int> GetFundAlertCountAsync()
    {
        return await _context.FundAccounts.AsNoTracking()
            .CountAsync(f => f.IsActive == true && f.TotalBalance < f.AlertThreshold);
    }

    // ===== V19.X 资金看板功能对齐 =====

    /// <summary>交易在途：Orders.Status='Paid' 未结算金额</summary>
    public async Task<decimal> GetPendingOrdersAmountAsync()
    {
        return await _context.Orders.AsNoTracking()
            .Where(o => o.Status == "Paid")
            .SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
    }

    /// <summary>提现中：PaymentRecords Status='Pending' & TransactionType='Transfer' 金额</summary>
    public async Task<decimal> GetPendingWithdrawAmountAsync()
    {
        return await _context.PaymentRecords.AsNoTracking()
            .Where(p => p.Status == "Pending" && p.TransactionType == "Transfer")
            .SumAsync(p => (decimal?)p.Amount) ?? 0;
    }

    /// <summary>大额转账检测：单笔≥10000元的转账笔数</summary>
    public async Task<int> GetLargeTransferCountAsync()
    {
        return await _context.PaymentRecords.AsNoTracking()
            .CountAsync(p => p.TransactionType == "Transfer" && p.Amount >= 10000);
    }

    /// <summary>资金预警明细（AvailableBalance < AlertThreshold 的活跃账户）</summary>
    public async Task<List<FundAlertRow>> GetFundAlertsDetailAsync()
    {
        return await _context.FundAccounts.AsNoTracking()
            .Where(f => f.IsActive == true && f.AvailableBalance < f.AlertThreshold)
            .OrderByDescending(f => (f.AlertThreshold ?? 0) - (f.AvailableBalance ?? 0))
            .Select(f => new FundAlertRow
            {
                AccountId = f.AccountId,
                AccountName = f.AccountName ?? "",
                AvailableBalance = f.AvailableBalance ?? 0,
                AlertThreshold = f.AlertThreshold ?? 0,
                DiffAmount = (f.AlertThreshold ?? 0) - (f.AvailableBalance ?? 0),
                AlertLevel = (f.AlertThreshold ?? 0) > 0 && (f.AvailableBalance ?? 0) < (f.AlertThreshold ?? 0) * 0.5m ? "高危" :
                             (f.AlertThreshold ?? 0) > 0 && (f.AvailableBalance ?? 0) < (f.AlertThreshold ?? 0) * 0.8m ? "中危" : "低危"
            })
            .ToListAsync();
    }

    /// <summary>交易在途明细（Orders.Status='Paid'）</summary>
    public async Task<List<PendingOrderRow>> GetPendingOrdersDetailAsync()
    {
        return await _context.Orders.AsNoTracking()
            .Where(o => o.Status == "Paid")
            .OrderByDescending(o => o.CreatedAt)
            .Take(50)
            .Select(o => new PendingOrderRow
            {
                OrderId = o.OrderId,
                OrderNo = o.OrderNo ?? "",
                Amount = o.TotalAmount,
                Method = o.PaymentMethod ?? "",
                CreatedAt = o.CreatedAt
            })
            .ToListAsync();
    }

    /// <summary>提现中明细</summary>
    public async Task<List<PendingWithdrawRow>> GetPendingWithdrawDetailAsync()
    {
        return await _context.PaymentRecords.AsNoTracking()
            .Where(p => p.Status == "Pending" && p.TransactionType == "Transfer")
            .OrderByDescending(p => p.CreatedAt)
            .Take(50)
            .Select(p => new PendingWithdrawRow
            {
                RecordId = p.RecordId,
                OrderNo = p.OrderNo ?? "",
                Amount = p.Amount ?? 0,
                CreatedAt = p.CreatedAt
            })
            .ToListAsync();
    }

    /// <summary>不可用保证金明细（FrozenAmount > 0）</summary>
    public async Task<List<PendingFrozenRow>> GetFrozenDetailAsync()
    {
        return await _context.FundAccounts.AsNoTracking()
            .Where(f => f.FrozenAmount > 0 && f.IsActive == true)
            .OrderByDescending(f => f.FrozenAmount)
            .Select(f => new PendingFrozenRow
            {
                AccountId = f.AccountId,
                AccountName = f.AccountName ?? "",
                FrozenAmount = f.FrozenAmount ?? 0
            })
            .ToListAsync();
    }

    /// <summary>保存资金账户（新增或编辑）</summary>
    public async Task SaveFundAccountAsync(FundAccount account)
    {
        if (account.AccountId == 0)
        {
            account.CreatedAt = DateTime.Now;
            account.UpdatedAt = DateTime.Now;
            _context.FundAccounts.Add(account);
        }
        else
        {
            var existing = await _context.FundAccounts.FindAsync(account.AccountId);
            if (existing != null)
            {
                existing.AccountName = account.AccountName;
                existing.AccountType = account.AccountType;
                existing.TotalBalance = account.TotalBalance;
                existing.AvailableBalance = account.AvailableBalance;
                existing.AlertThreshold = account.AlertThreshold;
                existing.UpdatedAt = DateTime.Now;
            }
        }
        await _context.SaveChangesAsync();
    }

    /// <summary>更新账户余额（自动计算可用余额 = 总余额 - 冻结金额）</summary>
    public async Task UpdateFundBalanceAsync(int accountId, decimal newBalance)
    {
        var account = await _context.FundAccounts.FindAsync(accountId);
        if (account != null)
        {
            var frozen = account.FrozenAmount ?? 0;
            account.TotalBalance = newBalance;
            account.AvailableBalance = Math.Max(0, newBalance - frozen);
            account.LastSyncAt = DateTime.Now;
            account.UpdatedAt = DateTime.Now;
            await _context.SaveChangesAsync();
        }
    }

    /// <summary>设置预警阈值</summary>
    public async Task UpdateFundThresholdAsync(int accountId, decimal threshold)
    {
        var account = await _context.FundAccounts.FindAsync(accountId);
        if (account != null)
        {
            account.AlertThreshold = threshold;
            account.UpdatedAt = DateTime.Now;
            await _context.SaveChangesAsync();
        }
    }

    /// <summary>启用/停用账户</summary>
    public async Task ToggleFundStatusAsync(int accountId, bool isActive)
    {
        var account = await _context.FundAccounts.FindAsync(accountId);
        if (account != null)
        {
            account.IsActive = isActive;
            account.UpdatedAt = DateTime.Now;
            await _context.SaveChangesAsync();
        }
    }

    // DTO: 资金预警行
    public class FundAlertRow
    {
        public int AccountId { get; set; }
        public string AccountName { get; set; } = "";
        public decimal AvailableBalance { get; set; }
        public decimal AlertThreshold { get; set; }
        public decimal DiffAmount { get; set; }
        public string AlertLevel { get; set; } = "正常";
    }

    // DTO: 交易在途明细行
    public class PendingOrderRow
    {
        public int OrderId { get; set; }
        public string OrderNo { get; set; } = "";
        public decimal Amount { get; set; }
        public string Method { get; set; } = "";
        public DateTime? CreatedAt { get; set; }
    }

    // DTO: 提现中明细行
    public class PendingWithdrawRow
    {
        public int RecordId { get; set; }
        public string OrderNo { get; set; } = "";
        public decimal Amount { get; set; }
        public DateTime? CreatedAt { get; set; }
    }

    // DTO: 不可用保证金明细行
    public class PendingFrozenRow
    {
        public int AccountId { get; set; }
        public string AccountName { get; set; } = "";
        public decimal FrozenAmount { get; set; }
    }

    public async Task<int> GetCostAlertCountAsync()
    {
        // 成本异动预警：ProductCosts 本月vs上月波动>5%的商品数
        var now = DateTime.Today;
        var thisMonth = new DateTime(now.Year, now.Month, 1);
        var lastMonth = thisMonth.AddMonths(-1);
        var nextMonth = thisMonth.AddMonths(1);
        
        var currentCosts = await _context.ProductCosts
            .Where(c => c.CreatedAt >= thisMonth && c.CreatedAt < nextMonth)
            .GroupBy(c => c.ProductId)
            .Select(g => new { ProductId = g.Key, Total = g.Sum(c => c.TotalCost ?? 0m) })
            .ToListAsync();
        var prevCosts = await _context.ProductCosts
            .Where(c => c.CreatedAt >= lastMonth && c.CreatedAt < thisMonth)
            .GroupBy(c => c.ProductId)
            .Select(g => new { ProductId = g.Key, Total = g.Sum(c => c.TotalCost ?? 0m) })
            .ToListAsync();
        
        var prevDict = prevCosts.ToDictionary(p => p.ProductId, p => p.Total);
        int count = 0;
        foreach (var c in currentCosts)
        {
            if (prevDict.TryGetValue(c.ProductId, out var prev) && prev > 0)
            {
                var diff = Math.Abs(c.Total - prev) / prev * 100m;
                if (diff > 5m) count++;
            }
        }
        return count;
    }

    public async Task<int> GetPendingPayableCountAsync()
    {
        return await _context.PurchaseReceipts.AsNoTracking()
            .CountAsync(r => r.Status == "Pending" || r.Status == "Partial");
    }

    public async Task<int> GetPendingReceivableCountAsync()
    {
        return await _context.Orders.AsNoTracking()
            .CountAsync(o => (o.Status == "Paid" || o.Status == "Processing") && o.ShippingStatus != "Delivered");
    }

    public async Task<decimal> GetInventoryValueAsync(string type)
    {
        return type switch
        {
            "raw" => await _context.RawMaterialInventories.AsNoTracking()
                .SumAsync(r => (decimal?)r.StockQty * (decimal?)r.UnitPrice) ?? 0,
            "note" => await _context.NoteInventories.AsNoTracking()
                .SumAsync(n => (decimal?)(n.StockQuantity ?? 0)) ?? 0,
            "product" => await _context.ProductInventories.AsNoTracking()
                .SumAsync(p => (decimal?)p.StockQty * (decimal?)p.UnitCost) ?? 0,
            _ => 0
        };
    }

    public async Task<List<MonthRevenueRow>> GetMonthlyRevenueTrendAsync(int months = 12)
    {
        var startDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-months + 1);
        var data = await _context.Orders.AsNoTracking()
            .Where(o => o.CreatedAt >= startDate
                && (new[] { "Paid", "Processing", "Shipped", "Completed" }).Contains(o.Status!))
            .GroupBy(o => new { o.CreatedAt!.Value.Year, o.CreatedAt!.Value.Month })
            .Select(g => new
            {
                g.Key.Year,
                g.Key.Month,
                Revenue = g.Sum(o => (decimal?)o.TotalAmount) ?? 0,
                Profit = g.Sum(o => (decimal?)(o.TotalAmount - o.CostAmount)) ?? 0
            })
            .ToListAsync();
        var allMonths = new List<MonthRevenueRow>();
        var cur = startDate;
        while (cur < DateTime.Today.AddMonths(1))
        {
            var y = cur.Year;
            var m = cur.Month;
            var d = data.FirstOrDefault(x => x.Year == y && x.Month == m);
            allMonths.Add(new MonthRevenueRow
            {
                Label = $"{m:D2}月",
                RevenueWan = (d?.Revenue ?? 0) / 10000m,
                ProfitWan = (d?.Profit ?? 0) / 10000m
            });
            cur = cur.AddMonths(1);
        }
        return allMonths;
    }

    public async Task<List<CategorySalesRow>> GetCategorySalesAsync()
    {
        // NOTE: Requires OrderDetails + Products + ProductTypeConfig join
        var data = await (
            from o in _context.Orders
            join od in _context.OrderDetails on o.OrderId equals od.OrderId
            join p in _context.Products on od.ProductId equals p.ProductId
            where (new[] { "Paid", "Processing", "Shipped", "Completed" }).Contains(o.Status!)
            group new { o, p } by new { p.ProductType } into g
            select new CategorySalesRow
            {
                CategoryName = g.Key.ProductType ?? "其他",
                SalesAmount = g.Sum(x => (decimal?)x.o.TotalAmount) ?? 0
            }
        ).OrderByDescending(x => x.SalesAmount).Take(8).ToListAsync();
        return data;
    }

    public async Task<List<Order>> GetRecentFinanceOrdersAsync(int limit = 10)
    {
        return await _context.Orders.AsNoTracking()
            .OrderByDescending(o => o.CreatedAt)
            .Take(limit).ToListAsync();
    }

    // ===== V19.5 成本管理 =====

    public async Task<string> GetCostMethodAsync()
    {
        var val = await _context.SiteSettings
            .Where(s => s.SettingKey == "CostCalculationMethod")
            .Select(s => s.SettingValue)
            .FirstOrDefaultAsync();
        return string.IsNullOrWhiteSpace(val) ? "FIFO" : val;
    }

    public async Task SaveCostMethodAsync(string method)
    {
        var existing = await _context.SiteSettings
            .FirstOrDefaultAsync(s => s.SettingKey == "CostCalculationMethod");
        if (existing != null) existing.SettingValue = method;
        else _context.SiteSettings.Add(new SiteSetting { SettingKey = "CostCalculationMethod", SettingValue = method });
        await _context.SaveChangesAsync();
    }

    public async Task<List<ProductCostSummary>> GetProductCostSummariesAsync(int page, int pageSize)
    {
        return await _context.Products.AsNoTracking()
            .OrderBy(p => p.ProductId)
            .Skip((page - 1) * pageSize).Take(pageSize)
            .Select(p => new ProductCostSummary
            {
                ProductId = p.ProductId,
                ProductName = p.ProductName ?? "",
                ProductType = p.ProductType ?? "standard",
                UnitCost = p.UnitCost ?? 0,
                BomCost = 0, // Product表暂无BomCost字段，从ProductCosts汇总计算
                CostItemCount = _context.ProductCosts.Count(c => c.ProductId == p.ProductId)
            })
            .ToListAsync();
    }

    public async Task<int> GetProductCountAsync()
        => await _context.Products.CountAsync();

    public async Task<List<ProductCost>> GetProductCostItemsAsync(int productId)
        => await _context.ProductCosts.AsNoTracking()
            .Where(c => c.ProductId == productId)
            .OrderBy(c => c.CostType).ThenBy(c => c.CostName)
            .ToListAsync();

    public async Task SaveProductCostsAsync(int productId, List<ProductCost> costs)
    {
        var existing = await _context.ProductCosts.Where(c => c.ProductId == productId).ToListAsync();
        _context.ProductCosts.RemoveRange(existing);
        if (costs.Any())
        {
            foreach (var c in costs) { c.ProductId = productId; c.CreatedAt = DateTime.Now; }
            _context.ProductCosts.AddRange(costs);
        }
        var totalUnit = costs.Sum(c => c.TotalCost ?? 0);
        var product = await _context.Products.FindAsync(productId);
        if (product != null) { product.UnitCost = totalUnit; }
        await _context.SaveChangesAsync();
    }

    public async Task<List<CostVarianceRow>> GetCostVarianceAsync(int page, int pageSize)
    {
        var now = DateTime.Today;
        var thisMonth = new DateTime(now.Year, now.Month, 1);
        var lastMonth = thisMonth.AddMonths(-1);
        var nextMonth = thisMonth.AddMonths(1);
        var current = await _context.ProductCosts
            .Where(c => c.CreatedAt >= thisMonth && c.CreatedAt < nextMonth)
            .GroupBy(c => c.ProductId)
            .Select(g => new { ProductId = g.Key, Total = g.Sum(c => c.TotalCost ?? 0m) })
            .ToDictionaryAsync(x => x.ProductId, x => x.Total);
        var previous = await _context.ProductCosts
            .Where(c => c.CreatedAt >= lastMonth && c.CreatedAt < thisMonth)
            .GroupBy(c => c.ProductId)
            .Select(g => new { ProductId = g.Key, Total = g.Sum(c => c.TotalCost ?? 0m) })
            .ToDictionaryAsync(x => x.ProductId, x => x.Total);
        var products = await _context.Products.AsNoTracking()
            .Where(p => current.Keys.Contains(p.ProductId))
            .ToDictionaryAsync(p => p.ProductId, p => p.ProductName ?? "");
        var result = new List<CostVarianceRow>();
        foreach (var kv in current)
        {
            var prev = previous.GetValueOrDefault(kv.Key, 0);
            var curr = kv.Value;
            var diff = curr - prev;
            var rate = prev > 0 ? (diff / prev) * 100m : 0;
            result.Add(new CostVarianceRow
            {
                ProductId = kv.Key,
                ProductName = products.GetValueOrDefault(kv.Key, "未知"),
                CurrentCost = curr,
                PreviousCost = prev,
                Variance = diff,
                VarianceRate = rate
            });
        }
        var paged = result.OrderByDescending(r => Math.Abs(r.VarianceRate))
            .Skip((page - 1) * pageSize).Take(pageSize).ToList();
        // V19.6: 加载本月归因备注（存于 SiteSettings，键 CostVarianceNote_{pid}_{yyyyMM}）
        var ym = DateTime.Today.ToString("yyyyMM");
        var noteKeys = paged.Select(r => $"CostVarianceNote_{r.ProductId}_{ym}").ToList();
        if (noteKeys.Count > 0)
        {
            var notes = await _context.SiteSettings.AsNoTracking()
                .Where(s => noteKeys.Contains(s.SettingKey))
                .ToDictionaryAsync(s => s.SettingKey, s => s.SettingValue ?? "");
            foreach (var r in paged)
                if (notes.TryGetValue($"CostVarianceNote_{r.ProductId}_{ym}", out var nv)) r.Note = nv;
        }
        return paged;
    }

    /// <summary>保存成本异动归因备注（本月）— 存 SiteSettings，对齐 V18</summary>
    public async Task SaveVarianceNoteAsync(int productId, string note)
    {
        var ym = DateTime.Today.ToString("yyyyMM");
        var key = $"CostVarianceNote_{productId}_{ym}";
        var existing = await _context.SiteSettings.FirstOrDefaultAsync(s => s.SettingKey == key);
        if (existing != null) existing.SettingValue = note ?? "";
        else _context.SiteSettings.Add(new SiteSetting { SettingKey = key, SettingValue = note ?? "" });
        await _context.SaveChangesAsync();
    }

    /// <summary>获取启用香调（成本传导链-香调成本明细用）</summary>
    public async Task<List<FragranceNote>> GetActiveFragranceNotesAsync()
        => await _context.FragranceNotes.AsNoTracking()
            .Where(n => n.IsActive == true)
            .OrderBy(n => n.NoteId)
            .ToListAsync();

    public async Task<List<PriceChangeLog>> GetPriceChangeLogsAsync(int page, int pageSize)
    {
        return await _context.PriceChangeLogs.AsNoTracking()
            .OrderByDescending(l => l.LogId)
            .Skip((page - 1) * pageSize).Take(pageSize)
            .ToListAsync();
    }

    public async Task<int> GetCostHistoryCountAsync()
        => await _context.ProductCosts.CountAsync();

    /// <summary>成本变更历史（ProductCosts 时间线，对齐 V18 “成本变更历史”）</summary>
    public async Task<List<CostHistoryRow>> GetCostChangeHistoryAsync(int page, int pageSize)
    {
        return await (from c in _context.ProductCosts.AsNoTracking()
                      join p in _context.Products.AsNoTracking() on c.ProductId equals p.ProductId into pj
                      from p in pj.DefaultIfEmpty()
                      orderby c.CreatedAt descending, c.CostId descending
                      select new CostHistoryRow
                      {
                          CostId = c.CostId,
                          ProductId = c.ProductId,
                          ProductName = p != null ? (p.ProductName ?? "") : "",
                          CostType = c.CostType,
                          CostName = c.CostName,
                          UnitCost = c.UnitCost ?? 0,
                          Quantity = c.Quantity ?? 0,
                          TotalCost = c.TotalCost ?? 0,
                          CreatedBy = c.CreatedBy,
                          CreatedAt = c.CreatedAt
                      })
                      .Skip((page - 1) * pageSize).Take(pageSize)
                      .ToListAsync();
    }

    public async Task<List<CostChainRow>> GetCostChainAsync()
    {
        var result = new List<CostChainRow>();
        var products = await _context.Products.AsNoTracking()
            .OrderBy(p => p.ProductId)
            .Select(p => new { p.ProductId, p.ProductName, p.UnitCost, p.ProductType })
            .ToListAsync();
        foreach (var p in products)
        {
            var costs = await _context.ProductCosts.AsNoTracking()
                .Where(c => c.ProductId == p.ProductId).ToListAsync();
            result.Add(new CostChainRow
            {
                ProductId = p.ProductId,
                ProductName = p.ProductName ?? "",
                ProductType = p.ProductType ?? "",
                UnitCost = p.UnitCost ?? 0,
                BomCost = 0,
                BomItems = costs.Where(c => c.CostType == "BOM").Sum(c => c.TotalCost ?? 0),
                PurchaseItems = costs.Where(c => c.CostType == "Purchase").Sum(c => c.TotalCost ?? 0),
                PackagingItems = costs.Where(c => c.CostType == "Packaging").Sum(c => c.TotalCost ?? 0),
                OtherItems = costs.Where(c => c.CostType == "Other").Sum(c => c.TotalCost ?? 0)
            });
        }
        return result;
    }

    // ===== V19.5 现金流预测 =====

    public async Task<CashFlowData> GetCashFlowDataAsync()
    {
        var thirtyDaysAgo = DateTime.Today.AddDays(-30);
        var thirtyDaysLater = DateTime.Today.AddDays(30);
        var today = DateTime.Today;

        var cashIn = await _context.PaymentRecords.AsNoTracking()
            .Where(p => p.Status == "Completed" && p.CreatedAt >= thirtyDaysAgo)
            .SumAsync(p => (decimal?)p.Amount) ?? 0;
        var cashOut = await _context.PaymentRecords.AsNoTracking()
            .Where(p => p.Status == "Completed" && p.CreatedAt >= thirtyDaysAgo)
            .SumAsync(p => (decimal?)p.Amount) ?? 0;
        // 简化：cashIn = 收款, cashOut = 付款（如PaymentRecord有PaymentType字段则区分）
        var fundBalance = await _context.FundAccounts.AsNoTracking()
            .Where(f => f.IsActive == true)
            .SumAsync(f => (decimal?)f.TotalBalance) ?? 0;
        var apBalance = await _context.AccountsPayables.AsNoTracking()
            .Where(a => a.Status == "Pending" || a.Status == "Partial")
            .SumAsync(a => (decimal?)(a.Amount - a.PaidAmount)) ?? 0;
        var arBalance = await _context.AccountsReceivables.AsNoTracking()
            .Where(a => a.Status == "Pending" || a.Status == "Partial")
            .SumAsync(a => (decimal?)(a.Amount - a.ReceivedAmount)) ?? 0;
        var apDue30 = await _context.AccountsPayables.AsNoTracking()
            .Where(a => (a.Status == "Pending" || a.Status == "Partial") && a.DueDate >= today && a.DueDate <= thirtyDaysLater)
            .SumAsync(a => (decimal?)(a.Amount - a.PaidAmount)) ?? 0;
        var arDue30 = await _context.AccountsReceivables.AsNoTracking()
            .Where(a => (a.Status == "Pending" || a.Status == "Partial") && a.DueDate >= today && a.DueDate <= thirtyDaysLater)
            .SumAsync(a => (decimal?)(a.Amount - a.ReceivedAmount)) ?? 0;
        var apOverdue = await _context.AccountsPayables.AsNoTracking()
            .Where(a => (a.Status == "Pending" || a.Status == "Partial") && a.DueDate < today)
            .SumAsync(a => (decimal?)(a.Amount - a.PaidAmount)) ?? 0;
        var arOverdue = await _context.AccountsReceivables.AsNoTracking()
            .Where(a => (a.Status == "Pending" || a.Status == "Partial") && a.DueDate < today)
            .SumAsync(a => (decimal?)(a.Amount - a.ReceivedAmount)) ?? 0;
        var projectedCash = fundBalance + arDue30 - apDue30;

        return new CashFlowData
        {
            FundBalance = fundBalance,
            CashIn = cashIn,
            CashOut = cashOut,
            NetCashFlow = cashIn - cashOut,
            ApBalance = apBalance,
            ArBalance = arBalance,
            ApDue30 = apDue30,
            ArDue30 = arDue30,
            ApOverdue = apOverdue,
            ArOverdue = arOverdue,
            ProjectedCash = projectedCash
        };
    }

    public class CashFlowData
    {
        public decimal FundBalance { get; set; }
        public decimal CashIn { get; set; }
        public decimal CashOut { get; set; }
        public decimal NetCashFlow { get; set; }
        public decimal ApBalance { get; set; }
        public decimal ArBalance { get; set; }
        public decimal ApDue30 { get; set; }
        public decimal ArDue30 { get; set; }
        public decimal ApOverdue { get; set; }
        public decimal ArOverdue { get; set; }
        public decimal ProjectedCash { get; set; }
    }

    // DTO 类型
    public class ProductCostSummary
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; } = "";
        public string ProductType { get; set; } = "";
        public decimal UnitCost { get; set; }
        public decimal BomCost { get; set; }
        public int CostItemCount { get; set; }
    }

    public class CostVarianceRow
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; } = "";
        public decimal CurrentCost { get; set; }
        public decimal PreviousCost { get; set; }
        public decimal Variance { get; set; }
        public decimal VarianceRate { get; set; }
        public string? Note { get; set; }
    }

    public class CostChainRow
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; } = "";
        public string ProductType { get; set; } = "";
        public decimal UnitCost { get; set; }
        public decimal BomCost { get; set; }
        public decimal BomItems { get; set; }
        public decimal PurchaseItems { get; set; }
        public decimal PackagingItems { get; set; }
        public decimal OtherItems { get; set; }
    }

    public class CostHistoryRow
    {
        public int CostId { get; set; }
        public int ProductId { get; set; }
        public string ProductName { get; set; } = "";
        public string? CostType { get; set; }
        public string? CostName { get; set; }
        public decimal UnitCost { get; set; }
        public double Quantity { get; set; }
        public decimal TotalCost { get; set; }
        public string? CreatedBy { get; set; }
        public DateTime? CreatedAt { get; set; }
    }

    // ===== 采购审核 (PurchaseReview) =====

    /// <summary>采购审核列表（按状态筛选，含供应商名）</summary>
    public async Task<List<PurchaseReviewRow>> GetPurchaseReviewListAsync(string status)
    {
        var q = _context.PurchaseOrders.AsNoTracking().AsQueryable();
        if (status == "pending") q = q.Where(o => o.Status == "Submitted" || o.Status == "Pending");
        else if (status == "approved") q = q.Where(o => o.Status == "FinanceApproved" || o.Status == "Approved" || o.Status == "Ordered");
        else if (status == "rejected") q = q.Where(o => o.Status == "Rejected");
        var orders = await q.OrderByDescending(o => o.PurchaseId).Take(100).ToListAsync();
        var supIds = orders.Where(o => o.SupplierId != null).Select(o => o.SupplierId!.Value).Distinct().ToList();
        var sups = await _context.Suppliers.AsNoTracking()
            .Where(s => supIds.Contains(s.SupplierId))
            .ToDictionaryAsync(s => s.SupplierId, s => s.SupplierName);
        return orders.Select(o => new PurchaseReviewRow
        {
            PurchaseId = o.PurchaseId,
            PurchaseNo = string.IsNullOrEmpty(o.PurchaseNo) ? ("PO#" + o.PurchaseId) : o.PurchaseNo!,
            SupplierName = o.SupplierId != null && sups.ContainsKey(o.SupplierId.Value) ? sups[o.SupplierId.Value] : "-",
            CategoryCode = o.CategoryCode ?? "-",
            TotalAmount = o.TotalAmount ?? 0,
            Status = o.Status ?? "",
            OrderDate = o.OrderDate ?? o.CreatedAt,
            ApprovedAt = o.ApprovedAt
        }).ToList();
    }

    /// <summary>采购审核三态计数</summary>
    public async Task<(int Pending, int Approved, int Rejected)> GetPurchaseReviewCountsAsync()
    {
        var pending = await _context.PurchaseOrders.CountAsync(o => o.Status == "Submitted" || o.Status == "Pending");
        var approved = await _context.PurchaseOrders.CountAsync(o => o.Status == "FinanceApproved" || o.Status == "Approved" || o.Status == "Ordered");
        var rejected = await _context.PurchaseOrders.CountAsync(o => o.Status == "Rejected");
        return (pending, approved, rejected);
    }

    /// <summary>审核通过：更新采购单状态 + 写入审核记录（PurchaseOrder/PurchaseCostReview 为无主键实体，用参数化原生 SQL）</summary>
    public async Task<bool> ApprovePurchaseAsync(int purchaseId, int reviewerId, decimal amount, string allocation, string comments)
    {
        var n = await _context.Database.ExecuteSqlInterpolatedAsync(
            $"UPDATE PurchaseOrders SET Status='FinanceApproved', ApprovedBy={reviewerId}, ApprovedAt=GETDATE(), UpdatedAt=GETDATE() WHERE PurchaseID={purchaseId}");
        if (n == 0) return false;
        await _context.Database.ExecuteSqlInterpolatedAsync(
            $"INSERT INTO PurchaseCostReview (PurchaseID, ReviewerID, ReviewStatus, ReviewAmount, CostAllocation, ReviewComments, ReviewedAt, CreatedAt) VALUES ({purchaseId}, {reviewerId}, 'Approved', {amount}, {allocation}, {comments}, GETDATE(), GETDATE())");
        return true;
    }

    /// <summary>审核驳回：更新采购单状态 + 写入驳回记录</summary>
    public async Task<bool> RejectPurchaseAsync(int purchaseId, int reviewerId, string reason)
    {
        var n = await _context.Database.ExecuteSqlInterpolatedAsync(
            $"UPDATE PurchaseOrders SET Status='Rejected', UpdatedAt=GETDATE() WHERE PurchaseID={purchaseId}");
        if (n == 0) return false;
        await _context.Database.ExecuteSqlInterpolatedAsync(
            $"INSERT INTO PurchaseCostReview (PurchaseID, ReviewerID, ReviewStatus, ReviewAmount, CostAllocation, ReviewComments, ReviewedAt, CreatedAt) VALUES ({purchaseId}, {reviewerId}, 'Rejected', 0, '', {reason}, GETDATE(), GETDATE())");
        return true;
    }

    public class PurchaseReviewRow
    {
        public int PurchaseId { get; set; }
        public string PurchaseNo { get; set; } = "";
        public string SupplierName { get; set; } = "";
        public string CategoryCode { get; set; } = "";
        public decimal TotalAmount { get; set; }
        public string Status { get; set; } = "";
        public DateTime? OrderDate { get; set; }
        public DateTime? ApprovedAt { get; set; }
    }

    // ===== 支付配置 (PaymentConfig) =====

    /// <summary>支付配置全部键（对齐 V18 payment_config.asp）</summary>
    public static readonly string[] PaymentConfigKeys = new[]
    {
        "EnableAlipay","EnableWechatPay","EnableBankTransfer","EnableStripe","EnableUnionPay","EnablePayPal",
        "AlipayAppId","AlipayMerchantId","AlipayFeeRate",
        "WechatAppId","WechatMchId","WechatFeeRate",
        "BankAccountName","BankAccountNo","BankName",
        "StripePublishableKey","StripeSecretKey","StripeWebhookSecret","StripeFeeRate","StripeFixedFee",
        "UnionPayMerchantId","UnionPayCertPath","UnionPayFeeRate",
        "PayPalClientId","PayPalSecret","PayPalSandbox","PayPalFeeRate","PayPalFixedFee",
        "PaymentTestMode","DefaultPaymentMethod"
    };

    /// <summary>读取支付配置（缺失键返回空串）</summary>
    public async Task<Dictionary<string, string>> GetPaymentConfigAsync()
    {
        var keys = PaymentConfigKeys.ToList();
        var settings = await _context.SiteSettings.AsNoTracking()
            .Where(s => s.SettingKey != null && keys.Contains(s.SettingKey))
            .ToListAsync();
        var dict = new Dictionary<string, string>();
        foreach (var k in PaymentConfigKeys)
            dict[k] = settings.FirstOrDefault(s => s.SettingKey == k)?.SettingValue ?? "";
        return dict;
    }

    /// <summary>保存支付配置（UPSERT SiteSettings）</summary>
    public async Task SavePaymentConfigAsync(Dictionary<string, string> config)
    {
        var keys = config.Keys.ToList();
        var existing = await _context.SiteSettings
            .Where(s => s.SettingKey != null && keys.Contains(s.SettingKey))
            .ToListAsync();
        foreach (var kv in config)
        {
            var s = existing.FirstOrDefault(x => x.SettingKey == kv.Key);
            if (s != null) { s.SettingValue = kv.Value; s.UpdatedAt = DateTime.Now; }
            else _context.SiteSettings.Add(new SiteSetting { SettingKey = kv.Key, SettingValue = kv.Value, UpdatedAt = DateTime.Now });
        }
        await _context.SaveChangesAsync();
    }

    // ===== 资金看板 (FundDashboard) 新增（唯一新增，其余复用已有） =====

    /// <summary>获取所有资金账户（有效优先）</summary>
    public async Task<List<FundAccount>> GetFundAccountsAsync()
        => await _context.FundAccounts.AsNoTracking()
            .OrderByDescending(a => a.IsActive).ThenByDescending(a => a.AccountId)
            .ToListAsync();

    public async Task<FundAccount?> GetFundAccountAsync(int id)
        => await _context.FundAccounts.FindAsync(id);

    /// <summary>待结算资金穿透聚合（3项一次返回，含合计）</summary>
    public async Task<PendingPenetrationDto> GetPendingPenetrationAsync()
    {
        var ordersAmount = await _context.Orders.AsNoTracking()
            .Where(o => o.Status == "Paid")
            .SumAsync(o => (decimal?)o.TotalAmount) ?? 0m;
        var withdrawAmount = await _context.PaymentRecords.AsNoTracking()
            .Where(p => p.Status == "Pending" && p.TransactionType == "Transfer")
            .SumAsync(p => (decimal?)(p.Amount ?? 0m)) ?? 0;
        var frozenAmount = await _context.FundAccounts.AsNoTracking()
            .Where(f => f.IsActive == true)
            .SumAsync(f => (decimal?)(f.FrozenAmount ?? 0m)) ?? 0;
        return new PendingPenetrationDto
        {
            OrdersInTransit = ordersAmount,
            PendingWithdraw = withdrawAmount,
            FrozenMargin = frozenAmount,
            Total = ordersAmount + withdrawAmount + frozenAmount
        };
    }

    /// <summary>待结算资金明细弹窗数据（统一入口，复用已有独立方法）</summary>
    public async Task<List<PendingDetailRow>> GetPendingDetailAsync(string type)
    {
        switch (type)
        {
            case "orders":
                var orders = await GetPendingOrdersDetailAsync();
                return orders.Select(o => new PendingDetailRow { Id = o.OrderId, No = o.OrderNo, Amount = o.Amount, Extra = o.Method }).ToList();
            case "withdraw":
                var wds = await GetPendingWithdrawDetailAsync();
                return wds.Select(w => new PendingDetailRow { Id = w.RecordId, No = w.OrderNo, Amount = w.Amount }).ToList();
            case "frozen":
                var fzs = await GetFrozenDetailAsync();
                return fzs.Select(f => new PendingDetailRow { Id = f.AccountId, No = f.AccountName, Amount = f.FrozenAmount }).ToList();
            default:
                return new();
        }
    }

    // DTO: 待结算聚合
    public class PendingPenetrationDto
    {
        public decimal OrdersInTransit { get; set; }
        public decimal PendingWithdraw { get; set; }
        public decimal FrozenMargin { get; set; }
        public decimal Total { get; set; }
    }

    // DTO: 明细弹窗通用行（映射已有各类型 DTO）
    public class PendingDetailRow
    {
        public int Id { get; set; }
        public string No { get; set; } = "";
        public decimal Amount { get; set; }
        public string Extra { get; set; } = "";
    }

    public class MonthRevenueRow
    {
        public string Label { get; set; } = "";
        public decimal RevenueWan { get; set; }
        public decimal ProfitWan { get; set; }
    }

    public class CategorySalesRow
    {
        public string CategoryName { get; set; } = "";
        public decimal SalesAmount { get; set; }
    }
}
