using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class FinanceController : ControllerBase
{
    private readonly FinanceRepository _repo;
    private readonly PerfumeShopContext _db;
    public FinanceController(FinanceRepository repo, PerfumeShopContext db) { _repo = repo; _db = db; }

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

    // ===== V20 扩展端点 =====

    [HttpGet("overview")]
    public async Task<IActionResult> GetOverview()
    {
        var revenue = await _repo.GetTotalRevenueAsync();
        var cost = await _repo.GetTotalCostAsync();
        var refund = await _repo.GetTotalRefundAsync();
        var pendingRecon = await _repo.GetPendingReconCountAsync();
        var alertCount = await _repo.GetReconAlertCountAsync() + await _repo.GetBudgetAlertCountAsync() + await _repo.GetFundAlertCountAsync();
        var invRaw = await _repo.GetInventoryValueAsync("raw");
        var invNote = await _repo.GetInventoryValueAsync("note");
        var invProduct = await _repo.GetInventoryValueAsync("product");
        return Ok(new { revenue, cost, refund, profit = revenue - cost - refund, pendingRecon, alertCount, inventoryValue = invRaw + invNote + invProduct });
    }

    [HttpGet("revenue-trend")]
    public async Task<IActionResult> GetRevenueTrend(int months = 12) => Ok(await _repo.GetMonthlyRevenueTrendAsync(months));

    [HttpGet("category-sales")]
    public async Task<IActionResult> GetCategorySales() => Ok(await _repo.GetCategorySalesAsync());

    [HttpGet("cash-flow")]
    public async Task<IActionResult> GetCashFlow() => Ok(await _repo.GetCashFlowDataAsync());

    [HttpGet("cost-variance")]
    public async Task<IActionResult> GetCostVariance(int page = 1, int pageSize = 20) => Ok(await _repo.GetCostVarianceAsync(page, pageSize));

    [HttpGet("cost-chain")]
    public async Task<IActionResult> GetCostChain() => Ok(await _repo.GetCostChainAsync());

    [HttpGet("cost-method")]
    public async Task<IActionResult> GetCostMethod() => Ok(await _repo.GetCostMethodAsync());

    [HttpPut("cost-method")]
    public async Task<IActionResult> SaveCostMethod([FromBody] string method) { await _repo.SaveCostMethodAsync(method); return Ok(); }

    [HttpGet("product-costs")]
    public async Task<IActionResult> GetProductCosts(int page = 1, int pageSize = 20) => Ok(await _repo.GetProductCostSummariesAsync(page, pageSize));

    [HttpGet("transactions")]
    public async Task<IActionResult> GetTransactions(int page = 1, int pageSize = 20) { var (items, total) = await _repo.GetPaymentRecordsAsync(page, pageSize); return Ok(new { items, total }); }

    [HttpGet("gl-transactions")]
    public async Task<IActionResult> GetGlTransactions(int page = 1, int pageSize = 20) { var (items, total) = await _repo.GetGlTransactionsAsync(page, pageSize); return Ok(new { items, total }); }

    [HttpGet("purchase-orders")]
    public async Task<IActionResult> GetPurchaseOrders(int page = 1, int pageSize = 20)
    {
        var items = await _db.PurchaseOrders.AsNoTracking().OrderByDescending(o => o.PurchaseId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        var total = await _db.PurchaseOrders.CountAsync();
        return Ok(new { items, total });
    }

    [HttpPut("purchase-orders/{id}/approve")]
    public async Task<IActionResult> ApprovePurchaseOrder(int id)
    {
        await _db.PurchaseOrders.Where(o => o.PurchaseId == id).ExecuteUpdateAsync(s => s.SetProperty(o => o.Status, "FinanceApproved").SetProperty(o => o.ApprovedAt, DateTime.Now));
        return Ok();
    }

    [HttpGet("fund-accounts")]
    public async Task<IActionResult> GetFundAccounts() => Ok(await _db.FundAccounts.AsNoTracking().OrderBy(a => a.AccountId).ToListAsync());

    // ===== V19.X 资金看板 CRUD =====

    [HttpPost("fund-accounts")]
    public async Task<IActionResult> SaveFundAccount([FromBody] FundAccount account)
    {
        if (string.IsNullOrWhiteSpace(account.AccountName))
            return BadRequest(new { success = false, message = "账户名称不能为空" });
        await _repo.SaveFundAccountAsync(account);
        return Ok(new { success = true, message = "保存成功" });
    }

    [HttpPut("fund-accounts/{id}/balance")]
    public async Task<IActionResult> UpdateFundBalance(int id, [FromBody] FundBalanceDto dto)
    {
        if (id <= 0) return BadRequest(new { success = false, message = "无效参数" });
        await _repo.UpdateFundBalanceAsync(id, dto.NewBalance);
        return Ok(new { success = true, message = "余额已更新" });
    }

    [HttpPut("fund-accounts/{id}/threshold")]
    public async Task<IActionResult> UpdateFundThreshold(int id, [FromBody] FundThresholdDto dto)
    {
        if (id <= 0) return BadRequest(new { success = false, message = "无效参数" });
        await _repo.UpdateFundThresholdAsync(id, dto.Threshold);
        return Ok(new { success = true, message = "阈值已设置" });
    }

    [HttpPut("fund-accounts/{id}/toggle-status")]
    public async Task<IActionResult> ToggleFundStatus(int id, [FromBody] FundStatusDto dto)
    {
        if (id <= 0) return BadRequest(new { success = false, message = "无效参数" });
        await _repo.ToggleFundStatusAsync(id, dto.IsActive);
        return Ok(new { success = true, message = dto.IsActive ? "已启用" : "已停用" });
    }

    [HttpGet("fund-accounts/pending-orders")]
    public async Task<IActionResult> GetPendingOrders() => Ok(await _repo.GetPendingOrdersDetailAsync());

    [HttpGet("fund-accounts/pending-withdraw")]
    public async Task<IActionResult> GetPendingWithdraw() => Ok(await _repo.GetPendingWithdrawDetailAsync());

    [HttpGet("fund-accounts/frozen")]
    public async Task<IActionResult> GetFrozenDetail() => Ok(await _repo.GetFrozenDetailAsync());

    [HttpGet("fund-accounts/alerts")]
    public async Task<IActionResult> GetFundAlerts() => Ok(await _repo.GetFundAlertsDetailAsync());

    [HttpGet("fund-accounts/stats")]
    public async Task<IActionResult> GetFundStats()
    {
        var accounts = await _db.FundAccounts.AsNoTracking().Where(a => a.IsActive == true).ToListAsync();
        var totalBookBalance = accounts.Sum(a => a.TotalBalance ?? 0);
        var totalAvailable = accounts.Sum(a => a.AvailableBalance ?? 0);
        var totalFrozen = accounts.Sum(a => a.FrozenAmount ?? 0);
        var totalPending = accounts.Sum(a => a.PendingSettlement ?? 0);
        var pendingOrders = await _repo.GetPendingOrdersAmountAsync();
        var pendingWithdraw = await _repo.GetPendingWithdrawAmountAsync();
        var largeTransferCount = await _repo.GetLargeTransferCountAsync();
        return Ok(new { totalBookBalance, totalAvailable, totalFrozen, totalPending, pendingOrders, pendingWithdraw, largeTransferCount, accountCount = accounts.Count });
    }

    [HttpGet("risk-control")]
    public async Task<IActionResult> GetRiskControl()
    {
        var abnormalOrders = await _db.Orders.CountAsync(o => o.Status == "Cancelled" || o.Status == "Refunded");
        var ipBlacklist = await _db.Ipblacklists.CountAsync();
        var loginAlerts = await _db.LoginAlerts.CountAsync();
        var refundAlerts = await _db.RefundRecords.CountAsync(r => r.Status == "Pending");
        return Ok(new { abnormalOrders, ipBlacklist, loginAlerts, refundAlerts });
    }

    [HttpGet("price-change-logs")]
    public async Task<IActionResult> GetPriceChangeLogs(int page = 1, int pageSize = 20) => Ok(await _repo.GetPriceChangeLogsAsync(page, pageSize));
}

// 资金看板 DTO
public class FundBalanceDto { public decimal NewBalance { get; set; } }
public class FundThresholdDto { public decimal Threshold { get; set; } }
public class FundStatusDto { public bool IsActive { get; set; } }
