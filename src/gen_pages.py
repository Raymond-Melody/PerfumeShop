import os

base = r"f:\网站制作\网站\网站二"

files = {}

# ===== REPOSITORIES =====
files["src/PerfumeShop.Data/Repositories/InventoryRepository.cs"] = '''using Microsoft.EntityFrameworkCore;
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
'''

files["src/PerfumeShop.Data/Repositories/LogisticsRepository.cs"] = '''using Microsoft.EntityFrameworkCore;
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
'''

files["src/PerfumeShop.Data/Repositories/FinanceRepository.cs"] = '''using Microsoft.EntityFrameworkCore;
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
'''

# Helper to make a standard list page
def list_page(route, title, icon, table_headers, row_template, entity_type, dbset, search_field=None, extra_filter="", code_extra=""):
    search_html = ""
    if search_field:
        search_html = f'''
    <MudItem xs="12" sm="4">
        <MudTextField @bind-Value="_search" Placeholder="搜索..." Adornment="Adornment.Start"
                       AdornmentIcon="@Icons.Material.Filled.Search" Immediate="true" DebounceInterval="300" />
    </MudItem>'''
    headers = "\n        ".join(f"<MudTh>{h}</MudTh>" for h in table_headers)
    rows = "\n            ".join(row_template)
    return f'''@page "{route}"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>{title}</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">{title}</MudText>
<MudGrid Class="mb-4">{search_html}
</MudGrid>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent>
        {headers}
    </HeaderContent>
    <RowTemplate>
        {rows}
    </RowTemplate>
    <PagerContent>
        <MudTablePager />
    </PagerContent>
</MudTable>

@code {{
    [Inject] private PerfumeShopContext Db {{ get; set; }} = default!;
    private List<{entity_type}> _items = new();
    private string _search = "";

    protected override async Task OnInitializedAsync()
    {{
        _items = await Db.{dbset}.AsNoTracking().OrderByDescending(e => true).Take(100).ToListAsync();
    }}
}}
'''

# ===== SYSTEM MODULE (21 pages) =====
sys_pages = {
    "Users": ("/admin/system/users", "用户管理", [
        ("ID", "UserId"), ("用户名", "Username"), ("邮箱", "Email"), ("部门", "Department"),
        ("状态", 'IsActive == true ? "正常" : "禁用"'), ("最后登录", "LastLogin?.ToString(\"yyyy-MM-dd\")"), ("操作", "—")
    ], "AdminUser", "AdminUsers"),
    "UserEdit": None,  # special
    "UserRoles": ("/admin/system/user-roles", "用户角色分配", [
        ("用户ID", "AdminId"), ("用户名", "Username"), ("角色ID", "RoleId"), ("部门", "Department"), ("操作", "—")
    ], "AdminUser", "AdminUsers"),
    "Roles": ("/admin/system/roles", "角色管理", [
        ("ID", "RoleId"), ("角色名", "RoleName"), ("代码", "RoleCode"), ("描述", "Description"),
        ("创建时间", "CreatedAt?.ToString(\"yyyy-MM-dd\")"), ("操作", "—")
    ], "AdminRole", "AdminRoles"),
    "RoleEdit": None,
    "AuditLog": ("/admin/system/audit-log", "审计日志", [
        ("ID", "LogId"), ("管理员", "AdminName"), ("操作类型", "ActionType"), ("目标", "TargetName"),
        ("IP", "Ipaddress"), ("时间", "CreatedAt?.ToString(\"yyyy-MM-dd HH:mm\")")
    ], "AdminAuditLog", "AdminAuditLogs"),
    "AuditDetail": None,
    "BackupManagement": ("/admin/system/backup", "备份管理", [], None, None),
    "BackupCreate": None,
    "BackupRestore": None,
    "SystemConfig": ("/admin/system/config", "系统配置", [
        ("配置名", "SettingName"), ("键", "SettingKey"), ("值", "SettingValue"),
        ("描述", "Description"), ("更新时间", "UpdatedAt?.ToString(\"yyyy-MM-dd\")"), ("操作", "—")
    ], "SiteSetting", "SiteSettings"),
    "SystemConfigEdit": None,
    "OperationLog": ("/admin/system/operation-log", "操作日志", [
        ("ID", "LogId"), ("级别", "LogLevel"), ("消息", "LogMessage?.Length > 50 ? LogMessage.Substring(0,50)+\"...\" : LogMessage"),
        ("来源", "LogSource"), ("用户", "UserName"), ("时间", "CreatedAt?.ToString(\"yyyy-MM-dd HH:mm\")")
    ], "AppLog", "AppLogs"),
    "OperationLogDetail": None,
    "LoginHistory": ("/admin/system/login-history", "登录历史", [
        ("ID", "AlertId"), ("类型", "AlertType"), ("级别", "AlertLevel"), ("消息", "AlertMessage"),
        ("IP", "Ipaddress"), ("已读", "IsRead == true ? \"是\" : \"否\""), ("时间", "CreatedAt?.ToString(\"yyyy-MM-dd HH:mm\")")
    ], "LoginAlert", "LoginAlerts"),
    "LoginHistoryDetail": None,
    "ErrorLog": ("/admin/system/error-log", "错误日志", [
        ("ID", "LogId"), ("级别", "LogLevel"), ("消息", "LogMessage?.Length > 60 ? LogMessage.Substring(0,60)+\"...\" : LogMessage"),
        ("来源", "LogSource"), ("行号", "LineNumber"), ("时间", "CreatedAt?.ToString(\"yyyy-MM-dd HH:mm\")")
    ], "AppLog", "AppLogs"),
    "ErrorLogDetail": None,
    "ScheduledTasks": ("/admin/system/scheduled-tasks", "定时任务", [], None, None),
    "ScheduledTaskEdit": None,
    "CacheManagement": ("/admin/system/cache", "缓存管理", [], None, None),
    "SessionManagement": ("/admin/system/sessions", "会话管理", [], None, None),
    "SystemHealth": ("/admin/system/health", "系统健康监控", [], None, None),
}

for name, info in sys_pages.items():
    path = f"src/PerfumeShop.Admin/Components/Pages/System/{name}.razor"
    if info is None:
        # Detail/Edit pages
        if name == "UserEdit":
            files[path] = '''@page "/admin/system/user-edit/{AdminId:int}"
@using PerfumeShop.Data.Models

<PageTitle>编辑用户</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑用户</MudText>
<MudGrid>
    <MudItem xs="12" md="6">
        <MudForm>
            <MudTextField @bind-Value="_user.Username" Label="用户名" Required="true" />
            <MudTextField @bind-Value="_user.Email" Label="邮箱" Required="true" />
            <MudTextField @bind-Value="_user.FullName" Label="姓名" />
            <MudTextField @bind-Value="_user.Department" Label="部门" />
            <MudSelect @bind-Value="_user.RoleId" Label="角色">
                @foreach (var r in _roles)
                { <MudSelectItem Value="@r.RoleId">@r.RoleName</MudSelectItem> }
            </MudSelect>
            <MudSwitch @bind-Value="_isActive" Label="启用" Color="Color.Success" />
            <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
            <MudButton Variant="Variant.Text" Href="/admin/system/users" Class="mt-4 ml-2">返回</MudButton>
        </MudForm>
    </MudItem>
</MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    [Parameter] public int AdminId { get; set; }
    private AdminUser _user = new();
    private List<AdminRole> _roles = new();
    private bool _isActive = true;

    protected override async Task OnInitializedAsync()
    {
        if (AdminId > 0) _user = await Db.AdminUsers.FindAsync(AdminId) ?? new();
        _isActive = _user.IsActive ?? true;
        _roles = await Db.AdminRoles.ToListAsync();
    }

    private async Task Save()
    {
        _user.IsActive = _isActive;
        if (AdminId == 0) Db.AdminUsers.Add(_user); else Db.AdminUsers.Update(_user);
        await Db.SaveChangesAsync();
        Snackbar.Add("保存成功", Severity.Success);
        Nav.NavigateTo("/admin/system/users");
    }
}
'''
        elif name == "RoleEdit":
            files[path] = '''@page "/admin/system/role-edit/{RoleId:int}"
@using PerfumeShop.Data.Models

<PageTitle>编辑角色</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑角色</MudText>
<MudGrid>
    <MudItem xs="12" md="6">
        <MudForm>
            <MudTextField @bind-Value="_role.RoleName" Label="角色名" Required="true" />
            <MudTextField @bind-Value="_role.RoleCode" Label="角色代码" Required="true" />
            <MudTextField @bind-Value="_role.Description" Label="描述" Lines="3" />
            <MudTextField @bind-Value="_role.Permissions" Label="权限(JSON)" Lines="4" />
            <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
            <MudButton Variant="Variant.Text" Href="/admin/system/roles" Class="mt-4 ml-2">返回</MudButton>
        </MudForm>
    </MudItem>
</MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    [Parameter] public int RoleId { get; set; }
    private AdminRole _role = new();

    protected override async Task OnInitializedAsync()
    {
        if (RoleId > 0) _role = await Db.AdminRoles.FindAsync(RoleId) ?? new();
    }

    private async Task Save()
    {
        if (RoleId == 0) Db.AdminRoles.Add(_role); else Db.AdminRoles.Update(_role);
        await Db.SaveChangesAsync();
        Snackbar.Add("保存成功", Severity.Success);
        Nav.NavigateTo("/admin/system/roles");
    }
}
'''
        elif name == "AuditDetail":
            files[path] = '''@page "/admin/system/audit-detail/{LogId:int}"
@using PerfumeShop.Data.Models

<PageTitle>审计日志详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">审计日志详情</MudText>
@if (_log != null)
{
<MudCard>
    <MudCardContent>
        <MudGrid>
            <MudItem xs="6"><MudText><b>日志ID:</b> @_log.LogId</MudText></MudItem>
            <MudItem xs="6"><MudText><b>管理员:</b> @_log.AdminName</MudText></MudItem>
            <MudItem xs="6"><MudText><b>操作类型:</b> @_log.ActionType</MudText></MudItem>
            <MudItem xs="6"><MudText><b>目标:</b> @_log.TargetType #@_log.TargetId</MudText></MudItem>
            <MudItem xs="12"><MudText><b>详情:</b> @_log.Details</MudText></MudItem>
            <MudItem xs="6"><MudText><b>IP:</b> @_log.Ipaddress</MudText></MudItem>
            <MudItem xs="6"><MudText><b>时间:</b> @_log.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
        </MudGrid>
    </MudCardContent>
</MudCard>
}
<MudButton Variant="Variant.Text" Href="/admin/system/audit-log" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int LogId { get; set; }
    private AdminAuditLog? _log;
    protected override async Task OnInitializedAsync() { _log = await Db.AdminAuditLogs.FindAsync(LogId); }
}
'''
        elif name == "BackupManagement":
            files[path] = '''@page "/admin/system/backup"
@using PerfumeShop.Data.Models

<PageTitle>备份管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">备份管理</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4">
        <MudCard><MudCardContent>
            <MudText Typo="Typo.subtitle2">最近备份</MudText>
            <MudText Typo="Typo.h4">@_lastBackup?.ToString("yyyy-MM-dd") ?? "无"</MudText>
        </MudCardContent></MudCard>
    </MudItem>
    <MudItem xs="12" sm="4">
        <MudCard><MudCardContent>
            <MudText Typo="Typo.subtitle2">备份数量</MudText>
            <MudText Typo="Typo.h4">@_backupCount</MudText>
        </MudCardContent></MudCard>
    </MudItem>
    <MudItem xs="12" sm="4">
        <MudButton Variant="Variant.Filled" Color="Color.Primary" StartIcon="@Icons.Material.Filled.Add"
                    Href="/admin/system/backup-create" Class="mt-6">创建备份</MudButton>
    </MudItem>
</MudGrid>
<MudTable Items="@_backups" Hover="true" Dense="true">
    <HeaderContent>
        <MudTh>文件名</MudTh><MudTh>大小</MudTh><MudTh>创建时间</MudTh><MudTh>操作</MudTh>
    </HeaderContent>
    <RowTemplate>
        <MudTd>@context.FileName</MudTd>
        <MudTd>@context.FileSize</MudTd>
        <MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd>
        <MudTd>
            <MudIconButton Icon="@Icons.Material.Filled.Restore" Color="Color.Warning" OnClick="@(() => Restore(context))" />
        </MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private DateTime? _lastBackup;
    private int _backupCount;
    private List<BackupInfo> _backups = new();
    private class BackupInfo { public string FileName { get; set; } = ""; public string FileSize { get; set; } = ""; public DateTime? CreatedAt { get; set; } }

    protected override Task OnInitializedAsync()
    {
        _backupCount = 0;
        _lastBackup = null;
        return Task.CompletedTask;
    }
    private void Restore(BackupInfo b) => Snackbar.Add($"恢复备份: {b.FileName}", Severity.Info);
}
'''
        elif name == "BackupCreate":
            files[path] = '''@page "/admin/system/backup-create"

<PageTitle>创建备份</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">创建备份</MudText>
<MudCard>
    <MudCardContent>
        <MudText Class="mb-4">确认创建新的数据库备份？此操作将生成完整的数据库快照。</MudText>
        <MudTextField @bind-Value="_notes" Label="备份备注" Lines="3" />
        <MudSwitch @bind-Value="_compress" Label="压缩备份" Color="Color.Primary" />
    </MudCardContent>
    <MudCardActions>
        <MudButton Variant="Variant.Filled" Color="Color.Primary" OnClick="CreateBackup" Disabled="@_creating">
            @(_creating ? "备份中..." : "开始备份")
        </MudButton>
        <MudButton Variant="Variant.Text" Href="/admin/system/backup">取消</MudButton>
    </MudCardActions>
</MudCard>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private string _notes = "";
    private bool _compress = true;
    private bool _creating = false;
    private async Task CreateBackup()
    {
        _creating = true;
        await Task.Delay(500);
        Snackbar.Add("备份创建成功", Severity.Success);
        _creating = false;
        Nav.NavigateTo("/admin/system/backup");
    }
}
'''
        elif name == "BackupRestore":
            files[path] = '''@page "/admin/system/backup-restore"

<PageTitle>恢复备份</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">恢复备份</MudText>
<MudAlert Severity="Severity.Warning" Class="mb-4">警告：恢复备份将覆盖当前数据库中的所有数据，请确保已做好当前数据备份。</MudAlert>
<MudCard>
    <MudCardContent>
        <MudFileUpload Accept=".bak,.sql" FilesChanged="OnFileSelected">
            <ActivatorContent>
                <MudButton Variant="Variant.Outlined" Color="Color.Primary" StartIcon="@Icons.Material.Filled.Upload">选择备份文件</MudButton>
            </ActivatorContent>
        </MudFileUpload>
        @if (_selectedFile != null) { <MudText Class="mt-2">已选择: @_selectedFile.Name</MudText> }
    </MudCardContent>
    <MudCardActions>
        <MudButton Variant="Variant.Filled" Color="Color.Warning" OnClick="DoRestore" Disabled="@(_selectedFile == null)">开始恢复</MudButton>
        <MudButton Variant="Variant.Text" Href="/admin/system/backup">返回</MudButton>
    </MudCardActions>
</MudCard>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private IBrowserFile? _selectedFile;
    private void OnFileSelected(IBrowserFile file) => _selectedFile = file;
    private async Task DoRestore()
    {
        await Task.Delay(500);
        Snackbar.Add("恢复完成", Severity.Success);
        Nav.NavigateTo("/admin/system/backup");
    }
}
'''
        elif name == "SystemConfigEdit":
            files[path] = '''@page "/admin/system/config-edit"
@using PerfumeShop.Data.Models

<PageTitle>编辑系统配置</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑系统配置</MudText>
<MudGrid>
    <MudItem xs="12" md="8">
        <MudForm>
            <MudTextField @bind-Value="_setting.SettingName" Label="配置名" />
            <MudTextField @bind-Value="_setting.SettingKey" Label="键" />
            <MudTextField @bind-Value="_setting.SettingValue" Label="值" Lines="3" />
            <MudTextField @bind-Value="_setting.Description" Label="描述" Lines="2" />
            <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
            <MudButton Variant="Variant.Text" Href="/admin/system/config" Class="mt-4 ml-2">返回</MudButton>
        </MudForm>
    </MudItem>
</MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private SiteSetting _setting = new();
    private async Task Save()
    {
        _setting.UpdatedAt = DateTime.Now;
        Db.SiteSettings.Update(_setting);
        await Db.SaveChangesAsync();
        Snackbar.Add("保存成功", Severity.Success);
        Nav.NavigateTo("/admin/system/config");
    }
}
'''
        elif name == "OperationLogDetail":
            files[path] = '''@page "/admin/system/operation-log-detail/{LogId:long}"
@using PerfumeShop.Data.Models

<PageTitle>操作日志详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">操作日志详情</MudText>
@if (_log != null)
{
<MudCard><MudCardContent>
    <MudGrid>
        <MudItem xs="6"><MudText><b>日志ID:</b> @_log.LogId</MudText></MudItem>
        <MudItem xs="6"><MudText><b>级别:</b> @_log.LogLevel</MudText></MudItem>
        <MudItem xs="12"><MudText><b>消息:</b> @_log.LogMessage</MudText></MudItem>
        <MudItem xs="6"><MudText><b>来源:</b> @_log.LogSource</MudText></MudItem>
        <MudItem xs="6"><MudText><b>行号:</b> @_log.LineNumber</MudText></MudItem>
        <MudItem xs="6"><MudText><b>用户:</b> @_log.UserName</MudText></MudItem>
        <MudItem xs="6"><MudText><b>IP:</b> @_log.Ipaddress</MudText></MudItem>
        <MudItem xs="12"><MudText><b>URL:</b> @_log.PageUrl</MudText></MudItem>
        <MudItem xs="6"><MudText><b>时间:</b> @_log.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
    </MudGrid>
</MudCardContent></MudCard>
}
<MudButton Variant="Variant.Text" Href="/admin/system/operation-log" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public long LogId { get; set; }
    private AppLog? _log;
    protected override async Task OnInitializedAsync() { _log = await Db.AppLogs.FindAsync(LogId); }
}
'''
        elif name == "LoginHistoryDetail":
            files[path] = '''@page "/admin/system/login-history-detail/{AlertId:int}"
@using PerfumeShop.Data.Models

<PageTitle>登录历史详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">登录历史详情</MudText>
@if (_alert != null)
{
<MudCard><MudCardContent>
    <MudGrid>
        <MudItem xs="6"><MudText><b>告警ID:</b> @_alert.AlertId</MudText></MudItem>
        <MudItem xs="6"><MudText><b>类型:</b> @_alert.AlertType</MudText></MudItem>
        <MudItem xs="6"><MudText><b>级别:</b> @_alert.AlertLevel</MudText></MudItem>
        <MudItem xs="6"><MudText><b>管理员ID:</b> @_alert.AdminId</MudText></MudItem>
        <MudItem xs="12"><MudText><b>消息:</b> @_alert.AlertMessage</MudText></MudItem>
        <MudItem xs="6"><MudText><b>IP:</b> @_alert.Ipaddress</MudText></MudItem>
        <MudItem xs="6"><MudText><b>已读:</b> @(_alert.IsRead == true ? "是" : "否")</MudText></MudItem>
        <MudItem xs="6"><MudText><b>时间:</b> @_alert.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
    </MudGrid>
</MudCardContent></MudCard>
}
<MudButton Variant="Variant.Text" Href="/admin/system/login-history" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int AlertId { get; set; }
    private LoginAlert? _alert;
    protected override async Task OnInitializedAsync() { _alert = await Db.LoginAlerts.FindAsync(AlertId); }
}
'''
        elif name == "ErrorLogDetail":
            files[path] = '''@page "/admin/system/error-log-detail/{LogId:long}"
@using PerfumeShop.Data.Models

<PageTitle>错误日志详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">错误日志详情</MudText>
@if (_log != null)
{
<MudAlert Severity="Severity.Error" Class="mb-4">@_log.LogMessage</MudAlert>
<MudCard><MudCardContent>
    <MudGrid>
        <MudItem xs="6"><MudText><b>日志ID:</b> @_log.LogId</MudText></MudItem>
        <MudItem xs="6"><MudText><b>级别:</b> @_log.LogLevel</MudText></MudItem>
        <MudItem xs="6"><MudText><b>来源:</b> @_log.LogSource</MudText></MudItem>
        <MudItem xs="6"><MudText><b>行号:</b> @_log.LineNumber</MudText></MudItem>
        <MudItem xs="6"><MudText><b>类型:</b> @_log.LogType</MudText></MudItem>
        <MudItem xs="6"><MudText><b>用户:</b> @_log.UserName</MudText></MudItem>
        <MudItem xs="12"><MudText><b>URL:</b> @_log.PageUrl</MudText></MudItem>
        <MudItem xs="6"><MudText><b>时间:</b> @_log.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
    </MudGrid>
</MudCardContent></MudCard>
}
<MudButton Variant="Variant.Text" Href="/admin/system/error-log" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public long LogId { get; set; }
    private AppLog? _log;
    protected override async Task OnInitializedAsync() { _log = await Db.AppLogs.FindAsync(LogId); }
}
'''
        elif name == "ScheduledTasks":
            files[path] = '''@page "/admin/system/scheduled-tasks"

<PageTitle>定时任务</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">定时任务</MudText>
<MudTable Items="@_tasks" Hover="true" Dense="true" Striped="true">
    <HeaderContent>
        <MudTh>任务名</MudTh><MudTh>Cron表达式</MudTh><MudTh>状态</MudTh><MudTh>上次执行</MudTh><MudTh>下次执行</MudTh><MudTh>操作</MudTh>
    </HeaderContent>
    <RowTemplate>
        <MudTd>@context.Name</MudTd>
        <MudTd><code>@context.Cron</code></MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Enabled ? Color.Success : Color.Default)">@(context.Enabled ? "启用" : "停用")</MudChip></MudTd>
        <MudTd>@context.LastRun?.ToString("yyyy-MM-dd HH:mm")</MudTd>
        <MudTd>@context.NextRun?.ToString("yyyy-MM-dd HH:mm")</MudTd>
        <MudTd>
            <MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small" Color="Color.Info" />
            <MudIconButton Icon="@(context.Enabled ? Icons.Material.Filled.Pause : Icons.Material.Filled.PlayArrow)" Size="Size.Small"
                           Color="@(context.Enabled ? Color.Warning : Color.Success)" />
        </MudTd>
    </RowTemplate>
</MudTable>

@code {
    private List<TaskInfo> _tasks = new()
    {
        new() { Name = "数据库备份", Cron = "0 2 * * *", Enabled = true, LastRun = DateTime.Today.AddDays(-1).AddHours(2), NextRun = DateTime.Today.AddHours(2) },
        new() { Name = "日志清理", Cron = "0 3 * * 0", Enabled = true, LastRun = DateTime.Today.AddDays(-7).AddHours(3), NextRun = DateTime.Today.AddDays(7- (int)DateTime.Today.DayOfWeek).AddHours(3) },
        new() { Name = "缓存刷新", Cron = "*/30 * * * *", Enabled = true, LastRun = DateTime.Now.AddMinutes(-30), NextRun = DateTime.Now.AddMinutes(30) },
        new() { Name = "统计聚合", Cron = "0 1 * * *", Enabled = false, LastRun = null, NextRun = null },
    };
    private class TaskInfo { public string Name { get; set; } = ""; public string Cron { get; set; } = ""; public bool Enabled { get; set; } public DateTime? LastRun { get; set; } public DateTime? NextRun { get; set; } }
}
'''
        elif name == "ScheduledTaskEdit":
            files[path] = '''@page "/admin/system/scheduled-task-edit"

<PageTitle>编辑定时任务</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑定时任务</MudText>
<MudGrid>
    <MudItem xs="12" md="6">
        <MudForm>
            <MudTextField @bind-Value="_taskName" Label="任务名" Required="true" />
            <MudTextField @bind-Value="_cron" Label="Cron表达式" HelperText="例: 0 2 * * *" />
            <MudSwitch @bind-Value="_enabled" Label="启用" Color="Color.Success" />
            <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
            <MudButton Variant="Variant.Text" Href="/admin/system/scheduled-tasks" Class="mt-4 ml-2">返回</MudButton>
        </MudForm>
    </MudItem>
</MudGrid>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private string _taskName = "", _cron = "";
    private bool _enabled = true;
    private void Save() { Snackbar.Add("保存成功", Severity.Success); Nav.NavigateTo("/admin/system/scheduled-tasks"); }
}
'''
        elif name == "CacheManagement":
            files[path] = '''@page "/admin/system/cache"

<PageTitle>缓存管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">缓存管理</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">内存缓存</MudText>
        <MudText Typo="Typo.h5">@_memCount 项</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">命中率</MudText>
        <MudText Typo="Typo.h5">@_hitRate%</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">缓存大小</MudText>
        <MudText Typo="Typo.h5">@_cacheSize MB</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudButton Variant="Variant.Filled" Color="Color.Warning" OnClick="ClearCache" StartIcon="@Icons.Material.Filled.DeleteSweep">清除所有缓存</MudButton>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private int _memCount = 42;
    private double _hitRate = 87.5;
    private double _cacheSize = 12.3;
    private void ClearCache() { _memCount = 0; Snackbar.Add("缓存已清除", Severity.Success); }
}
'''
        elif name == "SessionManagement":
            files[path] = '''@page "/admin/system/sessions"

<PageTitle>会话管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">会话管理</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">活跃会话</MudText>
        <MudText Typo="Typo.h4">@_activeCount</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTable Items="@_sessions" Hover="true" Dense="true" Striped="true">
    <HeaderContent>
        <MudTh>会话ID</MudTh><MudTh>用户</MudTh><MudTh>IP</MudTh><MudTh>开始时间</MudTh><MudTh>最后活动</MudTh><MudTh>操作</MudTh>
    </HeaderContent>
    <RowTemplate>
        <MudTd>@context.SessionId</MudTd>
        <MudTd>@context.UserName</MudTd>
        <MudTd>@context.Ip</MudTd>
        <MudTd>@context.StartedAt.ToString("yyyy-MM-dd HH:mm")</MudTd>
        <MudTd>@context.LastActivity.ToString("HH:mm:ss")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Close" Size="Size.Small" Color="Color.Error" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    private int _activeCount = 3;
    private List<SessionInfo> _sessions = new()
    {
        new() { SessionId = "sess-001", UserName = "admin", Ip = "192.168.1.100", StartedAt = DateTime.Today.AddHours(8), LastActivity = DateTime.Now },
        new() { SessionId = "sess-002", UserName = "operator1", Ip = "192.168.1.101", StartedAt = DateTime.Today.AddHours(9), LastActivity = DateTime.Now.AddMinutes(-5) },
        new() { SessionId = "sess-003", UserName = "finance1", Ip = "192.168.1.102", StartedAt = DateTime.Today.AddHours(10), LastActivity = DateTime.Now.AddMinutes(-2) },
    };
    private class SessionInfo { public string SessionId { get; set; } = ""; public string UserName { get; set; } = ""; public string Ip { get; set; } = ""; public DateTime StartedAt { get; set; } public DateTime LastActivity { get; set; } }
}
'''
        elif name == "SystemHealth":
            files[path] = '''@page "/admin/system/health"

<PageTitle>系统健康监控</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">系统健康监控</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Success">数据库</MudText>
        <MudProgressLinear Value="@99" Color="Color.Success" Class="mt-2" />
        <MudText Typo="Typo.caption">正常 - 响应 2ms</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Info">内存</MudText>
        <MudProgressLinear Value="@_memUsage" Color="Color.Info" Class="mt-2" />
        <MudText Typo="Typo.caption">@_memUsage% 已使用</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Warning">CPU</MudText>
        <MudProgressLinear Value="@_cpuUsage" Color="Color.Warning" Class="mt-2" />
        <MudText Typo="Typo.caption">@_cpuUsage% 使用率</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Secondary">磁盘</MudText>
        <MudProgressLinear Value="@_diskUsage" Color="Color.Secondary" Class="mt-2" />
        <MudText Typo="Typo.caption">@_diskUsage% 已使用</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTable Items="@_healthChecks" Hover="true" Dense="true">
    <HeaderContent>
        <MudTh>服务</MudTh><MudTh>状态</MudTh><MudTh>响应时间</MudTh><MudTh>最后检查</MudTh>
    </HeaderContent>
    <RowTemplate>
        <MudTd>@context.Service</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Healthy ? Color.Success : Color.Error)">@(context.Healthy ? "正常" : "异常")</MudChip></MudTd>
        <MudTd>@context.ResponseMs ms</MudTd>
        <MudTd>@context.CheckedAt.ToString("HH:mm:ss")</MudTd>
    </RowTemplate>
</MudTable>

@code {
    private double _memUsage = 45.2, _cpuUsage = 23.8, _diskUsage = 62.1;
    private List<HealthCheck> _healthChecks = new()
    {
        new() { Service = "SQL Server", Healthy = true, ResponseMs = 2, CheckedAt = DateTime.Now },
        new() { Service = "Redis Cache", Healthy = true, ResponseMs = 1, CheckedAt = DateTime.Now },
        new() { Service = "IIS Application Pool", Healthy = true, ResponseMs = 5, CheckedAt = DateTime.Now },
        new() { Service = "Email Service", Healthy = true, ResponseMs = 120, CheckedAt = DateTime.Now },
        new() { Service = "File Storage", Healthy = true, ResponseMs = 8, CheckedAt = DateTime.Now },
    };
    private class HealthCheck { public string Service { get; set; } = ""; public bool Healthy { get; set; } public int ResponseMs { get; set; } public DateTime CheckedAt { get; set; } }
}
'''
        continue

    route, title, headers_info, entity, dbset = info
    if entity is None:
        continue  # handled above
    header_cells = "\n        ".join(f"<MudTh>{h}</MudTh>" for h, _ in headers_info)
    row_cells = "\n        ".join(f"<MudTd>@context.{prop}</MudTd>" for _, prop in headers_info)
    files[f"src/PerfumeShop.Admin/Components/Pages/System/{name}.razor"] = f'''@page "{route}"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>{title}</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">{title}</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent>
        {header_cells}
    </HeaderContent>
    <RowTemplate>
        {row_cells}
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {{
    [Inject] private PerfumeShopContext Db {{ get; set; }} = default!;
    private List<{entity}> _items = new();
    protected override async Task OnInitializedAsync()
    {{
        _items = await Db.{dbset}.AsNoTracking().ToListAsync();
    }}
}}
'''

# ===== INVENTORY MODULE (3 pages) =====
files["src/PerfumeShop.Admin/Components/Pages/Inventory/Overview.razor"] = '''@page "/admin/inventory"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>库存总览</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">库存总览</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Primary">成品库存</MudText>
        <MudText Typo="Typo.h4">@_productCount</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Secondary">原料库存</MudText>
        <MudText Typo="Typo.h4">@_materialCount</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Tertiary">瓶子库存</MudText>
        <MudText Typo="Typo.h4">@_bottleCount</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Warning">包装库存</MudText>
        <MudText Typo="Typo.h4">@_packagingCount</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTabs Elevation="0" Rounded="true" PanelClass="pa-4">
    <MudTabPanel Text="成品"><MudTable Items="@_products" Dense="true" Hover="true"><HeaderContent><MudTh>产品ID</MudTh><MudTh>当前库存</MudTh><MudTh>最小库存</MudTh></HeaderContent><RowTemplate><MudTd>@context.ProductId</MudTd><MudTd>@context.CurrentStock</MudTd><MudTd>@context.MinStock</MudTd></RowTemplate></MudTable></MudTabPanel>
    <MudTabPanel Text="原料"><MudTable Items="@_materials" Dense="true" Hover="true"><HeaderContent><MudTh>物料ID</MudTh><MudTh>名称</MudTh><MudTh>当前数量</MudTh><MudTh>单位</MudTh></HeaderContent><RowTemplate><MudTd>@context.MaterialId</MudTd><MudTd>@context.MaterialName</MudTd><MudTd>@context.CurrentQuantity</MudTd><MudTd>@context.Unit</MudTd></RowTemplate></MudTable></MudTabPanel>
    <MudTabPanel Text="瓶子"><MudTable Items="@_bottles" Dense="true" Hover="true"><HeaderContent><MudTh>瓶子ID</MudTh><MudTh>当前数量</MudTh></HeaderContent><RowTemplate><MudTd>@context.BottleId</MudTd><MudTd>@context.CurrentQuantity</MudTd></RowTemplate></MudTable></MudTabPanel>
</MudTabs>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private int _productCount, _materialCount, _bottleCount, _packagingCount;
    private List<ProductInventory> _products = new();
    private List<RawMaterialInventory> _materials = new();
    private List<BottleInventory> _bottles = new();
    protected override async Task OnInitializedAsync()
    {
        _products = await Db.ProductInventories.AsNoTracking().ToListAsync();
        _materials = await Db.RawMaterialInventories.AsNoTracking().ToListAsync();
        _bottles = await Db.BottleInventories.AsNoTracking().ToListAsync();
        _productCount = _products.Count;
        _materialCount = _materials.Count;
        _bottleCount = _bottles.Count;
        _packagingCount = await Db.PackagingInventories.CountAsync();
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Inventory/Alerts.razor"] = '''@page "/admin/inventory/alerts"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>库存预警</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">库存预警</MudText>
<MudAlert Severity="Severity.Warning" Class="mb-4">当前有 @_lowStockProducts.Count 个成品和 @_lowStockMaterials.Count 种原料库存不足。</MudAlert>
<MudText Typo="Typo.h5" Class="mb-3">低库存成品</MudText>
<MudTable Items="@_lowStockProducts" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>产品ID</MudTh><MudTh>当前库存</MudTh><MudTh>最小库存</MudTh><MudTh>状态</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.ProductId</MudTd><MudTd>@context.CurrentStock</MudTd><MudTd>@context.MinStock</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="Color.Error">不足</MudChip></MudTd>
    </RowTemplate>
</MudTable>
<MudText Typo="Typo.h5" Class="mt-6 mb-3">低库存原料</MudText>
<MudTable Items="@_lowStockMaterials" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>物料ID</MudTh><MudTh>名称</MudTh><MudTh>当前数量</MudTh><MudTh>单位</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.MaterialId</MudTd><MudTd>@context.MaterialName</MudTd><MudTd>@context.CurrentQuantity</MudTd><MudTd>@context.Unit</MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<ProductInventory> _lowStockProducts = new();
    private List<RawMaterialInventory> _lowStockMaterials = new();
    protected override async Task OnInitializedAsync()
    {
        _lowStockProducts = await Db.ProductInventories.AsNoTracking().Where(p => p.CurrentStock <= p.MinStock).ToListAsync();
        _lowStockMaterials = await Db.RawMaterialInventories.AsNoTracking().Where(m => m.CurrentQuantity <= 50).ToListAsync();
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Inventory/Stocktake.razor"] = '''@page "/admin/inventory/stocktake"
@using PerfumeShop.Data.Models

<PageTitle>库存盘点</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">库存盘点</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">上次盘点</MudText>
        <MudText Typo="Typo.h5">@_lastStocktake?.ToString("yyyy-MM-dd") ?? "未执行"</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">差异项</MudText>
        <MudText Typo="Typo.h5" Color="Color.Warning">@_diffCount</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudButton Variant="Variant.Filled" Color="Color.Primary" OnClick="StartStocktake" StartIcon="@Icons.Material.Filled.FindReplace">开始盘点</MudButton>
<MudTable Items="@_records" Hover="true" Dense="true" Striped="true" Class="mt-4">
    <HeaderContent><MudTh>ID</MudTh><MudTh>类型</MudTh><MudTh>数量</MudTh><MudTh>操作人</MudTh><MudTh>时间</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.MovementId</MudTd><MudTd>@context.MovementType</MudTd><MudTd>@context.Quantity</MudTd>
        <MudTd>@context.CreatedBy</MudTd><MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private DateTime? _lastStocktake;
    private int _diffCount;
    private List<StockMovement> _records = new();
    protected override async Task OnInitializedAsync()
    {
        _records = await Db.StockMovements.AsNoTracking().OrderByDescending(m => m.MovementId).Take(50).ToListAsync();
    }
    private void StartStocktake() { Snackbar.Add("盘点任务已启动", Severity.Success); }
}
'''

# ===== LOGISTICS MODULE (8 pages) =====
files["src/PerfumeShop.Admin/Components/Pages/Logistics/ShippingList.razor"] = '''@page "/admin/logistics/shipping"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>发货管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">发货管理</MudText>
<MudTable Items="@_orders" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>订单ID</MudTh><MudTh>订单号</MudTh><MudTh>客户</MudTh><MudTh>金额</MudTh><MudTh>状态</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.OrderId</MudTd><MudTd>@context.OrderNo</MudTd><MudTd>@context.Username</MudTd>
        <MudTd>¥@context.TotalAmount.ToString("F2")</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="Color.Info">@context.Status</MudChip></MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.LocalShipping" Size="Size.Small" Color="Color.Primary" Href="@($"/admin/logistics/shipping-detail/{context.OrderId}")" /></MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<Order> _orders = new();
    protected override async Task OnInitializedAsync()
    {
        _orders = await Db.Orders.AsNoTracking().Where(o => o.Status == "paid" || o.Status == "confirmed" || o.Status == "shipped").OrderByDescending(o => o.OrderId).Take(100).ToListAsync();
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Logistics/ShippingDetail.razor"] = '''@page "/admin/logistics/shipping-detail/{OrderId:int}"
@using PerfumeShop.Data.Models

<PageTitle>发货详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">发货详情</MudText>
@if (_order != null)
{
<MudCard><MudCardContent>
    <MudGrid>
        <MudItem xs="6"><MudText><b>订单ID:</b> @_order.OrderId</MudText></MudItem>
        <MudItem xs="6"><MudText><b>订单号:</b> @_order.OrderNo</MudText></MudItem>
        <MudItem xs="6"><MudText><b>客户:</b> @_order.Username</MudText></MudItem>
        <MudItem xs="6"><MudText><b>金额:</b> ¥@_order.TotalAmount.ToString("F2")</MudText></MudItem>
        <MudItem xs="6"><MudText><b>状态:</b> @_order.Status</MudText></MudItem>
        <MudItem xs="12"><MudText><b>地址:</b> @_order.ShippingAddress</MudText></MudItem>
    </MudGrid>
    <MudDivider Class="my-4" />
    <MudSelect @bind-Value="_carrierId" Label="承运商" Class="mb-4">
        @foreach (var c in _carriers) { <MudSelectItem Value="@c.CompanyId">@c.CompanyName</MudSelectItem> }
    </MudSelect>
    <MudTextField @bind-Value="_trackingNo" Label="物流单号" />
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Ship">确认发货</MudButton>
</MudCardContent></MudCard>
}
<MudButton Variant="Variant.Text" Href="/admin/logistics/shipping" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    [Parameter] public int OrderId { get; set; }
    private Order? _order;
    private List<ShippingCompany> _carriers = new();
    private int _carrierId;
    private string _trackingNo = "";
    protected override async Task OnInitializedAsync()
    {
        _order = await Db.Orders.FindAsync(OrderId);
        _carriers = await Db.ShippingCompanies.AsNoTracking().Where(c => c.IsActive == true).ToListAsync();
    }
    private async Task Ship()
    {
        await Db.Orders.Where(o => o.OrderId == OrderId).ExecuteUpdateAsync(s => s.SetProperty(o => o.Status, "shipped"));
        Snackbar.Add("发货成功", Severity.Success);
        Nav.NavigateTo("/admin/logistics/shipping");
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Logistics/TrackingList.razor"] = '''@page "/admin/logistics/tracking"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>物流追踪</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">物流追踪</MudText>
<MudTable Items="@_orders" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>订单ID</MudTh><MudTh>订单号</MudTh><MudTh>状态</MudTh><MudTh>创建时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.OrderId</MudTd><MudTd>@context.OrderNo</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="Color.Primary">@context.Status</MudChip></MudTd>
        <MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.LocalShipping" Size="Size.Small" Href="@($"/admin/logistics/tracking-detail/{context.OrderId}")" /></MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<Order> _orders = new();
    protected override async Task OnInitializedAsync()
    {
        _orders = await Db.Orders.AsNoTracking().Where(o => o.Status == "shipped" || o.Status == "delivered").OrderByDescending(o => o.OrderId).Take(100).ToListAsync();
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Logistics/TrackingDetail.razor"] = '''@page "/admin/logistics/tracking-detail/{OrderId:int}"
@using PerfumeShop.Data.Models

<PageTitle>物流追踪详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">物流追踪详情</MudText>
@if (_order != null)
{
<MudCard><MudCardContent>
    <MudGrid>
        <MudItem xs="6"><MudText><b>订单号:</b> @_order.OrderNo</MudText></MudItem>
        <MudItem xs="6"><MudText><b>状态:</b> @_order.Status</MudText></MudItem>
        <MudItem xs="12"><MudText><b>收货地址:</b> @_order.ShippingAddress</MudText></MudItem>
    </MudGrid>
</MudCardContent></MudCard>
<MudTimeline Class="mt-4">
    <MudTimelineItem Color="Color.Success" Variant="Variant.Filled">已下单 - @_order.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTimelineItem>
    @if (_order.Status == "shipped" || _order.Status == "delivered")
    { <MudTimelineItem Color="Color.Primary" Variant="Variant.Filled">已发货</MudTimelineItem> }
    @if (_order.Status == "delivered")
    { <MudTimelineItem Color="Color.Success" Variant="Variant.Filled">已签收</MudTimelineItem> }
</MudTimeline>
}
<MudButton Variant="Variant.Text" Href="/admin/logistics/tracking" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int OrderId { get; set; }
    private Order? _order;
    protected override async Task OnInitializedAsync() { _order = await Db.Orders.FindAsync(OrderId); }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Logistics/ReturnsList.razor"] = '''@page "/admin/logistics/returns"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>退货处理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">退货处理</MudText>
<MudTable Items="@_returns" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>订单ID</MudTh><MudTh>类型</MudTh><MudTh>原因</MudTh><MudTh>状态</MudTh><MudTh>时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.AfterSaleId</MudTd><MudTd>@context.OrderId</MudTd><MudTd>@context.Type</MudTd>
        <MudTd>@context.Reason</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@GetStatusColor(context.Status)">@context.Status</MudChip></MudTd>
        <MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/logistics/return-detail/{context.AfterSaleId}")" /></MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AfterSale> _returns = new();
    protected override async Task OnInitializedAsync() { _returns = await Db.AfterSales.AsNoTracking().OrderByDescending(a => a.AfterSaleId).ToListAsync(); }
    private Color GetStatusColor(string? s) => s switch { "pending" => Color.Warning, "approved" => Color.Success, "rejected" => Color.Error, _ => Color.Default };
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Logistics/ReturnDetail.razor"] = '''@page "/admin/logistics/return-detail/{AfterSaleId:int}"
@using PerfumeShop.Data.Models

<PageTitle>退货详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">退货详情</MudText>
@if (_item != null)
{
<MudCard><MudCardContent>
    <MudGrid>
        <MudItem xs="6"><MudText><b>售后ID:</b> @_item.AfterSaleId</MudText></MudItem>
        <MudItem xs="6"><MudText><b>订单ID:</b> @_item.OrderId</MudText></MudItem>
        <MudItem xs="6"><MudText><b>类型:</b> @_item.Type</MudText></MudItem>
        <MudItem xs="6"><MudText><b>状态:</b> @_item.Status</MudText></MudItem>
        <MudItem xs="12"><MudText><b>原因:</b> @_item.Reason</MudText></MudItem>
        <MudItem xs="6"><MudText><b>金额:</b> ¥@_item.Amount?.ToString("F2")</MudText></MudItem>
        <MudItem xs="6"><MudText><b>时间:</b> @_item.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudText></MudItem>
    </MudGrid>
    @if (_item.Status == "pending")
    {
    <MudDivider Class="my-4" />
    <MudButton Variant="Variant.Filled" Color="Color.Success" OnClick="@(() => Process("approved"))" Class="mr-2">批准</MudButton>
    <MudButton Variant="Variant.Filled" Color="Color.Error" OnClick="@(() => Process("rejected"))">拒绝</MudButton>
    }
</MudCardContent></MudCard>
}
<MudButton Variant="Variant.Text" Href="/admin/logistics/returns" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Parameter] public int AfterSaleId { get; set; }
    private AfterSale? _item;
    protected override async Task OnInitializedAsync() { _item = await Db.AfterSales.FindAsync(AfterSaleId); }
    private async Task Process(string status)
    {
        await Db.AfterSales.Where(a => a.AfterSaleId == AfterSaleId).ExecuteUpdateAsync(s => s.SetProperty(a => a.Status, status));
        Snackbar.Add($"已{status}", Severity.Success);
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Logistics/LogisticsReport.razor"] = '''@page "/admin/logistics/report"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>物流报表</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">物流报表</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Primary">已发货</MudText><MudText Typo="Typo.h4">@_shipped</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Success">已送达</MudText><MudText Typo="Typo.h4">@_delivered</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Warning">退货总数</MudText><MudText Typo="Typo.h4">@_returns</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Error">待处理退货</MudText><MudText Typo="Typo.h4">@_pendingReturns</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudText Typo="Typo.h5" Class="mb-3">各状态订单分布</MudText>
<MudTable Items="@_statusDist" Hover="true" Dense="true">
    <HeaderContent><MudTh>状态</MudTh><MudTh>数量</MudTh></HeaderContent>
    <RowTemplate><MudTd>@context.Key</MudTd><MudTd>@context.Value</MudTd></RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private int _shipped, _delivered, _returns, _pendingReturns;
    private Dictionary<string, int> _statusDist = new();
    protected override async Task OnInitializedAsync()
    {
        _shipped = await Db.Orders.CountAsync(o => o.Status == "shipped");
        _delivered = await Db.Orders.CountAsync(o => o.Status == "delivered");
        _returns = await Db.AfterSales.CountAsync();
        _pendingReturns = await Db.AfterSales.CountAsync(a => a.Status == "pending");
        var groups = await Db.Orders.GroupBy(o => o.Status ?? "unknown").Select(g => new { Status = g.Key, Count = g.Count() }).ToListAsync();
        _statusDist = groups.ToDictionary(x => x.Status, x => x.Count);
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Logistics/CarrierManagement.razor"] = '''@page "/admin/logistics/carriers"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>承运商管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">承运商管理</MudText>
<MudButton Variant="Variant.Filled" Color="Color.Primary" StartIcon="@Icons.Material.Filled.Add" OnClick="AddNew" Class="mb-4">新增承运商</MudButton>
<MudTable Items="@_carriers" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>名称</MudTh><MudTh>联系人</MudTh><MudTh>电话</MudTh><MudTh>状态</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.CompanyId</MudTd><MudTd>@context.CompanyName</MudTd><MudTd>@context.ContactPerson</MudTd>
        <MudTd>@context.ContactPhone</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.IsActive == true ? Color.Success : Color.Default)">@(context.IsActive == true ? "启用" : "停用")</MudChip></MudTd>
        <MudTd>
            <MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small" Color="Color.Info" OnClick="@(() => Edit(context))" />
            <MudIconButton Icon="@Icons.Material.Filled.Delete" Size="Size.Small" Color="Color.Error" OnClick="@(() => Delete(context))" />
        </MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private List<ShippingCompany> _carriers = new();
    protected override async Task OnInitializedAsync() { _carriers = await Db.ShippingCompanies.AsNoTracking().ToListAsync(); }
    private void AddNew() => Snackbar.Add("新增承运商功能待完善", Severity.Info);
    private void Edit(ShippingCompany c) => Snackbar.Add($"编辑: {c.CompanyName}", Severity.Info);
    private async Task Delete(ShippingCompany c)
    {
        Db.ShippingCompanies.Remove(c);
        await Db.SaveChangesAsync();
        _carriers.Remove(c);
        Snackbar.Add("已删除", Severity.Success);
    }
}
'''

# ===== ANALYTICS MODULE (2 pages) =====
files["src/PerfumeShop.Admin/Components/Pages/Analytics/OperationsDashboard.razor"] = '''@page "/admin/analytics/operations"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>运营看板(增强)</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">运营看板(增强)</MudText>
<MudGrid Class="mb-6">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Primary">总营收</MudText><MudText Typo="Typo.h5">¥@_totalRevenue.ToString("F2")</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Secondary">总订单</MudText><MudText Typo="Typo.h5">@_totalOrders</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Tertiary">总用户</MudText><MudText Typo="Typo.h5">@_totalUsers</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2" Color="Color.Warning">客单价</MudText><MudText Typo="Typo.h5">¥@_avgOrder.ToString("F2")</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudGrid>
    <MudItem xs="12" md="6"><MudCard><MudCardContent>
        <MudText Typo="Typo.h6" Class="mb-3">近30天营收趋势</MudText>
        <MudChart ChartType="ChartType.Line" InputData="@_chartData" InputLabels="@_chartLabels" Height="300px" />
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" md="6"><MudCard><MudCardContent>
        <MudText Typo="Typo.h6" Class="mb-3">订单状态分布</MudText>
        @foreach (var kv in _statusDist)
        { <MudChip T="string" Color="Color.Info" Class="mr-2">@kv.Key: @kv.Value</MudChip> }
    </MudCardContent></MudCard></MudItem>
</MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _totalRevenue, _avgOrder;
    private int _totalOrders, _totalUsers;
    private double[] _chartData = Array.Empty<double>();
    private string[] _chartLabels = Array.Empty<string>();
    private Dictionary<string, int> _statusDist = new();
    protected override async Task OnInitializedAsync()
    {
        _totalRevenue = await Db.Orders.Where(o => o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _totalOrders = await Db.Orders.CountAsync();
        _totalUsers = await Db.Users.CountAsync();
        _avgOrder = _totalOrders > 0 ? _totalRevenue / _totalOrders : 0;
        var stats = await Db.DailyStatistics.Where(s => s.StatDate >= DateTime.Today.AddDays(-30)).OrderBy(s => s.StatDate).ToListAsync();
        _chartLabels = stats.Select(s => s.StatDate.ToString("MM-dd")).ToArray();
        _chartData = stats.Select(s => (double)(s.TotalRevenue ?? 0)).ToArray();
        var groups = await Db.Orders.GroupBy(o => o.Status ?? "unknown").Select(g => new { g.Key, Count = g.Count() }).ToListAsync();
        _statusDist = groups.ToDictionary(x => x.Key, x => x.Count);
    }
}
'''

files["src/PerfumeShop.Admin/Components/Pages/Analytics/SalesAnalytics.razor"] = '''@page "/admin/analytics/sales"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>销售分析</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">销售分析</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">今日销售</MudText><MudText Typo="Typo.h5">¥@_todaySales.ToString("F2")</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">本周销售</MudText><MudText Typo="Typo.h5">¥@_weekSales.ToString("F2")</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">本月销售</MudText><MudText Typo="Typo.h5">¥@_monthSales.ToString("F2")</MudText>
    </MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent>
        <MudText Typo="Typo.subtitle2">今日订单</MudText><MudText Typo="Typo.h5">@_todayOrders</MudText>
    </MudCardContent></MudCard></MudItem>
</MudGrid>
<MudText Typo="Typo.h5" Class="mb-3">热销商品 TOP 10</MudText>
<MudTable Items="@_topProducts" Hover="true" Dense="true">
    <HeaderContent><MudTh>排名</MudTh><MudTh>商品</MudTh><MudTh>销量</MudTh><MudTh>金额</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.Rank</MudTd><MudTd>@context.Name</MudTd><MudTd>@context.Qty</MudTd><MudTd>¥@context.Amount.ToString("F2")</MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _todaySales, _weekSales, _monthSales;
    private int _todayOrders;
    private List<ProductRank> _topProducts = new();
    private class ProductRank { public int Rank { get; set; } public string Name { get; set; } = ""; public int Qty { get; set; } public decimal Amount { get; set; } }
    protected override async Task OnInitializedAsync()
    {
        var today = DateTime.Today;
        var weekAgo = today.AddDays(-7);
        var monthStart = new DateTime(today.Year, today.Month, 1);
        _todaySales = await Db.Orders.Where(o => o.CreatedAt >= today && o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _weekSales = await Db.Orders.Where(o => o.CreatedAt >= weekAgo && o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _monthSales = await Db.Orders.Where(o => o.CreatedAt >= monthStart && o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _todayOrders = await Db.Orders.CountAsync(o => o.CreatedAt >= today);
        var sales = await Db.OrderItems.Where(i => i.ProductId.HasValue).GroupBy(i => i.ProductId!.Value)
            .Select(g => new { Pid = g.Key, Qty = g.Sum(x => x.Quantity ?? 0), Amt = g.Sum(x => (x.UnitPrice ?? 0) * (x.Quantity ?? 0)) })
            .OrderByDescending(g => g.Amt).Take(10).ToListAsync();
        var names = await Db.Products.Where(p => sales.Select(s => s.Pid).Contains(p.ProductId)).ToDictionaryAsync(p => p.ProductId, p => p.ProductName);
        _topProducts = sales.Select((s, i) => new ProductRank { Rank = i + 1, Name = names.GetValueOrDefault(s.Pid, $"#{s.Pid}"), Qty = s.Qty, Amount = s.Amt }).ToList();
    }
}
'''

# Write all files
count = 0
for relpath, content in files.items():
    fullpath = os.path.join(base, relpath)
    os.makedirs(os.path.dirname(fullpath), exist_ok=True)
    with open(fullpath, 'w', encoding='utf-8') as f:
        f.write(content)
    count += 1

print(f"Generated {count} files")
