const fs = require('fs');
const path = require('path');
const base = String.raw`f:\网站制作\网站\网站二`;
function w(rp, c) { const fp = path.join(base, rp); fs.mkdirSync(path.dirname(fp), { recursive: true }); fs.writeFileSync(fp, c, 'utf-8'); }
let n = 0; function W(rp, c) { w(rp, c); n++; }
const fin = "src/PerfumeShop.Admin/Components/Pages/Finance";

// 1. PayableList
W(`${fin}/PayableList.razor`, `@page "/admin/finance/payables"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>应付管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">应付管理</MudText>
<MudButton Variant="Variant.Filled" Color="Color.Primary" Href="/admin/finance/payable-create" StartIcon="@Icons.Material.Filled.Add" Class="mb-4">新建应付</MudButton>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>ID</MudTh><MudTh>应付单号</MudTh><MudTh>供应商</MudTh><MudTh>金额</MudTh><MudTh>已付</MudTh><MudTh>状态</MudTh><MudTh>到期日</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.PayableId</MudTd><MudTd>@context.PayableNo</MudTd><MudTd>@context.SupplierName</MudTd><MudTd>¥@context.Amount?.ToString("F2")</MudTd><MudTd>¥@context.PaidAmount?.ToString("F2")</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Status == "paid" ? Color.Success : context.Status == "overdue" ? Color.Error : Color.Warning)">@context.Status</MudChip></MudTd>
        <MudTd>@context.DueDate?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/finance/payable-detail/{context.PayableId}")" /></MudTd>
    </RowTemplate><PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AccountsPayable> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AccountsPayables.AsNoTracking().OrderByDescending(p => p.PayableId).ToListAsync(); }
}
`);

// 2. PayableDetail
W(`${fin}/PayableDetail.razor`, `@page "/admin/finance/payable-detail/{PayableId:int}"
@using PerfumeShop.Data.Models

<PageTitle>应付详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">应付详情</MudText>
@if (_item != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>应付ID:</b> @_item.PayableId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>单号:</b> @_item.PayableNo</MudText></MudItem>
    <MudItem xs="6"><MudText><b>供应商:</b> @_item.SupplierName</MudText></MudItem>
    <MudItem xs="6"><MudText><b>金额:</b> ¥@_item.Amount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>已付:</b> ¥@_item.PaidAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_item.Status</MudText></MudItem>
    <MudItem xs="6"><MudText><b>到期日:</b> @_item.DueDate?.ToString("yyyy-MM-dd")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>创建时间:</b> @_item.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/finance/payables" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int PayableId { get; set; }
    private AccountsPayable? _item;
    protected override async Task OnInitializedAsync() { _item = await Db.AccountsPayables.FindAsync(PayableId); }
}
`);

// 3. PayableCreate
W(`${fin}/PayableCreate.razor`, `@page "/admin/finance/payable-create"
@using PerfumeShop.Data.Models

<PageTitle>新建应付</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">新建应付</MudText>
<MudGrid><MudItem xs="12" md="6"><MudForm>
    <MudTextField @bind-Value="_item.PayableNo" Label="应付单号" Required="true" />
    <MudTextField @bind-Value="_item.SupplierName" Label="供应商" />
    <MudNumericField @bind-Value="_amount" Label="金额" />
    <MudDatePicker @bind-Value="_dueDate" Label="到期日" />
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/finance/payables" Class="mt-4 ml-2">返回</MudButton>
</MudForm></MudItem></MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private AccountsPayable _item = new(); private decimal _amount; private DateTime? _dueDate;
    private async Task Save() { _item.Amount = _amount; _item.DueDate = _dueDate; _item.Status = "pending"; _item.CreatedAt = DateTime.Now; Db.AccountsPayables.Add(_item); await Db.SaveChangesAsync(); Snackbar.Add("创建成功", Severity.Success); Nav.NavigateTo("/admin/finance/payables"); }
}
`);

// 4. ReceivableList
W(`${fin}/ReceivableList.razor`, `@page "/admin/finance/receivables"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>应收管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">应收管理</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>ID</MudTh><MudTh>应收单号</MudTh><MudTh>客户</MudTh><MudTh>金额</MudTh><MudTh>已收</MudTh><MudTh>状态</MudTh><MudTh>到期日</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.ReceivableId</MudTd><MudTd>@context.ReceivableNo</MudTd><MudTd>@context.CustomerName</MudTd><MudTd>¥@context.Amount?.ToString("F2")</MudTd><MudTd>¥@context.ReceivedAmount?.ToString("F2")</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Status == "received" ? Color.Success : Color.Warning)">@context.Status</MudChip></MudTd>
        <MudTd>@context.DueDate?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/finance/receivable-detail/{context.ReceivableId}")" /></MudTd>
    </RowTemplate><PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AccountsReceivable> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AccountsReceivables.AsNoTracking().OrderByDescending(r => r.ReceivableId).ToListAsync(); }
}
`);

// 5. ReceivableDetail
W(`${fin}/ReceivableDetail.razor`, `@page "/admin/finance/receivable-detail/{ReceivableId:int}"
@using PerfumeShop.Data.Models

<PageTitle>应收详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">应收详情</MudText>
@if (_item != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>应收ID:</b> @_item.ReceivableId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>单号:</b> @_item.ReceivableNo</MudText></MudItem>
    <MudItem xs="6"><MudText><b>客户:</b> @_item.CustomerName</MudText></MudItem>
    <MudItem xs="6"><MudText><b>金额:</b> ¥@_item.Amount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>已收:</b> ¥@_item.ReceivedAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_item.Status</MudText></MudItem>
    <MudItem xs="6"><MudText><b>到期日:</b> @_item.DueDate?.ToString("yyyy-MM-dd")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/finance/receivables" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int ReceivableId { get; set; }
    private AccountsReceivable? _item;
    protected override async Task OnInitializedAsync() { _item = await Db.AccountsReceivables.FindAsync(ReceivableId); }
}
`);

// 6-8: Reports
W(`${fin}/RevenueReport.razor`, `@page "/admin/finance/revenue-report"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>营收报表</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">营收报表</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Primary">总营收</MudText><MudText Typo="Typo.h5">¥@_total.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Success">本月营收</MudText><MudText Typo="Typo.h5">¥@_month.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">订单总数</MudText><MudText Typo="Typo.h5">@_orderCount</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudChart ChartType="ChartType.Bar" InputData="@_chartData" InputLabels="@_chartLabels" Height="300px" />

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _total, _month; private int _orderCount;
    private double[] _chartData = Array.Empty<double>(); private string[] _chartLabels = Array.Empty<string>();
    protected override async Task OnInitializedAsync()
    {
        _total = await Db.Orders.Where(o => o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        var ms = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _month = await Db.Orders.Where(o => o.CreatedAt >= ms && o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _orderCount = await Db.Orders.CountAsync();
        var stats = await Db.DailyStatistics.Where(s => s.StatDate >= DateTime.Today.AddDays(-30)).OrderBy(s => s.StatDate).ToListAsync();
        _chartLabels = stats.Select(s => s.StatDate.ToString("MM-dd")).ToArray(); _chartData = stats.Select(s => (double)(s.TotalRevenue ?? 0)).ToArray();
    }
}
`);

W(`${fin}/CostReport.razor`, `@page "/admin/finance/cost-report"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>成本报表</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">成本报表</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Error">总成本</MudText><MudText Typo="Typo.h5">¥@_totalCost.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">成本中心数</MudText><MudText Typo="Typo.h5">@_centerCount</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTable Items="@_expenses" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>费用名</MudTh><MudTh>类型</MudTh><MudTh>金额</MudTh><MudTh>期间</MudTh></HeaderContent>
    <RowTemplate><MudTd>@context.ExpenseId</MudTd><MudTd>@context.ExpenseName</MudTd><MudTd>@context.ExpenseType</MudTd><MudTd>¥@context.Amount?.ToString("F2")</MudTd><MudTd>@context.Period</MudTd></RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _totalCost; private int _centerCount;
    private List<ExpenseRecord> _expenses = new();
    protected override async Task OnInitializedAsync()
    {
        _expenses = await Db.ExpenseRecords.AsNoTracking().OrderByDescending(e => e.ExpenseId).Take(100).ToListAsync();
        _totalCost = _expenses.Sum(e => e.Amount ?? 0);
        _centerCount = await Db.CostCenters.CountAsync();
    }
}
`);

W(`${fin}/ProfitReport.razor`, `@page "/admin/finance/profit-report"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>利润报表</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">利润报表</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Success">总营收</MudText><MudText Typo="Typo.h5">¥@_revenue.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Error">总成本</MudText><MudText Typo="Typo.h5">¥@_cost.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Primary">净利润</MudText><MudText Typo="Typo.h5">¥@(_revenue - _cost).ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudAlert Severity="@( (_revenue - _cost) > 0 ? Severity.Success : Severity.Error)">利润率: @(_revenue > 0 ? ((_revenue - _cost) / _revenue * 100).ToString("F1") : "0.0")%</MudAlert>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _revenue, _cost;
    protected override async Task OnInitializedAsync()
    {
        _revenue = await Db.Orders.Where(o => o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _cost = await Db.ExpenseRecords.SumAsync(e => (decimal?)e.Amount) ?? 0;
    }
}
`);

// 9-11: Cost Analysis
W(`${fin}/CostAnalysis.razor`, `@page "/admin/finance/cost-analysis"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>成本分析</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">成本分析</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">费用总数</MudText><MudText Typo="Typo.h5">@_expenseCount</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">总费用</MudText><MudText Typo="Typo.h5">¥@_totalExpense.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">成本中心</MudText><MudText Typo="Typo.h5">@_centerCount</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTable Items="@_expenses" Hover="true" Dense="true">
    <HeaderContent><MudTh>费用名</MudTh><MudTh>类型</MudTh><MudTh>金额</MudTh><MudTh>分摊方式</MudTh><MudTh>期间</MudTh></HeaderContent>
    <RowTemplate><MudTd>@context.ExpenseName</MudTd><MudTd>@context.ExpenseType</MudTd><MudTd>¥@context.Amount?.ToString("F2")</MudTd><MudTd>@context.AllocationMethod</MudTd><MudTd>@context.Period</MudTd></RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private int _expenseCount, _centerCount; private decimal _totalExpense;
    private List<ExpenseRecord> _expenses = new();
    protected override async Task OnInitializedAsync()
    {
        _expenses = await Db.ExpenseRecords.AsNoTracking().OrderByDescending(e => e.ExpenseId).Take(100).ToListAsync();
        _expenseCount = _expenses.Count; _totalExpense = _expenses.Sum(e => e.Amount ?? 0);
        _centerCount = await Db.CostCenters.CountAsync();
    }
}
`);

W(`${fin}/CostBreakdown.razor`, `@page "/admin/finance/cost-breakdown"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>成本明细</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">成本明细</MudText>
<MudTable Items="@_centers" Hover="true" Dense="true" Striped="true" Class="mb-4">
    <HeaderContent><MudTh>中心代码</MudTh><MudTh>名称</MudTh><MudTh>类型</MudTh><MudTh>预算</MudTh><MudTh>状态</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.CenterCode</MudTd><MudTd>@context.CenterName</MudTd><MudTd>@context.CenterType</MudTd>
        <MudTd>¥@context.BudgetAmount?.ToString("F2")</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.IsActive == true ? Color.Success : Color.Default)">@(context.IsActive == true ? "启用" : "停用")</MudChip></MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<CostCenter> _centers = new();
    protected override async Task OnInitializedAsync() { _centers = await Db.CostCenters.AsNoTracking().ToListAsync(); }
}
`);

W(`${fin}/CostTrend.razor`, `@page "/admin/finance/cost-trend"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>成本趋势</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">成本趋势</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="6"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">近30天费用趋势</MudText><MudChart ChartType="ChartType.Line" InputData="@_chartData" InputLabels="@_chartLabels" Height="300px" /></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="6"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">按类型分布</MudText>
    @foreach (var kv in _typeDist) { <MudChip T="string" Color="Color.Info" Class="mr-2 mb-2">@kv.Key: ¥@kv.Value.ToString("F2")</MudChip> }
    </MudCardContent></MudCard></MudItem>
</MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private double[] _chartData = Array.Empty<double>(); private string[] _chartLabels = Array.Empty<string>();
    private Dictionary<string, decimal> _typeDist = new();
    protected override async Task OnInitializedAsync()
    {
        var expenses = await Db.ExpenseRecords.AsNoTracking().ToListAsync();
        var groups = expenses.GroupBy(e => e.ExpenseType ?? "other").Select(g => new { Type = g.Key, Total = g.Sum(x => x.Amount ?? 0) });
        _typeDist = groups.ToDictionary(x => x.Type, x => x.Total);
        var byPeriod = expenses.Where(e => e.Period != null).GroupBy(e => e.Period!).OrderBy(g => g.Key).Select(g => new { Period = g.Key, Total = g.Sum(x => x.Amount ?? 0) }).Take(30).ToList();
        _chartLabels = byPeriod.Select(b => b.Period).ToArray(); _chartData = byPeriod.Select(b => (double)b.Total).ToArray();
    }
}
`);

// 12-15: Budget
W(`${fin}/BudgetList.razor`, `@page "/admin/finance/budgets"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>预算管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">预算管理</MudText>
<MudButton Variant="Variant.Filled" Color="Color.Primary" Href="/admin/finance/budget-create" StartIcon="@Icons.Material.Filled.Add" Class="mb-4">新建预算</MudButton>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>预算名</MudTh><MudTh>分类</MudTh><MudTh>预算额</MudTh><MudTh>实际额</MudTh><MudTh>执行率</MudTh><MudTh>状态</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.BudgetId</MudTd><MudTd>@context.BudgetName</MudTd><MudTd>@context.Category</MudTd><MudTd>¥@context.BudgetAmount?.ToString("F2")</MudTd><MudTd>¥@context.ActualAmount?.ToString("F2")</MudTd>
        <MudTd>@(context.BudgetAmount > 0 ? (context.ActualAmount / context.BudgetAmount * 100)?.ToString("F1") + "%" : "-")</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Status == "approved" ? Color.Success : Color.Warning)">@context.Status</MudChip></MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/finance/budget-detail/{context.BudgetId}")" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<BudgetPlan> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.BudgetPlans.AsNoTracking().OrderByDescending(b => b.BudgetId).ToListAsync(); }
}
`);

W(`${fin}/BudgetCreate.razor`, `@page "/admin/finance/budget-create"
@using PerfumeShop.Data.Models

<PageTitle>新建预算</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">新建预算</MudText>
<MudGrid><MudItem xs="12" md="6"><MudForm>
    <MudTextField @bind-Value="_b.BudgetName" Label="预算名" Required="true" />
    <MudTextField @bind-Value="_b.Category" Label="分类" />
    <MudNumericField @bind-Value="_budgetAmt" Label="预算金额" />
    <MudTextField @bind-Value="_b.Period" Label="期间" HelperText="例: 2026-Q1" />
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/finance/budgets" Class="mt-4 ml-2">返回</MudButton>
</MudForm></MudItem></MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private BudgetPlan _b = new(); private decimal _budgetAmt;
    private async Task Save() { _b.BudgetAmount = _budgetAmt; _b.Status = "draft"; _b.CreatedAt = DateTime.Now; Db.BudgetPlans.Add(_b); await Db.SaveChangesAsync(); Snackbar.Add("创建成功", Severity.Success); Nav.NavigateTo("/admin/finance/budgets"); }
}
`);

W(`${fin}/BudgetDetail.razor`, `@page "/admin/finance/budget-detail/{BudgetId:int}"
@using PerfumeShop.Data.Models

<PageTitle>预算详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">预算详情</MudText>
@if (_b != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>预算ID:</b> @_b.BudgetId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>预算名:</b> @_b.BudgetName</MudText></MudItem>
    <MudItem xs="6"><MudText><b>分类:</b> @_b.Category</MudText></MudItem>
    <MudItem xs="6"><MudText><b>预算额:</b> ¥@_b.BudgetAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>实际额:</b> ¥@_b.ActualAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>ROI:</b> @_b.Roi?.ToString("F2")%</MudText></MudItem>
    <MudItem xs="6"><MudText><b>期间:</b> @_b.Period</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_b.Status</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>
<MudButton Variant="Variant.Outlined" Color="Color.Info" Href="@($"/admin/finance/budget-compare/{_b.BudgetId}")" Class="mt-4">对比分析</MudButton>}
<MudButton Variant="Variant.Text" Href="/admin/finance/budgets" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int BudgetId { get; set; }
    private BudgetPlan? _b;
    protected override async Task OnInitializedAsync() { _b = await Db.BudgetPlans.FindAsync(BudgetId); }
}
`);

W(`${fin}/BudgetCompare.razor`, `@page "/admin/finance/budget-compare/{BudgetId:int}"
@using PerfumeShop.Data.Models

<PageTitle>预算对比</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">预算对比</MudText>
@if (_b != null) {
<MudGrid>
    <MudItem xs="12" sm="6"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">预算金额</MudText><MudText Typo="Typo.h4" Color="Color.Primary">¥@_b.BudgetAmount?.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="6"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">实际金额</MudText><MudText Typo="Typo.h4" Color="@((_b.ActualAmount ?? 0) > (_b.BudgetAmount ?? 0) ? Color.Error : Color.Success)">¥@_b.ActualAmount?.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudAlert Severity="@((_b.ActualAmount ?? 0) <= (_b.BudgetAmount ?? 0) ? Severity.Success : Severity.Warning)" Class="mt-4">
    差异: ¥@((_b.ActualAmount - _b.BudgetAmount)?.ToString("F2")) (@((_b.BudgetAmount > 0 ? (_b.ActualAmount - _b.BudgetAmount) / _b.BudgetAmount * 100 : 0)?.ToString("F1"))%)
</MudAlert>}
<MudButton Variant="Variant.Text" Href="/admin/finance/budgets" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int BudgetId { get; set; }
    private BudgetPlan? _b;
    protected override async Task OnInitializedAsync() { _b = await Db.BudgetPlans.FindAsync(BudgetId); }
}
`);

// 16-18: Reconciliation
W(`${fin}/ReconciliationList.razor`, `@page "/admin/finance/reconciliations"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>对账管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">对账管理</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>订单号</MudTh><MudTh>订单金额</MudTh><MudTh>支付金额</MudTh><MudTh>差异</MudTh><MudTh>状态</MudTh><MudTh>对账日期</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.LogId</MudTd><MudTd>@context.OrderNo</MudTd><MudTd>¥@context.OrderAmount?.ToString("F2")</MudTd><MudTd>¥@context.PaymentAmount?.ToString("F2")</MudTd>
        <MudTd Color="@((context.Difference ?? 0) != 0 ? Color.Error : Color.Success)">¥@context.Difference?.ToString("F2")</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Status == "matched" ? Color.Success : Color.Warning)">@context.Status</MudChip></MudTd>
        <MudTd>@context.ReconcileDate?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/finance/reconciliation-detail/{context.LogId}")" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<ReconciliationLog> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.ReconciliationLogs.AsNoTracking().OrderByDescending(r => r.LogId).ToListAsync(); }
}
`);

W(`${fin}/ReconciliationDetail.razor`, `@page "/admin/finance/reconciliation-detail/{LogId:int}"
@using PerfumeShop.Data.Models

<PageTitle>对账详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">对账详情</MudText>
@if (_item != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>日志ID:</b> @_item.LogId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>订单号:</b> @_item.OrderNo</MudText></MudItem>
    <MudItem xs="6"><MudText><b>订单金额:</b> ¥@_item.OrderAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>支付金额:</b> ¥@_item.PaymentAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>差异:</b> ¥@_item.Difference?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_item.Status</MudText></MudItem>
    <MudItem xs="12"><MudText><b>解决方案:</b> @_item.Resolution</MudText></MudItem>
    <MudItem xs="6"><MudText><b>对账日期:</b> @_item.ReconcileDate?.ToString("yyyy-MM-dd")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>处理人:</b> @_item.ResolvedBy</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/finance/reconciliations" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int LogId { get; set; }
    private ReconciliationLog? _item;
    protected override async Task OnInitializedAsync() { _item = await Db.ReconciliationLogs.FindAsync(LogId); }
}
`);

W(`${fin}/ReconciliationCreate.razor`, `@page "/admin/finance/reconciliation-create"

<PageTitle>新建对账</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">新建对账</MudText>
<MudCard><MudCardContent>
    <MudDatePicker @bind-Value="_date" Label="对账日期" />
    <MudSelect @bind-Value="_scope" Label="对账范围" Class="mt-4">
        <MudSelectItem Value="@("daily")">日对账</MudSelectItem><MudSelectItem Value="@("weekly")">周对账</MudSelectItem><MudSelectItem Value="@("monthly")">月对账</MudSelectItem>
    </MudSelect>
</MudCardContent>
<MudCardActions>
    <MudButton Variant="Variant.Filled" Color="Color.Primary" OnClick="Create">开始对账</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/finance/reconciliations">返回</MudButton>
</MudCardActions></MudCard>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private DateTime? _date = DateTime.Today; private string _scope = "daily";
    private async Task Create() { await Task.Delay(500); Snackbar.Add("对账完成", Severity.Success); Nav.NavigateTo("/admin/finance/reconciliations"); }
}
`);

// 19-21: Invoice
W(`${fin}/InvoiceList.razor`, `@page "/admin/finance/invoices"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>发票管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">发票管理</MudText>
<MudButton Variant="Variant.Filled" Color="Color.Primary" Href="/admin/finance/invoice-create" StartIcon="@Icons.Material.Filled.Add" Class="mb-4">新建发票</MudButton>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>支付单号</MudTh><MudTh>订单ID</MudTh><MudTh>金额</MudTh><MudTh>方式</MudTh><MudTh>状态</MudTh><MudTh>时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.RecordId</MudTd><MudTd>@context.TransactionNo</MudTd><MudTd>@context.OrderId</MudTd><MudTd>¥@context.Amount?.ToString("F2")</MudTd>
        <MudTd>@context.PaymentMethod</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Status == "success" ? Color.Success : Color.Warning)">@context.Status</MudChip></MudTd>
        <MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/finance/invoice-detail/{context.RecordId}")" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<PaymentRecord> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.PaymentRecords.AsNoTracking().OrderByDescending(p => p.RecordId).ToListAsync(); }
}
`);

W(`${fin}/InvoiceDetail.razor`, `@page "/admin/finance/invoice-detail/{RecordId:int}"
@using PerfumeShop.Data.Models

<PageTitle>发票详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">发票详情</MudText>
@if (_item != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>记录ID:</b> @_item.RecordId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>交易号:</b> @_item.TransactionNo</MudText></MudItem>
    <MudItem xs="6"><MudText><b>订单ID:</b> @_item.OrderId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>金额:</b> ¥@_item.Amount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>方式:</b> @_item.PaymentMethod</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_item.Status</MudText></MudItem>
    <MudItem xs="6"><MudText><b>时间:</b> @_item.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/finance/invoices" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int RecordId { get; set; }
    private PaymentRecord? _item;
    protected override async Task OnInitializedAsync() { _item = await Db.PaymentRecords.FindAsync(RecordId); }
}
`);

W(`${fin}/InvoiceCreate.razor`, `@page "/admin/finance/invoice-create"

<PageTitle>新建发票</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">新建发票</MudText>
<MudGrid><MudItem xs="12" md="6"><MudForm>
    <MudNumericField @bind-Value="_orderId" Label="订单ID" Required="true" />
    <MudNumericField @bind-Value="_amount" Label="金额" />
    <MudSelect @bind-Value="_method" Label="支付方式" Class="mt-4">
        <MudSelectItem Value="@("alipay")">支付宝</MudSelectItem><MudSelectItem Value="@("wechat")">微信支付</MudSelectItem><MudSelectItem Value="@("bank")">银行转账</MudSelectItem>
    </MudSelect>
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Create">创建</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/finance/invoices" Class="mt-4 ml-2">返回</MudButton>
</MudForm></MudItem></MudGrid>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private int _orderId; private decimal _amount; private string _method = "alipay";
    private void Create() { Snackbar.Add("发票创建成功", Severity.Success); Nav.NavigateTo("/admin/finance/invoices"); }
}
`);

// 22-24: Refund
W(`${fin}/RefundList.razor`, `@page "/admin/finance/refunds"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>退款管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">退款管理</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>订单ID</MudTh><MudTh>金额</MudTh><MudTh>原因</MudTh><MudTh>状态</MudTh><MudTh>时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.RefundId</MudTd><MudTd>@context.OrderId</MudTd><MudTd>¥@context.RefundAmount?.ToString("F2")</MudTd><MudTd>@context.Reason</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Status == "approved" ? Color.Success : context.Status == "rejected" ? Color.Error : Color.Warning)">@context.Status</MudChip></MudTd>
        <MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd>
            <MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/finance/refund-detail/{context.RefundId}")" />
            @if (context.Status == "pending") { <MudIconButton Icon="@Icons.Material.Filled.CheckCircle" Size="Size.Small" Color="Color.Success" Href="@($"/admin/finance/refund-approve/{context.RefundId}")" /> }
        </MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<RefundRecord> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.RefundRecords.AsNoTracking().OrderByDescending(r => r.RefundId).ToListAsync(); }
}
`);

W(`${fin}/RefundDetail.razor`, `@page "/admin/finance/refund-detail/{RefundId:int}"
@using PerfumeShop.Data.Models

<PageTitle>退款详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">退款详情</MudText>
@if (_item != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>退款ID:</b> @_item.RefundId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>订单ID:</b> @_item.OrderId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>退款金额:</b> ¥@_item.RefundAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_item.Status</MudText></MudItem>
    <MudItem xs="12"><MudText><b>原因:</b> @_item.Reason</MudText></MudItem>
    <MudItem xs="6"><MudText><b>创建时间:</b> @_item.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/finance/refunds" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int RefundId { get; set; }
    private RefundRecord? _item;
    protected override async Task OnInitializedAsync() { _item = await Db.RefundRecords.FindAsync(RefundId); }
}
`);

W(`${fin}/RefundApprove.razor`, `@page "/admin/finance/refund-approve/{RefundId:int}"
@using PerfumeShop.Data.Models

<PageTitle>退款审批</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">退款审批</MudText>
@if (_item != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>退款ID:</b> @_item.RefundId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>订单ID:</b> @_item.OrderId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>退款金额:</b> ¥@_item.RefundAmount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="12"><MudText><b>原因:</b> @_item.Reason</MudText></MudItem>
</MudGrid>
<MudDivider Class="my-4" />
<MudTextField @bind-Value="_notes" Label="审批备注" Lines="3" />
<MudStack Row="true" Class="mt-4" Spacing="2">
    <MudButton Variant="Variant.Filled" Color="Color.Success" OnClick="@(() => Process("approved"))">批准退款</MudButton>
    <MudButton Variant="Variant.Filled" Color="Color.Error" OnClick="@(() => Process("rejected"))">拒绝</MudButton>
</MudStack>
</MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/finance/refunds" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    [Parameter] public int RefundId { get; set; }
    private RefundRecord? _item; private string _notes = "";
    protected override async Task OnInitializedAsync() { _item = await Db.RefundRecords.FindAsync(RefundId); }
    private async Task Process(string status)
    {
        await Db.RefundRecords.Where(r => r.RefundId == RefundId).ExecuteUpdateAsync(s => s.SetProperty(r => r.Status, status));
        Snackbar.Add($"退款已{status}", Severity.Success); Nav.NavigateTo("/admin/finance/refunds");
    }
}
`);

// 25. FinancialDashboard
W(`${fin}/FinancialDashboard.razor`, `@page "/admin/finance/dashboard"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>财务看板</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">财务看板</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Success">总营收</MudText><MudText Typo="Typo.h5">¥@_revenue.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Error">应付总额</MudText><MudText Typo="Typo.h5">¥@_payable.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Primary">应收总额</MudText><MudText Typo="Typo.h5">¥@_receivable.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Warning">退款总额</MudText><MudText Typo="Typo.h5">¥@_refunds.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudGrid>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">待对账</MudText><MudText Typo="Typo.h5">@_pendingRecon 笔</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">待审批退款</MudText><MudText Typo="Typo.h5">@_pendingRefunds 笔</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">逾期应付</MudText><MudText Typo="Typo.h5" Color="Color.Error">@_overduePayables 笔</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _revenue, _payable, _receivable, _refunds;
    private int _pendingRecon, _pendingRefunds, _overduePayables;
    protected override async Task OnInitializedAsync()
    {
        _revenue = await Db.Orders.Where(o => o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _payable = await Db.AccountsPayables.SumAsync(p => (decimal?)p.Amount) ?? 0;
        _receivable = await Db.AccountsReceivables.SumAsync(r => (decimal?)r.Amount) ?? 0;
        _refunds = await Db.RefundRecords.SumAsync(r => (decimal?)r.RefundAmount) ?? 0;
        _pendingRecon = await Db.ReconciliationLogs.CountAsync(r => r.Status == "pending");
        _pendingRefunds = await Db.RefundRecords.CountAsync(r => r.Status == "pending");
        _overduePayables = await Db.AccountsPayables.CountAsync(p => p.Status == "overdue");
    }
}
`);

console.log(`Finance module: ${n} pages created`);
