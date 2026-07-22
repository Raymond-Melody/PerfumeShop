const fs = require('fs');
const path = require('path');
const base = String.raw`f:\网站制作\网站\网站二`;
function w(rp, c) {
  const fp = path.join(base, rp);
  fs.mkdirSync(path.dirname(fp), { recursive: true });
  fs.writeFileSync(fp, c, 'utf-8');
}
let n = 0;
function W(rp, c) { w(rp, c); n++; }
const sys = "src/PerfumeShop.Admin/Components/Pages/System";

// SessionManagement
W(`${sys}/SessionManagement.razor`, `@page "/admin/system/sessions"

<PageTitle>会话管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">会话管理</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">活跃会话</MudText><MudText Typo="Typo.h4">@_sessions.Count</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTable Items="@_sessions" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>会话ID</MudTh><MudTh>用户</MudTh><MudTh>IP</MudTh><MudTh>开始时间</MudTh><MudTh>最后活动</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.SessionId</MudTd><MudTd>@context.UserName</MudTd><MudTd>@context.Ip</MudTd>
        <MudTd>@context.StartedAt.ToString("yyyy-MM-dd HH:mm")</MudTd><MudTd>@context.LastActivity.ToString("HH:mm:ss")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Close" Size="Size.Small" Color="Color.Error" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    private List<SessionInfo> _sessions = new()
    {
        new() { SessionId = "sess-001", UserName = "admin", Ip = "192.168.1.100", StartedAt = DateTime.Today.AddHours(8), LastActivity = DateTime.Now },
        new() { SessionId = "sess-002", UserName = "operator1", Ip = "192.168.1.101", StartedAt = DateTime.Today.AddHours(9), LastActivity = DateTime.Now.AddMinutes(-5) },
        new() { SessionId = "sess-003", UserName = "finance1", Ip = "192.168.1.102", StartedAt = DateTime.Today.AddHours(10), LastActivity = DateTime.Now.AddMinutes(-2) },
    };
    private class SessionInfo { public string SessionId { get; set; } = ""; public string UserName { get; set; } = ""; public string Ip { get; set; } = ""; public DateTime StartedAt { get; set; } public DateTime LastActivity { get; set; } }
}
`);

// SystemHealth
W(`${sys}/SystemHealth.razor`, `@page "/admin/system/health"

<PageTitle>系统健康监控</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">系统健康监控</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Success">数据库</MudText><MudProgressLinear Value="@99" Color="Color.Success" Class="mt-2" /><MudText Typo="Typo.caption">正常 - 响应 2ms</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Info">内存</MudText><MudProgressLinear Value="@45" Color="Color.Info" Class="mt-2" /><MudText Typo="Typo.caption">45% 已使用</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Warning">CPU</MudText><MudProgressLinear Value="@24" Color="Color.Warning" Class="mt-2" /><MudText Typo="Typo.caption">24% 使用率</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Secondary">磁盘</MudText><MudProgressLinear Value="@62" Color="Color.Secondary" Class="mt-2" /><MudText Typo="Typo.caption">62% 已使用</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTable Items="@_checks" Hover="true" Dense="true">
    <HeaderContent><MudTh>服务</MudTh><MudTh>状态</MudTh><MudTh>响应时间</MudTh><MudTh>最后检查</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.Service</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Healthy ? Color.Success : Color.Error)">@(context.Healthy ? "正常" : "异常")</MudChip></MudTd>
        <MudTd>@context.ResponseMs ms</MudTd><MudTd>@context.CheckedAt.ToString("HH:mm:ss")</MudTd>
    </RowTemplate>
</MudTable>

@code {
    private List<HC> _checks = new()
    {
        new() { Service = "SQL Server", Healthy = true, ResponseMs = 2, CheckedAt = DateTime.Now },
        new() { Service = "Redis Cache", Healthy = true, ResponseMs = 1, CheckedAt = DateTime.Now },
        new() { Service = "IIS AppPool", Healthy = true, ResponseMs = 5, CheckedAt = DateTime.Now },
        new() { Service = "Email Service", Healthy = true, ResponseMs = 120, CheckedAt = DateTime.Now },
        new() { Service = "File Storage", Healthy = true, ResponseMs = 8, CheckedAt = DateTime.Now },
    };
    private class HC { public string Service { get; set; } = ""; public bool Healthy { get; set; } public int ResponseMs { get; set; } public DateTime CheckedAt { get; set; } }
}
`);

// ===== INVENTORY MODULE (3 pages) =====
const inv = "src/PerfumeShop.Admin/Components/Pages/Inventory";
W(`${inv}/Overview.razor`, `@page "/admin/inventory"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>库存总览</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">库存总览</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Primary">成品库存</MudText><MudText Typo="Typo.h4">@_productCount</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Secondary">原料库存</MudText><MudText Typo="Typo.h4">@_materialCount</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Tertiary">瓶子库存</MudText><MudText Typo="Typo.h4">@_bottleCount</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Warning">包装库存</MudText><MudText Typo="Typo.h4">@_packagingCount</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudTabs Elevation="0" Rounded="true" PanelClass="pa-4">
    <MudTabPanel Text="成品"><MudTable Items="@_products" Dense="true" Hover="true"><HeaderContent><MudTh>产品ID</MudTh><MudTh>当前库存</MudTh><MudTh>最小库存</MudTh></HeaderContent><RowTemplate><MudTd>@context.ProductId</MudTd><MudTd>@context.CurrentStock</MudTd><MudTd>@context.MinStock</MudTd></RowTemplate></MudTable></MudTabPanel>
    <MudTabPanel Text="原料"><MudTable Items="@_materials" Dense="true" Hover="true"><HeaderContent><MudTh>物料ID</MudTh><MudTh>名称</MudTh><MudTh>数量</MudTh><MudTh>单位</MudTh></HeaderContent><RowTemplate><MudTd>@context.MaterialId</MudTd><MudTd>@context.MaterialName</MudTd><MudTd>@context.CurrentQuantity</MudTd><MudTd>@context.Unit</MudTd></RowTemplate></MudTable></MudTabPanel>
    <MudTabPanel Text="瓶子"><MudTable Items="@_bottles" Dense="true" Hover="true"><HeaderContent><MudTh>瓶子ID</MudTh><MudTh>数量</MudTh></HeaderContent><RowTemplate><MudTd>@context.BottleId</MudTd><MudTd>@context.CurrentQuantity</MudTd></RowTemplate></MudTable></MudTabPanel>
</MudTabs>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private int _productCount, _materialCount, _bottleCount, _packagingCount;
    private List<ProductInventory> _products = new();
    private List<RawMaterialInventory> _materials = new();
    private List<BottleInventory> _bottles = new();
    protected override async Task OnInitializedAsync()
    {
        _products = await Db.ProductInventories.AsNoTracking().ToListAsync(); _materials = await Db.RawMaterialInventories.AsNoTracking().ToListAsync();
        _bottles = await Db.BottleInventories.AsNoTracking().ToListAsync(); _productCount = _products.Count; _materialCount = _materials.Count;
        _bottleCount = _bottles.Count; _packagingCount = await Db.PackagingInventories.CountAsync();
    }
}
`);

W(`${inv}/Alerts.razor`, `@page "/admin/inventory/alerts"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>库存预警</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">库存预警</MudText>
<MudAlert Severity="Severity.Warning" Class="mb-4">当前有 @_lowProducts.Count 个成品和 @_lowMaterials.Count 种原料库存不足。</MudAlert>
<MudText Typo="Typo.h5" Class="mb-3">低库存成品</MudText>
<MudTable Items="@_lowProducts" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>产品ID</MudTh><MudTh>当前库存</MudTh><MudTh>最小库存</MudTh><MudTh>状态</MudTh></HeaderContent>
    <RowTemplate><MudTd>@context.ProductId</MudTd><MudTd>@context.CurrentStock</MudTd><MudTd>@context.MinStock</MudTd><MudTd><MudChip T="string" Size="Size.Small" Color="Color.Error">不足</MudChip></MudTd></RowTemplate>
</MudTable>
<MudText Typo="Typo.h5" Class="mt-6 mb-3">低库存原料</MudText>
<MudTable Items="@_lowMaterials" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>物料ID</MudTh><MudTh>名称</MudTh><MudTh>数量</MudTh><MudTh>单位</MudTh></HeaderContent>
    <RowTemplate><MudTd>@context.MaterialId</MudTd><MudTd>@context.MaterialName</MudTd><MudTd>@context.CurrentQuantity</MudTd><MudTd>@context.Unit</MudTd></RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<ProductInventory> _lowProducts = new();
    private List<RawMaterialInventory> _lowMaterials = new();
    protected override async Task OnInitializedAsync()
    {
        _lowProducts = await Db.ProductInventories.AsNoTracking().Where(p => p.CurrentStock <= p.MinStock).ToListAsync();
        _lowMaterials = await Db.RawMaterialInventories.AsNoTracking().Where(m => m.CurrentQuantity <= 50).ToListAsync();
    }
}
`);

W(`${inv}/Stocktake.razor`, `@page "/admin/inventory/stocktake"
@using PerfumeShop.Data.Models

<PageTitle>库存盘点</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">库存盘点</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">上次盘点</MudText><MudText Typo="Typo.h5">@_lastDate?.ToString("yyyy-MM-dd") ?? "未执行"</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">差异项</MudText><MudText Typo="Typo.h5" Color="Color.Warning">@_diffCount</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudButton Variant="Variant.Filled" Color="Color.Primary" OnClick="Start" StartIcon="@Icons.Material.Filled.FindReplace" Class="mb-4">开始盘点</MudButton>
<MudTable Items="@_records" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>类型</MudTh><MudTh>数量</MudTh><MudTh>操作人</MudTh><MudTh>时间</MudTh></HeaderContent>
    <RowTemplate><MudTd>@context.MovementId</MudTd><MudTd>@context.MovementType</MudTd><MudTd>@context.Quantity</MudTd><MudTd>@context.CreatedBy</MudTd><MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd></RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private DateTime? _lastDate; private int _diffCount;
    private List<StockMovement> _records = new();
    protected override async Task OnInitializedAsync() { _records = await Db.StockMovements.AsNoTracking().OrderByDescending(m => m.MovementId).Take(50).ToListAsync(); }
    private void Start() { Snackbar.Add("盘点任务已启动", Severity.Success); }
}
`);

// ===== LOGISTICS MODULE (8 pages) =====
const log = "src/PerfumeShop.Admin/Components/Pages/Logistics";
W(`${log}/ShippingList.razor`, `@page "/admin/logistics/shipping"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>发货管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">发货管理</MudText>
<MudTable Items="@_orders" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>订单ID</MudTh><MudTh>订单号</MudTh><MudTh>客户</MudTh><MudTh>金额</MudTh><MudTh>状态</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.OrderId</MudTd><MudTd>@context.OrderNo</MudTd><MudTd>@context.Username</MudTd><MudTd>¥@context.TotalAmount.ToString("F2")</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="Color.Info">@context.Status</MudChip></MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.LocalShipping" Size="Size.Small" Color="Color.Primary" Href="@($"/admin/logistics/shipping-detail/{context.OrderId}")" /></MudTd>
    </RowTemplate><PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<Order> _orders = new();
    protected override async Task OnInitializedAsync() { _orders = await Db.Orders.AsNoTracking().Where(o => o.Status == "paid" || o.Status == "confirmed" || o.Status == "shipped").OrderByDescending(o => o.OrderId).Take(100).ToListAsync(); }
}
`);

W(`${log}/ShippingDetail.razor`, `@page "/admin/logistics/shipping-detail/{OrderId:int}"
@using PerfumeShop.Data.Models

<PageTitle>发货详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">发货详情</MudText>
@if (_order != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>订单ID:</b> @_order.OrderId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>订单号:</b> @_order.OrderNo</MudText></MudItem>
    <MudItem xs="6"><MudText><b>客户:</b> @_order.Username</MudText></MudItem>
    <MudItem xs="6"><MudText><b>金额:</b> ¥@_order.TotalAmount.ToString("F2")</MudText></MudItem>
    <MudItem xs="12"><MudText><b>地址:</b> @_order.ShippingAddress</MudText></MudItem>
</MudGrid>
<MudDivider Class="my-4" />
<MudSelect @bind-Value="_carrierId" Label="承运商" Class="mb-4">@foreach (var c in _carriers) { <MudSelectItem Value="@c.CompanyId">@c.CompanyName</MudSelectItem> }</MudSelect>
<MudTextField @bind-Value="_trackingNo" Label="物流单号" />
<MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Ship">确认发货</MudButton>
</MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/logistics/shipping" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    [Parameter] public int OrderId { get; set; }
    private Order? _order; private List<ShippingCompany> _carriers = new(); private int _carrierId; private string _trackingNo = "";
    protected override async Task OnInitializedAsync() { _order = await Db.Orders.FindAsync(OrderId); _carriers = await Db.ShippingCompanies.AsNoTracking().Where(c => c.IsActive == true).ToListAsync(); }
    private async Task Ship() { await Db.Orders.Where(o => o.OrderId == OrderId).ExecuteUpdateAsync(s => s.SetProperty(o => o.Status, "shipped")); Snackbar.Add("发货成功", Severity.Success); Nav.NavigateTo("/admin/logistics/shipping"); }
}
`);

W(`${log}/TrackingList.razor`, `@page "/admin/logistics/tracking"
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
    </RowTemplate><PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<Order> _orders = new();
    protected override async Task OnInitializedAsync() { _orders = await Db.Orders.AsNoTracking().Where(o => o.Status == "shipped" || o.Status == "delivered").OrderByDescending(o => o.OrderId).Take(100).ToListAsync(); }
}
`);

W(`${log}/TrackingDetail.razor`, `@page "/admin/logistics/tracking-detail/{OrderId:int}"
@using PerfumeShop.Data.Models

<PageTitle>物流追踪详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">物流追踪详情</MudText>
@if (_order != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>订单号:</b> @_order.OrderNo</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_order.Status</MudText></MudItem>
    <MudItem xs="12"><MudText><b>收货地址:</b> @_order.ShippingAddress</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>
<MudTimeline Class="mt-4">
    <MudTimelineItem Color="Color.Success" Variant="Variant.Filled">已下单 - @_order.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTimelineItem>
    @if (_order.Status == "shipped" || _order.Status == "delivered") { <MudTimelineItem Color="Color.Primary" Variant="Variant.Filled">已发货</MudTimelineItem> }
    @if (_order.Status == "delivered") { <MudTimelineItem Color="Color.Success" Variant="Variant.Filled">已签收</MudTimelineItem> }
</MudTimeline>}
<MudButton Variant="Variant.Text" Href="/admin/logistics/tracking" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int OrderId { get; set; }
    private Order? _order;
    protected override async Task OnInitializedAsync() { _order = await Db.Orders.FindAsync(OrderId); }
}
`);

W(`${log}/ReturnsList.razor`, `@page "/admin/logistics/returns"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>退货处理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">退货处理</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>订单ID</MudTh><MudTh>类型</MudTh><MudTh>原因</MudTh><MudTh>状态</MudTh><MudTh>时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.AfterSaleId</MudTd><MudTd>@context.OrderId</MudTd><MudTd>@context.Type</MudTd><MudTd>@context.Reason</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Status == "pending" ? Color.Warning : context.Status == "approved" ? Color.Success : Color.Error)">@context.Status</MudChip></MudTd>
        <MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Visibility" Size="Size.Small" Href="@($"/admin/logistics/return-detail/{context.AfterSaleId}")" /></MudTd>
    </RowTemplate><PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AfterSale> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AfterSales.AsNoTracking().OrderByDescending(a => a.AfterSaleId).ToListAsync(); }
}
`);

W(`${log}/ReturnDetail.razor`, `@page "/admin/logistics/return-detail/{AfterSaleId:int}"
@using PerfumeShop.Data.Models

<PageTitle>退货详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">退货详情</MudText>
@if (_item != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>售后ID:</b> @_item.AfterSaleId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>订单ID:</b> @_item.OrderId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>类型:</b> @_item.Type</MudText></MudItem>
    <MudItem xs="6"><MudText><b>状态:</b> @_item.Status</MudText></MudItem>
    <MudItem xs="12"><MudText><b>原因:</b> @_item.Reason</MudText></MudItem>
    <MudItem xs="6"><MudText><b>金额:</b> ¥@_item.Amount?.ToString("F2")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>时间:</b> @_item.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudText></MudItem>
</MudGrid>
@if (_item.Status == "pending") { <MudDivider Class="my-4" />
    <MudButton Variant="Variant.Filled" Color="Color.Success" OnClick="@(() => Process("approved"))" Class="mr-2">批准</MudButton>
    <MudButton Variant="Variant.Filled" Color="Color.Error" OnClick="@(() => Process("rejected"))">拒绝</MudButton> }
</MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/logistics/returns" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Parameter] public int AfterSaleId { get; set; }
    private AfterSale? _item;
    protected override async Task OnInitializedAsync() { _item = await Db.AfterSales.FindAsync(AfterSaleId); }
    private async Task Process(string status) { await Db.AfterSales.Where(a => a.AfterSaleId == AfterSaleId).ExecuteUpdateAsync(s => s.SetProperty(a => a.Status, status)); Snackbar.Add($"已{status}", Severity.Success); }
}
`);

W(`${log}/LogisticsReport.razor`, `@page "/admin/logistics/report"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>物流报表</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">物流报表</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Primary">已发货</MudText><MudText Typo="Typo.h4">@_shipped</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Success">已送达</MudText><MudText Typo="Typo.h4">@_delivered</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Warning">退货总数</MudText><MudText Typo="Typo.h4">@_returns</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Error">待处理</MudText><MudText Typo="Typo.h4">@_pending</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudText Typo="Typo.h5" Class="mb-3">订单状态分布</MudText>
<MudTable Items="@_statusDist" Hover="true" Dense="true"><HeaderContent><MudTh>状态</MudTh><MudTh>数量</MudTh></HeaderContent><RowTemplate><MudTd>@context.Key</MudTd><MudTd>@context.Value</MudTd></RowTemplate></MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private int _shipped, _delivered, _returns, _pending;
    private Dictionary<string, int> _statusDist = new();
    protected override async Task OnInitializedAsync()
    {
        _shipped = await Db.Orders.CountAsync(o => o.Status == "shipped"); _delivered = await Db.Orders.CountAsync(o => o.Status == "delivered");
        _returns = await Db.AfterSales.CountAsync(); _pending = await Db.AfterSales.CountAsync(a => a.Status == "pending");
        var groups = await Db.Orders.GroupBy(o => o.Status ?? "unknown").Select(g => new { g.Key, Count = g.Count() }).ToListAsync();
        _statusDist = groups.ToDictionary(x => x.Key, x => x.Count);
    }
}
`);

W(`${log}/CarrierManagement.razor`, `@page "/admin/logistics/carriers"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>承运商管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">承运商管理</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>名称</MudTh><MudTh>联系人</MudTh><MudTh>电话</MudTh><MudTh>状态</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.CompanyId</MudTd><MudTd>@context.CompanyName</MudTd><MudTd>@context.ContactPerson</MudTd><MudTd>@context.ContactPhone</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.IsActive == true ? Color.Success : Color.Default)">@(context.IsActive == true ? "启用" : "停用")</MudChip></MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Delete" Size="Size.Small" Color="Color.Error" OnClick="@(() => Delete(context))" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private List<ShippingCompany> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.ShippingCompanies.AsNoTracking().ToListAsync(); }
    private async Task Delete(ShippingCompany c) { Db.ShippingCompanies.Remove(c); await Db.SaveChangesAsync(); _items.Remove(c); Snackbar.Add("已删除", Severity.Success); }
}
`);

// ===== ANALYTICS MODULE (2 pages) =====
const ana = "src/PerfumeShop.Admin/Components/Pages/Analytics";
W(`${ana}/OperationsDashboard.razor`, `@page "/admin/analytics/operations"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>运营看板(增强)</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">运营看板(增强)</MudText>
<MudGrid Class="mb-6">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Primary">总营收</MudText><MudText Typo="Typo.h5">¥@_totalRevenue.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Secondary">总订单</MudText><MudText Typo="Typo.h5">@_totalOrders</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Tertiary">总用户</MudText><MudText Typo="Typo.h5">@_totalUsers</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2" Color="Color.Warning">客单价</MudText><MudText Typo="Typo.h5">¥@_avgOrder.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudGrid>
    <MudItem xs="12" md="6"><MudCard><MudCardContent><MudText Typo="Typo.h6" Class="mb-3">近30天营收趋势</MudText><MudChart ChartType="ChartType.Line" InputData="@_chartData" InputLabels="@_chartLabels" Height="300px" /></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" md="6"><MudCard><MudCardContent><MudText Typo="Typo.h6" Class="mb-3">订单状态分布</MudText>@foreach (var kv in _statusDist) { <MudChip T="string" Color="Color.Info" Class="mr-2">@kv.Key: @kv.Value</MudChip> }</MudCardContent></MudCard></MudItem>
</MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _totalRevenue, _avgOrder; private int _totalOrders, _totalUsers;
    private double[] _chartData = Array.Empty<double>(); private string[] _chartLabels = Array.Empty<string>(); private Dictionary<string, int> _statusDist = new();
    protected override async Task OnInitializedAsync()
    {
        _totalRevenue = await Db.Orders.Where(o => o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _totalOrders = await Db.Orders.CountAsync(); _totalUsers = await Db.Users.CountAsync();
        _avgOrder = _totalOrders > 0 ? _totalRevenue / _totalOrders : 0;
        var stats = await Db.DailyStatistics.Where(s => s.StatDate >= DateTime.Today.AddDays(-30)).OrderBy(s => s.StatDate).ToListAsync();
        _chartLabels = stats.Select(s => s.StatDate.ToString("MM-dd")).ToArray(); _chartData = stats.Select(s => (double)(s.TotalRevenue ?? 0)).ToArray();
        var groups = await Db.Orders.GroupBy(o => o.Status ?? "unknown").Select(g => new { g.Key, Count = g.Count() }).ToListAsync();
        _statusDist = groups.ToDictionary(x => x.Key, x => x.Count);
    }
}
`);

W(`${ana}/SalesAnalytics.razor`, `@page "/admin/analytics/sales"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>销售分析</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">销售分析</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">今日销售</MudText><MudText Typo="Typo.h5">¥@_todaySales.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">本周销售</MudText><MudText Typo="Typo.h5">¥@_weekSales.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">本月销售</MudText><MudText Typo="Typo.h5">¥@_monthSales.ToString("F2")</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">今日订单</MudText><MudText Typo="Typo.h5">@_todayOrders</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudText Typo="Typo.h5" Class="mb-3">热销商品 TOP 10</MudText>
<MudTable Items="@_top" Hover="true" Dense="true"><HeaderContent><MudTh>排名</MudTh><MudTh>商品</MudTh><MudTh>销量</MudTh><MudTh>金额</MudTh></HeaderContent>
<RowTemplate><MudTd>@context.Rank</MudTd><MudTd>@context.Name</MudTd><MudTd>@context.Qty</MudTd><MudTd>¥@context.Amount.ToString("F2")</MudTd></RowTemplate></MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private decimal _todaySales, _weekSales, _monthSales; private int _todayOrders;
    private List<PR> _top = new();
    private class PR { public int Rank { get; set; } public string Name { get; set; } = ""; public int Qty { get; set; } public decimal Amount { get; set; } }
    protected override async Task OnInitializedAsync()
    {
        var today = DateTime.Today; var weekAgo = today.AddDays(-7); var monthStart = new DateTime(today.Year, today.Month, 1);
        _todaySales = await Db.Orders.Where(o => o.CreatedAt >= today && o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _weekSales = await Db.Orders.Where(o => o.CreatedAt >= weekAgo && o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _monthSales = await Db.Orders.Where(o => o.CreatedAt >= monthStart && o.Status != "cancelled").SumAsync(o => (decimal?)o.TotalAmount) ?? 0;
        _todayOrders = await Db.Orders.CountAsync(o => o.CreatedAt >= today);
        var sales = await Db.OrderItems.Where(i => i.ProductId.HasValue).GroupBy(i => i.ProductId!.Value)
            .Select(g => new { Pid = g.Key, Qty = g.Sum(x => x.Quantity ?? 0), Amt = g.Sum(x => (x.UnitPrice ?? 0) * (x.Quantity ?? 0)) })
            .OrderByDescending(g => g.Amt).Take(10).ToListAsync();
        var names = await Db.Products.Where(p => sales.Select(s => s.Pid).Contains(p.ProductId)).ToDictionaryAsync(p => p.ProductId, p => p.ProductName);
        _top = sales.Select((s, i) => new PR { Rank = i + 1, Name = names.GetValueOrDefault(s.Pid, $"#{s.Pid}"), Qty = s.Qty, Amount = s.Amt }).ToList();
    }
}
`);

console.log(`Remaining modules: ${n} pages created (2 sys + 3 inv + 8 log + 2 ana)`);
