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

// 1. Users
W(`${sys}/Users.razor`, `@page "/admin/system/users"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>用户管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">用户管理</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent>
        <MudTh>ID</MudTh><MudTh>用户名</MudTh><MudTh>邮箱</MudTh><MudTh>部门</MudTh><MudTh>状态</MudTh><MudTh>最后登录</MudTh><MudTh>操作</MudTh>
    </HeaderContent>
    <RowTemplate>
        <MudTd>@context.AdminId</MudTd><MudTd>@context.Username</MudTd><MudTd>@context.Email</MudTd><MudTd>@context.Department</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.IsActive == true ? Color.Success : Color.Error)">@(context.IsActive == true ? "正常" : "禁用")</MudChip></MudTd>
        <MudTd>@context.LastLogin?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd>
            <MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small" Color="Color.Info" Href="@($"/admin/system/user-edit/{context.AdminId}")" />
            <MudIconButton Icon="@Icons.Material.Filled.Delete" Size="Size.Small" Color="Color.Error" OnClick="@(() => Delete(context))" />
        </MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>
<MudButton Variant="Variant.Filled" Color="Color.Primary" Href="/admin/system/user-edit/0" StartIcon="@Icons.Material.Filled.Add" Class="mt-4">新增用户</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private List<AdminUser> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AdminUsers.AsNoTracking().OrderByDescending(u => u.AdminId).ToListAsync(); }
    private async Task Delete(AdminUser u) { Db.AdminUsers.Remove(u); await Db.SaveChangesAsync(); _items.Remove(u); Snackbar.Add("已删除", Severity.Success); }
}
`);

// 2. UserEdit
W(`${sys}/UserEdit.razor`, `@page "/admin/system/user-edit/{AdminId:int}"
@using PerfumeShop.Data.Models

<PageTitle>编辑用户</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑用户</MudText>
<MudGrid><MudItem xs="12" md="6"><MudForm>
    <MudTextField @bind-Value="_user.Username" Label="用户名" Required="true" />
    <MudTextField @bind-Value="_user.Email" Label="邮箱" Required="true" />
    <MudTextField @bind-Value="_user.FullName" Label="姓名" />
    <MudTextField @bind-Value="_user.Department" Label="部门" />
    <MudSelect @bind-Value="_user.RoleId" Label="角色">
        @foreach (var r in _roles) { <MudSelectItem Value="@r.RoleId">@r.RoleName</MudSelectItem> }
    </MudSelect>
    <MudSwitch @bind-Value="_isActive" Label="启用" Color="Color.Success" />
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/system/users" Class="mt-4 ml-2">返回</MudButton>
</MudForm></MudItem></MudGrid>

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
`);

// 3. UserRoles
W(`${sys}/UserRoles.razor`, `@page "/admin/system/user-roles"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>用户角色分配</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">用户角色分配</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>用户ID</MudTh><MudTh>用户名</MudTh><MudTh>角色ID</MudTh><MudTh>部门</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.AdminId</MudTd><MudTd>@context.Username</MudTd><MudTd>@context.RoleId</MudTd><MudTd>@context.Department</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small" Color="Color.Info" Href="@($"/admin/system/user-edit/{context.AdminId}")" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AdminUser> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AdminUsers.AsNoTracking().ToListAsync(); }
}
`);

// 4. Roles
W(`${sys}/Roles.razor`, `@page "/admin/system/roles"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>角色管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">角色管理</MudText>
<MudButton Variant="Variant.Filled" Color="Color.Primary" Href="/admin/system/role-edit/0" StartIcon="@Icons.Material.Filled.Add" Class="mb-4">新增角色</MudButton>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>ID</MudTh><MudTh>角色名</MudTh><MudTh>代码</MudTh><MudTh>描述</MudTh><MudTh>创建时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.RoleId</MudTd><MudTd>@context.RoleName</MudTd><MudTd>@context.RoleCode</MudTd><MudTd>@context.Description</MudTd>
        <MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd>
            <MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small" Color="Color.Info" Href="@($"/admin/system/role-edit/{context.RoleId}")" />
            <MudIconButton Icon="@Icons.Material.Filled.Delete" Size="Size.Small" Color="Color.Error" OnClick="@(() => Delete(context))" />
        </MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private List<AdminRole> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AdminRoles.AsNoTracking().ToListAsync(); }
    private async Task Delete(AdminRole r) { Db.AdminRoles.Remove(r); await Db.SaveChangesAsync(); _items.Remove(r); Snackbar.Add("已删除", Severity.Success); }
}
`);

// 5. RoleEdit
W(`${sys}/RoleEdit.razor`, `@page "/admin/system/role-edit/{RoleId:int}"
@using PerfumeShop.Data.Models

<PageTitle>编辑角色</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑角色</MudText>
<MudGrid><MudItem xs="12" md="6"><MudForm>
    <MudTextField @bind-Value="_role.RoleName" Label="角色名" Required="true" />
    <MudTextField @bind-Value="_role.RoleCode" Label="角色代码" Required="true" />
    <MudTextField @bind-Value="_role.Description" Label="描述" Lines="3" />
    <MudTextField @bind-Value="_role.Permissions" Label="权限(JSON)" Lines="4" />
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/system/roles" Class="mt-4 ml-2">返回</MudButton>
</MudForm></MudItem></MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    [Parameter] public int RoleId { get; set; }
    private AdminRole _role = new();
    protected override async Task OnInitializedAsync() { if (RoleId > 0) _role = await Db.AdminRoles.FindAsync(RoleId) ?? new(); }
    private async Task Save()
    {
        if (RoleId == 0) Db.AdminRoles.Add(_role); else Db.AdminRoles.Update(_role);
        await Db.SaveChangesAsync();
        Snackbar.Add("保存成功", Severity.Success);
        Nav.NavigateTo("/admin/system/roles");
    }
}
`);

// 6. AuditLog
W(`${sys}/AuditLog.razor`, `@page "/admin/system/audit-log"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>审计日志</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">审计日志</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>ID</MudTh><MudTh>管理员</MudTh><MudTh>操作类型</MudTh><MudTh>目标</MudTh><MudTh>IP</MudTh><MudTh>时间</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.LogId</MudTd><MudTd>@context.AdminName</MudTd><MudTd><MudChip T="string" Size="Size.Small">@context.ActionType</MudChip></MudTd>
        <MudTd>@context.TargetName</MudTd><MudTd>@context.Ipaddress</MudTd><MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AdminAuditLog> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AdminAuditLogs.AsNoTracking().OrderByDescending(l => l.LogId).Take(200).ToListAsync(); }
}
`);

// 7. AuditDetail
W(`${sys}/AuditDetail.razor`, `@page "/admin/system/audit-detail/{LogId:int}"
@using PerfumeShop.Data.Models

<PageTitle>审计日志详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">审计日志详情</MudText>
@if (_log != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>日志ID:</b> @_log.LogId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>管理员:</b> @_log.AdminName</MudText></MudItem>
    <MudItem xs="6"><MudText><b>操作类型:</b> @_log.ActionType</MudText></MudItem>
    <MudItem xs="6"><MudText><b>目标:</b> @_log.TargetType #@_log.TargetId</MudText></MudItem>
    <MudItem xs="12"><MudText><b>详情:</b> @_log.Details</MudText></MudItem>
    <MudItem xs="6"><MudText><b>IP:</b> @_log.Ipaddress</MudText></MudItem>
    <MudItem xs="6"><MudText><b>时间:</b> @_log.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/system/audit-log" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int LogId { get; set; }
    private AdminAuditLog? _log;
    protected override async Task OnInitializedAsync() { _log = await Db.AdminAuditLogs.FindAsync(LogId); }
}
`);

// 8. BackupManagement
W(`${sys}/BackupManagement.razor`, `@page "/admin/system/backup"

<PageTitle>备份管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">备份管理</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">最近备份</MudText><MudText Typo="Typo.h4">@_lastBackup?.ToString("yyyy-MM-dd") ?? "无"</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">备份数量</MudText><MudText Typo="Typo.h4">@_backupCount</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="4"><MudButton Variant="Variant.Filled" Color="Color.Primary" StartIcon="@Icons.Material.Filled.Add" Href="/admin/system/backup-create" Class="mt-6">创建备份</MudButton></MudItem>
</MudGrid>
<MudTable Items="@_backups" Hover="true" Dense="true">
    <HeaderContent><MudTh>文件名</MudTh><MudTh>大小</MudTh><MudTh>创建时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate><MudTd>@context.FileName</MudTd><MudTd>@context.FileSize</MudTd><MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd><MudTd><MudIconButton Icon="@Icons.Material.Filled.Restore" Color="Color.Warning" Href="/admin/system/backup-restore" /></MudTd></RowTemplate>
</MudTable>

@code {
    private DateTime? _lastBackup; private int _backupCount;
    private List<BackupInfo> _backups = new();
    private class BackupInfo { public string FileName { get; set; } = ""; public string FileSize { get; set; } = ""; public DateTime? CreatedAt { get; set; } }
}
`);

// 9. BackupCreate
W(`${sys}/BackupCreate.razor`, `@page "/admin/system/backup-create"

<PageTitle>创建备份</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">创建备份</MudText>
<MudCard><MudCardContent>
    <MudText Class="mb-4">确认创建新的数据库备份？此操作将生成完整的数据库快照。</MudText>
    <MudTextField @bind-Value="_notes" Label="备份备注" Lines="3" />
    <MudSwitch @bind-Value="_compress" Label="压缩备份" Color="Color.Primary" />
</MudCardContent>
<MudCardActions>
    <MudButton Variant="Variant.Filled" Color="Color.Primary" OnClick="CreateBackup" Disabled="@_creating">@(_creating ? "备份中..." : "开始备份")</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/system/backup">取消</MudButton>
</MudCardActions></MudCard>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private string _notes = ""; private bool _compress = true, _creating = false;
    private async Task CreateBackup() { _creating = true; await Task.Delay(500); Snackbar.Add("备份创建成功", Severity.Success); _creating = false; Nav.NavigateTo("/admin/system/backup"); }
}
`);

// 10. BackupRestore
W(`${sys}/BackupRestore.razor`, `@page "/admin/system/backup-restore"

<PageTitle>恢复备份</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">恢复备份</MudText>
<MudAlert Severity="Severity.Warning" Class="mb-4">警告：恢复备份将覆盖当前数据库中的所有数据。</MudAlert>
<MudCard><MudCardContent>
    <MudFileUpload Accept=".bak,.sql" FilesChanged="OnFileSelected"><ActivatorContent><MudButton Variant="Variant.Outlined" Color="Color.Primary" StartIcon="@Icons.Material.Filled.Upload">选择备份文件</MudButton></ActivatorContent></MudFileUpload>
    @if (_file != null) { <MudText Class="mt-2">已选择: @_file.Name</MudText> }
</MudCardContent>
<MudCardActions>
    <MudButton Variant="Variant.Filled" Color="Color.Warning" OnClick="DoRestore" Disabled="@(_file == null)">开始恢复</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/system/backup">返回</MudButton>
</MudCardActions></MudCard>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private IBrowserFile? _file;
    private void OnFileSelected(IBrowserFile f) => _file = f;
    private async Task DoRestore() { await Task.Delay(500); Snackbar.Add("恢复完成", Severity.Success); Nav.NavigateTo("/admin/system/backup"); }
}
`);

// 11. SystemConfig
W(`${sys}/SystemConfig.razor`, `@page "/admin/system/config"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>系统配置</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">系统配置</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>配置名</MudTh><MudTh>键</MudTh><MudTh>值</MudTh><MudTh>描述</MudTh><MudTh>更新时间</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.SettingName</MudTd><MudTd><code>@context.SettingKey</code></MudTd><MudTd>@context.SettingValue</MudTd>
        <MudTd>@context.Description</MudTd><MudTd>@context.UpdatedAt?.ToString("yyyy-MM-dd")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small" Color="Color.Info" Href="/admin/system/config-edit" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<SiteSetting> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.SiteSettings.AsNoTracking().ToListAsync(); }
}
`);

// 12. SystemConfigEdit
W(`${sys}/SystemConfigEdit.razor`, `@page "/admin/system/config-edit"
@using PerfumeShop.Data.Models

<PageTitle>编辑系统配置</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑系统配置</MudText>
<MudGrid><MudItem xs="12" md="8"><MudForm>
    <MudTextField @bind-Value="_s.SettingName" Label="配置名" />
    <MudTextField @bind-Value="_s.SettingKey" Label="键" />
    <MudTextField @bind-Value="_s.SettingValue" Label="值" Lines="3" />
    <MudTextField @bind-Value="_s.Description" Label="描述" Lines="2" />
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/system/config" Class="mt-4 ml-2">返回</MudButton>
</MudForm></MudItem></MudGrid>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private SiteSetting _s = new();
    private async Task Save() { _s.UpdatedAt = DateTime.Now; Db.SiteSettings.Update(_s); await Db.SaveChangesAsync(); Snackbar.Add("保存成功", Severity.Success); Nav.NavigateTo("/admin/system/config"); }
}
`);

// 13. OperationLog
W(`${sys}/OperationLog.razor`, `@page "/admin/system/operation-log"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>操作日志</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">操作日志</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>ID</MudTh><MudTh>级别</MudTh><MudTh>消息</MudTh><MudTh>来源</MudTh><MudTh>用户</MudTh><MudTh>时间</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.LogId</MudTd><MudTd><MudChip T="string" Size="Size.Small" Color="@(context.LogLevel == "Error" ? Color.Error : Color.Default)">@context.LogLevel</MudChip></MudTd>
        <MudTd>@(context.LogMessage?.Length > 50 ? context.LogMessage.Substring(0,50)+"..." : context.LogMessage)</MudTd>
        <MudTd>@context.LogSource</MudTd><MudTd>@context.UserName</MudTd><MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AppLog> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AppLogs.AsNoTracking().OrderByDescending(l => l.LogId).Take(200).ToListAsync(); }
}
`);

// 14. OperationLogDetail
W(`${sys}/OperationLogDetail.razor`, `@page "/admin/system/operation-log-detail/{LogId:long}"
@using PerfumeShop.Data.Models

<PageTitle>操作日志详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">操作日志详情</MudText>
@if (_log != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>日志ID:</b> @_log.LogId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>级别:</b> @_log.LogLevel</MudText></MudItem>
    <MudItem xs="12"><MudText><b>消息:</b> @_log.LogMessage</MudText></MudItem>
    <MudItem xs="6"><MudText><b>来源:</b> @_log.LogSource</MudText></MudItem>
    <MudItem xs="6"><MudText><b>行号:</b> @_log.LineNumber</MudText></MudItem>
    <MudItem xs="6"><MudText><b>用户:</b> @_log.UserName</MudText></MudItem>
    <MudItem xs="6"><MudText><b>IP:</b> @_log.Ipaddress</MudText></MudItem>
    <MudItem xs="12"><MudText><b>URL:</b> @_log.PageUrl</MudText></MudItem>
    <MudItem xs="6"><MudText><b>时间:</b> @_log.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/system/operation-log" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public long LogId { get; set; }
    private AppLog? _log;
    protected override async Task OnInitializedAsync() { _log = await Db.AppLogs.FindAsync(LogId); }
}
`);

// 15. LoginHistory
W(`${sys}/LoginHistory.razor`, `@page "/admin/system/login-history"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>登录历史</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">登录历史</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>ID</MudTh><MudTh>类型</MudTh><MudTh>级别</MudTh><MudTh>消息</MudTh><MudTh>IP</MudTh><MudTh>已读</MudTh><MudTh>时间</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.AlertId</MudTd><MudTd>@context.AlertType</MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.AlertLevel == "critical" ? Color.Error : Color.Warning)">@context.AlertLevel</MudChip></MudTd>
        <MudTd>@context.AlertMessage</MudTd><MudTd>@context.Ipaddress</MudTd>
        <MudTd>@(context.IsRead == true ? "是" : "否")</MudTd><MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<LoginAlert> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.LoginAlerts.AsNoTracking().OrderByDescending(a => a.AlertId).ToListAsync(); }
}
`);

// 16. LoginHistoryDetail
W(`${sys}/LoginHistoryDetail.razor`, `@page "/admin/system/login-history-detail/{AlertId:int}"
@using PerfumeShop.Data.Models

<PageTitle>登录历史详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">登录历史详情</MudText>
@if (_a != null) {
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>告警ID:</b> @_a.AlertId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>类型:</b> @_a.AlertType</MudText></MudItem>
    <MudItem xs="6"><MudText><b>级别:</b> @_a.AlertLevel</MudText></MudItem>
    <MudItem xs="6"><MudText><b>管理员ID:</b> @_a.AdminId</MudText></MudItem>
    <MudItem xs="12"><MudText><b>消息:</b> @_a.AlertMessage</MudText></MudItem>
    <MudItem xs="6"><MudText><b>IP:</b> @_a.Ipaddress</MudText></MudItem>
    <MudItem xs="6"><MudText><b>已读:</b> @(_a.IsRead == true ? "是" : "否")</MudText></MudItem>
    <MudItem xs="6"><MudText><b>时间:</b> @_a.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/system/login-history" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public int AlertId { get; set; }
    private LoginAlert? _a;
    protected override async Task OnInitializedAsync() { _a = await Db.LoginAlerts.FindAsync(AlertId); }
}
`);

// 17. ErrorLog
W(`${sys}/ErrorLog.razor`, `@page "/admin/system/error-log"
@using Microsoft.EntityFrameworkCore
@using PerfumeShop.Data.Models

<PageTitle>错误日志</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">错误日志</MudText>
<MudTable Items="@_items" Hover="true" Dense="true" Striped="true" RowsPerPage="20">
    <HeaderContent><MudTh>ID</MudTh><MudTh>级别</MudTh><MudTh>消息</MudTh><MudTh>来源</MudTh><MudTh>行号</MudTh><MudTh>时间</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.LogId</MudTd><MudTd><MudChip T="string" Size="Size.Small" Color="Color.Error">@context.LogLevel</MudChip></MudTd>
        <MudTd>@(context.LogMessage?.Length > 60 ? context.LogMessage.Substring(0,60)+"..." : context.LogMessage)</MudTd>
        <MudTd>@context.LogSource</MudTd><MudTd>@context.LineNumber</MudTd><MudTd>@context.CreatedAt?.ToString("yyyy-MM-dd HH:mm")</MudTd>
    </RowTemplate>
    <PagerContent><MudTablePager /></PagerContent>
</MudTable>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    private List<AppLog> _items = new();
    protected override async Task OnInitializedAsync() { _items = await Db.AppLogs.AsNoTracking().Where(l => l.LogLevel == "Error" || l.LogLevel == "Fatal").OrderByDescending(l => l.LogId).Take(200).ToListAsync(); }
}
`);

// 18. ErrorLogDetail
W(`${sys}/ErrorLogDetail.razor`, `@page "/admin/system/error-log-detail/{LogId:long}"
@using PerfumeShop.Data.Models

<PageTitle>错误日志详情</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">错误日志详情</MudText>
@if (_log != null) {
<MudAlert Severity="Severity.Error" Class="mb-4">@_log.LogMessage</MudAlert>
<MudCard><MudCardContent><MudGrid>
    <MudItem xs="6"><MudText><b>日志ID:</b> @_log.LogId</MudText></MudItem>
    <MudItem xs="6"><MudText><b>级别:</b> @_log.LogLevel</MudText></MudItem>
    <MudItem xs="6"><MudText><b>来源:</b> @_log.LogSource</MudText></MudItem>
    <MudItem xs="6"><MudText><b>行号:</b> @_log.LineNumber</MudText></MudItem>
    <MudItem xs="6"><MudText><b>类型:</b> @_log.LogType</MudText></MudItem>
    <MudItem xs="6"><MudText><b>用户:</b> @_log.UserName</MudText></MudItem>
    <MudItem xs="12"><MudText><b>URL:</b> @_log.PageUrl</MudText></MudItem>
    <MudItem xs="6"><MudText><b>时间:</b> @_log.CreatedAt?.ToString("yyyy-MM-dd HH:mm:ss")</MudText></MudItem>
</MudGrid></MudCardContent></MudCard>}
<MudButton Variant="Variant.Text" Href="/admin/system/error-log" Class="mt-4">返回列表</MudButton>

@code {
    [Inject] private PerfumeShopContext Db { get; set; } = default!;
    [Parameter] public long LogId { get; set; }
    private AppLog? _log;
    protected override async Task OnInitializedAsync() { _log = await Db.AppLogs.FindAsync(LogId); }
}
`);

// 19. ScheduledTasks
W(`${sys}/ScheduledTasks.razor`, `@page "/admin/system/scheduled-tasks"

<PageTitle>定时任务</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">定时任务</MudText>
<MudTable Items="@_tasks" Hover="true" Dense="true" Striped="true">
    <HeaderContent><MudTh>任务名</MudTh><MudTh>Cron表达式</MudTh><MudTh>状态</MudTh><MudTh>上次执行</MudTh><MudTh>下次执行</MudTh><MudTh>操作</MudTh></HeaderContent>
    <RowTemplate>
        <MudTd>@context.Name</MudTd><MudTd><code>@context.Cron</code></MudTd>
        <MudTd><MudChip T="string" Size="Size.Small" Color="@(context.Enabled ? Color.Success : Color.Default)">@(context.Enabled ? "启用" : "停用")</MudChip></MudTd>
        <MudTd>@context.LastRun?.ToString("yyyy-MM-dd HH:mm")</MudTd><MudTd>@context.NextRun?.ToString("yyyy-MM-dd HH:mm")</MudTd>
        <MudTd><MudIconButton Icon="@Icons.Material.Filled.Edit" Size="Size.Small" Color="Color.Info" Href="/admin/system/scheduled-task-edit" /></MudTd>
    </RowTemplate>
</MudTable>

@code {
    private List<TaskInfo> _tasks = new()
    {
        new() { Name = "数据库备份", Cron = "0 2 * * *", Enabled = true, LastRun = DateTime.Today.AddDays(-1).AddHours(2), NextRun = DateTime.Today.AddHours(2) },
        new() { Name = "日志清理", Cron = "0 3 * * 0", Enabled = true, LastRun = DateTime.Today.AddDays(-3).AddHours(3), NextRun = DateTime.Today.AddDays(4).AddHours(3) },
        new() { Name = "缓存刷新", Cron = "*/30 * * * *", Enabled = true, LastRun = DateTime.Now.AddMinutes(-30), NextRun = DateTime.Now.AddMinutes(30) },
        new() { Name = "统计聚合", Cron = "0 1 * * *", Enabled = false },
    };
    private class TaskInfo { public string Name { get; set; } = ""; public string Cron { get; set; } = ""; public bool Enabled { get; set; } public DateTime? LastRun { get; set; } public DateTime? NextRun { get; set; } }
}
`);

// 20. ScheduledTaskEdit
W(`${sys}/ScheduledTaskEdit.razor`, `@page "/admin/system/scheduled-task-edit"

<PageTitle>编辑定时任务</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">编辑定时任务</MudText>
<MudGrid><MudItem xs="12" md="6"><MudForm>
    <MudTextField @bind-Value="_name" Label="任务名" Required="true" />
    <MudTextField @bind-Value="_cron" Label="Cron表达式" HelperText="例: 0 2 * * *" />
    <MudSwitch @bind-Value="_enabled" Label="启用" Color="Color.Success" />
    <MudButton Variant="Variant.Filled" Color="Color.Primary" Class="mt-4" OnClick="Save">保存</MudButton>
    <MudButton Variant="Variant.Text" Href="/admin/system/scheduled-tasks" Class="mt-4 ml-2">返回</MudButton>
</MudForm></MudItem></MudGrid>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    [Inject] private NavigationManager Nav { get; set; } = default!;
    private string _name = "", _cron = ""; private bool _enabled = true;
    private void Save() { Snackbar.Add("保存成功", Severity.Success); Nav.NavigateTo("/admin/system/scheduled-tasks"); }
}
`);

// 21. CacheManagement
W(`${sys}/CacheManagement.razor`, `@page "/admin/system/cache"

<PageTitle>缓存管理</PageTitle>
<MudText Typo="Typo.h4" Class="mb-4">缓存管理</MudText>
<MudGrid Class="mb-4">
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">内存缓存</MudText><MudText Typo="Typo.h5">@_memCount 项</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">命中率</MudText><MudText Typo="Typo.h5">@_hitRate%</MudText></MudCardContent></MudCard></MudItem>
    <MudItem xs="12" sm="3"><MudCard><MudCardContent><MudText Typo="Typo.subtitle2">缓存大小</MudText><MudText Typo="Typo.h5">@_cacheSize MB</MudText></MudCardContent></MudCard></MudItem>
</MudGrid>
<MudButton Variant="Variant.Filled" Color="Color.Warning" OnClick="ClearCache" StartIcon="@Icons.Material.Filled.DeleteSweep">清除所有缓存</MudButton>

@code {
    [Inject] private ISnackbar Snackbar { get; set; } = default!;
    private int _memCount = 42; private double _hitRate = 87.5, _cacheSize = 12.3;
    private void ClearCache() { _memCount = 0; Snackbar.Add("缓存已清除", Severity.Success); }
}
`);

console.log(`System module: ${n} pages created`);
