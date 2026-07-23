using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using Microsoft.EntityFrameworkCore;
using MudBlazor.Services;
using PerfumeShop.Admin.Components;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;
using PerfumeShop.Data.Services;
using PerfumeShop.Shared;
using PerfumeShop.Shared.Auth;
using PerfumeShop.Shared.Services;
using PerfumeShop.Admin.Services;
using PerfumeShop.Admin.Authorization;
using PerfumeShop.Admin.Middleware;
using Microsoft.AspNetCore.Authorization;

var builder = WebApplication.CreateBuilder(args);

// V19.1: Windows Service 支持（开机自启 + 崩溃自动恢复）
builder.Host.UseWindowsService();

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Output caching for admin pages
builder.Services.AddOutputCache(options =>
{
    options.AddPolicy("AdminPage", b => b.Expire(TimeSpan.FromMinutes(5)));
});

// MudBlazor
builder.Services.AddMudServices();

// HttpContextAccessor (Login 页面需要)
builder.Services.AddHttpContextAccessor();

// EF Core DbContext — V19: 连接弹性配置（对齐 V18 connection.asp 3次重试）
builder.Services.AddDbContext<PerfumeShopContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("PerfumeShop"),
        sqlOptions => sqlOptions.EnableRetryOnFailure(
            maxRetryCount: 3,
            maxRetryDelay: TimeSpan.FromMilliseconds(500),
            errorNumbersToAdd: null)));

// Repository Pattern
builder.Services.AddScoped(typeof(IRepository<>), typeof(Repository<>));
builder.Services.AddScoped<IProductRepository, ProductRepository>();
builder.Services.AddScoped<IOrderRepository, OrderRepository>();
builder.Services.AddScoped<IUserRepository, UserRepository>();
builder.Services.AddScoped<PurchaseRepository>();

// M5-A: System / Inventory / Logistics / Finance Repositories
builder.Services.AddScoped<SystemRepository>();
builder.Services.AddScoped<InventoryRepository>();
builder.Services.AddScoped<LogisticsRepository>();
builder.Services.AddScoped<FinanceRepository>();

// V19.8: Operation / Production / SemiFinished / TechCenter Repositories
builder.Services.AddScoped<OperationRepository>();
builder.Services.AddScoped<ProductionRepository>();
builder.Services.AddScoped<SemiFinishedRepository>();
builder.Services.AddScoped<TechCenterRepository>();

// V19.8: Data Seeder for test data
builder.Services.AddScoped<DataSeeder>();

// V19 M1: 安全基座（统一当前用户 / 操作审计 / 数据驱动 RBAC / 登录安全）
builder.Services.AddScoped<ICurrentUserAccessor, CurrentUserAccessor>();
builder.Services.AddScoped<IAdminAuditor, AdminAuditor>();
builder.Services.AddScoped<IPermissionService, PermissionService>();
builder.Services.AddScoped<ILoginSecurityService, LoginSecurityService>();
builder.Services.AddSingleton<IAuthorizationHandler, ModuleAccessHandler>();

// Business Services
builder.Services.AddMemoryCache();
builder.Services.AddScoped<ExpenseAllocationService>(); // V20: 费用分摊引擎
builder.Services.AddScoped<ICostEngine, CostEngine>();
builder.Services.AddScoped<IPromotionEngine, PromotionEngine>();
builder.Services.AddScoped<IPaymentHandler, PaymentHandler>();
builder.Services.AddScoped<IProductTypeService, ProductTypeService>();
builder.Services.AddScoped<IRecommendationEngine, RecommendationEngine>();

// Marketing & Extended Services
builder.Services.AddScoped<IFlashSaleService, FlashSaleService>();
builder.Services.AddScoped<IGroupBuyService, GroupBuyService>();
builder.Services.AddScoped<ISubscriptionService, SubscriptionService>();
builder.Services.AddScoped<IPointsService, PointsService>();
builder.Services.AddScoped<IPointsEngine, PointsEngine>();
builder.Services.AddScoped<ICartService, CartService>();

// V18.3 国际化服务 (i18n)
builder.Services.AddLocaleService();

// V19 M2-A: 共享基础设施服务
var connStr = builder.Configuration.GetConnectionString("PerfumeShop") ?? "";
builder.Services.AddCacheService(builder.Configuration);
builder.Services.AddAuditService(connStr);
builder.Services.AddTrackingService(connStr);
var smtpOpts = new SmtpOptions();
builder.Configuration.GetSection("Smtp").Bind(smtpOpts);
builder.Services.AddEmailService(smtpOpts);
builder.Services.AddSingleton<IStorageProvider>(sp =>
{
    var env = sp.GetRequiredService<IWebHostEnvironment>();
    var uploadsPath = Path.Combine(env.ContentRootPath, "wwwroot", "uploads");
    if (!Directory.Exists(uploadsPath)) Directory.CreateDirectory(uploadsPath);
    return new LocalStorageProvider(uploadsPath);
});
builder.Services.AddScoped<IUploadService>(sp =>
    new UploadService(sp.GetRequiredService<IStorageProvider>()));

// ========== V19 M2: Cookie Authentication + AuthBridge ==========
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.Cookie.Name = "V19_AUTH";
        options.Cookie.HttpOnly = true;
        options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
        options.ExpireTimeSpan = TimeSpan.FromHours(8);
        options.SlidingExpiration = true;
        options.LoginPath = "/login";
        options.LogoutPath = "/logout";
        options.AccessDeniedPath = "/login";
        // SignalR WebSocket 连接不重定向，直接返回状态码
        options.Events.OnRedirectToLogin = ctx =>
        {
            if (ctx.Request.Path.StartsWithSegments("/_blazor") ||
                ctx.Request.Path.StartsWithSegments("/_framework"))
            {
                ctx.Response.StatusCode = 401;
                return Task.CompletedTask;
            }
            ctx.Response.Redirect(ctx.RedirectUri);
            return Task.CompletedTask;
        };
        options.Events.OnRedirectToAccessDenied = ctx =>
        {
            if (ctx.Request.Path.StartsWithSegments("/_blazor") ||
                ctx.Request.Path.StartsWithSegments("/_framework"))
            {
                ctx.Response.StatusCode = 403;
                return Task.CompletedTask;
            }
            ctx.Response.Redirect(ctx.RedirectUri);
            return Task.CompletedTask;
        };
    });
builder.Services.AddAuthorization(options =>
{
    // V19 M1: 按模块注册数据驱动授权策略（Module.system / Module.finance / ...）
    foreach (var m in ModulePolicies.Modules)
        options.AddPolicy(ModulePolicies.PolicyName(m), p =>
            p.RequireAuthenticatedUser().AddRequirements(new ModuleAccessRequirement(m)));
});
builder.Services.Configure<AuthBridgeOptions>(builder.Configuration.GetSection(AuthBridgeOptions.SectionName));
builder.Services.AddScoped<IAuthTokenStore, DbAuthTokenStore>();

var app = builder.Build();

// V19 M2: Authentication MUST come before any response-writing middleware
app.UseAuthentication();
// V19 M1: IP 黑名单运行时拦截（认证后、授权前）
app.UseIpBlacklist();
app.UseAuthorization();
app.UseAuthBridge();

// V19: 请求追踪中间件
app.UseRequestTracking();

// V19 M2-D: 静态文件长缓存（Blazor 资产 + CSS/JS 均启用 immutable 长缓存，
// 配合 _Layout.cshtml 中的 ?v=19.0 查询串实现版本化失效）
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers.Append("Cache-Control", "public, max-age=31536000, immutable");
    }
});

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}

app.UseOutputCache();
app.UseAntiforgery();

// V19: 缓存响应头中间件
app.UseCacheHeaders();

// V18.3 结构化日志中间件
app.UseStructuredLogging();

// V20 对账CSV导出端点
app.MapGet("/api/admin/reconciliation/export", async (HttpContext ctx, PerfumeShopContext db) => {
    var sb = new System.Text.StringBuilder();
    sb.AppendLine("ID,对账日期,订单号,订单金额,支付金额,差异,状态,处理说明");
    var all = await db.ReconciliationLogs.AsNoTracking().OrderByDescending(r => r.LogId).Take(500).ToListAsync();
    foreach(var r in all){
        var st = r.Status switch {"Matched"=>"匹配","ShortPay"=>"短款","OverPay"=>"长款","Missing"=>"未达","Resolved"=>"已解决",_=>r.Status??"-"};
        sb.AppendLine($"{r.LogId},{r.ReconcileDate:yyyy-MM-dd},\"{r.OrderNo}\",{r.OrderAmount:F2},{r.PaymentAmount:F2},{r.Difference:F2},\"{st}\",\"{r.Resolution}\"");
    }
    var csv = System.Text.Encoding.UTF8.GetBytes(sb.ToString());
    ctx.Response.ContentType = "text/csv; charset=utf-8";
    ctx.Response.Headers.ContentDisposition = $"attachment; filename=reconciliation_{DateTime.Now:yyyyMMdd}.csv";
    await ctx.Response.Body.WriteAsync(System.Text.Encoding.UTF8.Preamble.ToArray());
    await ctx.Response.Body.WriteAsync(csv);
});

// V19 M1: 备份文件下载（鉴权 + 文件名白名单，防目录穿越）
app.MapGet("/admin/system/backup/download", async (HttpContext ctx, string file) =>
{
    if (string.IsNullOrWhiteSpace(file) || file.Contains("..") || file.Contains('/') || file.Contains('\\'))
    { ctx.Response.StatusCode = 400; return; }
    string? dir = null;
    foreach (var start in new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory })
    {
        var d = new DirectoryInfo(start);
        for (int i = 0; i < 8 && d != null; i++, d = d.Parent)
        {
            var candidate = Path.Combine(d.FullName, "database", "backups");
            if (Directory.Exists(candidate)) { dir = candidate; break; }
        }
        if (dir != null) break;
    }
    var match = dir != null ? Directory.GetFiles(dir, file, SearchOption.AllDirectories).FirstOrDefault() : null;
    if (match == null || !File.Exists(match)) { ctx.Response.StatusCode = 404; return; }
    ctx.Response.ContentType = "application/octet-stream";
    ctx.Response.Headers.ContentDisposition = $"attachment; filename=\"{Path.GetFileName(match)}\"";
    await ctx.Response.SendFileAsync(match);
}).RequireAuthorization(ModulePolicies.PolicyName("system"));

// V19: 登出端点——清除认证 Cookie 后回到登录页（修复 LogoutPath 无实现导致登出无效）
app.MapGet("/logout", async (HttpContext ctx) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    return Results.Redirect("/login");
});

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

// V18.3 系统指标端点
app.MapMetricsEndpoint();

app.Run();
