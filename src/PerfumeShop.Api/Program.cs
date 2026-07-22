using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.AI;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;
using PerfumeShop.Data.Services;
using PerfumeShop.Shared;
using PerfumeShop.Shared.Auth;
using PerfumeShop.Shared.Services;

var builder = WebApplication.CreateBuilder(args);

// V19.1: Windows Service 支持（开机自启 + 崩溃自动恢复）
builder.Host.UseWindowsService();

// Add services to the container.
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register EF Core DbContext with SQL Server
// V19: 连接弹性配置 — 对齐 V18 connection.asp OpenConnection() 3次重试+健康检查
builder.Services.AddDbContext<PerfumeShopContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("PerfumeShop"),
        sqlOptions => sqlOptions.EnableRetryOnFailure(
            maxRetryCount: 3,
            maxRetryDelay: TimeSpan.FromMilliseconds(500),
            errorNumbersToAdd: null)));

// ========== Repository Pattern DI 注册 ==========
// 泛型仓储 (开放泛型注册)
builder.Services.AddScoped(typeof(IRepository<>), typeof(Repository<>));

// 特化仓储接口
builder.Services.AddScoped<IProductRepository, ProductRepository>();
builder.Services.AddScoped<IOrderRepository, OrderRepository>();
builder.Services.AddScoped<IUserRepository, UserRepository>();

// ========== Business Services DI 注册 ==========
builder.Services.AddMemoryCache(o => o.SizeLimit = 1024);
builder.Services.AddScoped<ICostEngine, CostEngine>();
builder.Services.AddScoped<IPromotionEngine, PromotionEngine>();
builder.Services.AddScoped<IPaymentHandler, PaymentHandler>();
builder.Services.AddScoped<IProductTypeService, ProductTypeService>();
builder.Services.AddScoped<IRecommendationEngine, RecommendationEngine>();

// ========== AI Services DI 注册 ==========
builder.Services.AddSingleton<ISentimentAnalyzer, SentimentAnalyzer>();
builder.Services.AddSingleton<IChatbotEngine, ChatbotEngine>();
builder.Services.AddSingleton<IFragranceMatcher, FragranceMatcher>();

// ========== Marketing & Extended Services DI 注册 ==========
builder.Services.AddScoped<IFlashSaleService, FlashSaleService>();
builder.Services.AddScoped<IGroupBuyService, GroupBuyService>();
builder.Services.AddScoped<ISubscriptionService, SubscriptionService>();
builder.Services.AddScoped<IPointsService, PointsService>();
builder.Services.AddScoped<IPointsEngine, PointsEngine>();
builder.Services.AddScoped<ICartService, CartService>();

// V18.3 国际化服务 (i18n)
builder.Services.AddLocaleService();

// V19 M3-A: Session (购物车合并需要)
builder.Services.AddSession();

// V19 M2-A: 共享基础设施服务
var connStr = builder.Configuration.GetConnectionString("PerfumeShop") ?? "";
builder.Services.AddCacheService(builder.Configuration);
builder.Services.AddAuditService(connStr);
builder.Services.AddTrackingService(connStr);
var smtpOpts = new SmtpOptions();
builder.Configuration.GetSection("Smtp").Bind(smtpOpts);
builder.Services.AddEmailService(smtpOpts, connectionString: connStr);
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
        options.Events.OnRedirectToLogin = ctx =>
        {
            // API/AJAX 请求返回 401 JSON，浏览器请求重定向到登录页
            if (ctx.Request.Path.StartsWithSegments("/api"))
            {
                ctx.Response.StatusCode = 401;
                return Task.CompletedTask;
            }
            var redirectUri = $"/login?returnUrl={Uri.EscapeDataString(ctx.Request.Path + ctx.Request.QueryString)}";
            ctx.Response.Redirect(redirectUri);
            return Task.CompletedTask;
        };
        options.Events.OnRedirectToAccessDenied = ctx =>
        {
            if (ctx.Request.Path.StartsWithSegments("/api"))
            {
                ctx.Response.StatusCode = 403;
                return Task.CompletedTask;
            }
            ctx.Response.Redirect("/login?returnUrl=" + Uri.EscapeDataString(ctx.Request.Path));
            return Task.CompletedTask;
        };
    });
builder.Services.AddAuthorization();
builder.Services.Configure<AuthBridgeOptions>(builder.Configuration.GetSection(AuthBridgeOptions.SectionName));
builder.Services.AddScoped<IAuthTokenStore, DbAuthTokenStore>();

// Add controllers + Razor Pages support
builder.Services.AddControllers();
builder.Services.AddRazorPages();

// V19: HttpContextAccessor（CacheService需要用于X-Cache响应头）
builder.Services.AddHttpContextAccessor();

// V19: 缓存预热服务（对齐 V18 CM_Warmup）
builder.Services.AddHostedService<PerfumeShop.Api.Services.CacheWarmupService>();

// ========== Performance Optimizations ==========
// Response compression (gzip/brotli)
builder.Services.AddResponseCompression(options =>
{
    options.EnableForHttps = true;
});

// Output caching for Razor Pages and API responses
builder.Services.AddOutputCache(options =>
{
    // 首页缓存 2 分钟
    options.AddPolicy("HomePage", b => b.Expire(TimeSpan.FromMinutes(2)));
    // 商品列表缓存 5 分钟
    options.AddPolicy("ProductList", b => b.Expire(TimeSpan.FromMinutes(5)));
    // 商品详情缓存 10 分钟
    options.AddPolicy("ProductDetail", b => b.Expire(TimeSpan.FromMinutes(10)));
    // 营销页面缓存 3 分钟（秒杀/拼团/订阅/社区）
    options.AddPolicy("MarketingPage", b => b.Expire(TimeSpan.FromMinutes(3)));
    // 静态内容缓存 30 分钟
    options.AddPolicy("StaticContent", b => b.Expire(TimeSpan.FromMinutes(30)));
});

var app = builder.Build();

// Middleware pipeline

// V19: 请求追踪（X-Request-ID + X-Response-Time + X-Server）— 必须最先注册
app.UseRequestTracking();

app.UseResponseCompression();

// V19: 缓存响应头中间件（X-Cache: HIT/MISS/BYPASS）
app.UseCacheHeaders();

// Serve static files (images, CSS, JS) from wwwroot
// V19 M2-D: 静态文件长缓存（配合 ?v=19.0 查询串实现版本化失效）
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers.Append("Cache-Control", "public, max-age=31536000, immutable");
    }
});

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// Health check endpoint — verifies database connectivity
app.MapGet("/api/health", async (PerfumeShopContext db) =>
{
    try
    {
        var productCount = await db.Products.CountAsync();
        var userCount = await db.Users.CountAsync();
        return Results.Ok(new
        {
            status = "ok",
            service = "PerfumeShop API v2",
            version = "0.1.0",
            database = "connected",
            products = productCount,
            users = userCount,
            timestamp = DateTime.UtcNow
        });
    }
    catch (Exception ex)
    {
        return Results.Ok(new
        {
            status = "degraded",
            service = "PerfumeShop API v2",
            database = "disconnected",
            error = ex.Message
        });
    }
});

app.UseOutputCache();

// V18.3 结构化日志中间件
app.UseStructuredLogging();

// V19 M3-A: Session (购物车合并需要)
app.UseSession();

// V19 M2: Authentication + AuthBridge（必须在 UseApiKeyAuth 之前）
app.UseAuthentication();
app.UseAuthorization();
app.UseAuthBridge();

// V18.3 API Key + HMAC 认证中间件（双模式：Cookie Bridge 优先，API Key 回退）
app.UseApiKeyAuth();

// V18.3 速率限制中间件（令牌桶算法 + CSRF 防护，60req/60s）
app.UseRateLimiter(maxRequests: 60, windowSeconds: 60);

app.MapControllers();
app.MapRazorPages();

// V18.3 系统指标端点
app.MapMetricsEndpoint();

app.Run();

// Make Program accessible to WebApplicationFactory in test projects
public partial class Program { }
