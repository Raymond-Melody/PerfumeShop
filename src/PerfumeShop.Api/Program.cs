using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;
using PerfumeShop.Data.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register EF Core DbContext with SQL Server
builder.Services.AddDbContext<PerfumeShopContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("PerfumeShop")));

// ========== Repository Pattern DI 注册 ==========
// 泛型仓储 (开放泛型注册)
builder.Services.AddScoped(typeof(IRepository<>), typeof(Repository<>));

// 特化仓储接口
builder.Services.AddScoped<IProductRepository, ProductRepository>();
builder.Services.AddScoped<IOrderRepository, OrderRepository>();
builder.Services.AddScoped<IUserRepository, UserRepository>();

// ========== Business Services DI 注册 ==========
builder.Services.AddScoped<ICostEngine, CostEngine>();
builder.Services.AddScoped<IPromotionEngine, PromotionEngine>();
builder.Services.AddScoped<IPaymentHandler, PaymentHandler>();
builder.Services.AddScoped<IProductTypeService, ProductTypeService>();
builder.Services.AddScoped<IRecommendationEngine, RecommendationEngine>();

// ========== Marketing & Extended Services DI 注册 ==========
builder.Services.AddScoped<IFlashSaleService, FlashSaleService>();
builder.Services.AddScoped<IGroupBuyService, GroupBuyService>();
builder.Services.AddScoped<ISubscriptionService, SubscriptionService>();
builder.Services.AddScoped<IPointsService, PointsService>();
builder.Services.AddScoped<ICartService, CartService>();

// Add controllers + Razor Pages support
builder.Services.AddControllers();
builder.Services.AddRazorPages();

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
app.UseResponseCompression();

// Serve static files (images, CSS, JS) from wwwroot
app.UseStaticFiles();

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

app.MapControllers();
app.MapRazorPages();

app.Run();
