using Microsoft.EntityFrameworkCore;
using MudBlazor.Services;
using PerfumeShop.Admin.Components;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;
using PerfumeShop.Data.Services;

var builder = WebApplication.CreateBuilder(args);

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

// EF Core DbContext
builder.Services.AddDbContext<PerfumeShopContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("PerfumeShop")));

// Repository Pattern
builder.Services.AddScoped(typeof(IRepository<>), typeof(Repository<>));
builder.Services.AddScoped<IProductRepository, ProductRepository>();
builder.Services.AddScoped<IOrderRepository, OrderRepository>();
builder.Services.AddScoped<IUserRepository, UserRepository>();

// Business Services
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
builder.Services.AddScoped<ICartService, CartService>();

var app = builder.Build();

// Cache Blazor framework assets for 1 year (content-hashed)
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        if (ctx.File.Name.StartsWith("blazor") || ctx.File.Name.EndsWith(".wasm") || ctx.File.Name.EndsWith(".dll"))
            ctx.Context.Response.Headers.CacheControl = "public,max-age=31536000,immutable";
    }
});

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}

app.UseOutputCache();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
