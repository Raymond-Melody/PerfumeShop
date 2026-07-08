using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register EF Core DbContext with SQL Server
builder.Services.AddDbContext<PerfumeShopContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("PerfumeShop")));

// Add controllers support (for future REST API endpoints)
builder.Services.AddControllers();

var app = builder.Build();

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

app.MapControllers();

app.Run();
