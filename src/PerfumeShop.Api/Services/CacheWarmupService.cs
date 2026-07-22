using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using PerfumeShop.Data.Models;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.Api.Services;

/// <summary>
/// V19 缓存预热服务 — 对齐 V18 cache_manager.asp CM_Warmup()
/// 应用启动时预加载热门数据到 L1 内存缓存
/// </summary>
public class CacheWarmupService : IHostedService
{
    private readonly IServiceProvider _services;
    private readonly ILogger<CacheWarmupService> _logger;

    public CacheWarmupService(IServiceProvider services, ILogger<CacheWarmupService> logger)
    {
        _services = services;
        _logger = logger;
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("=== V19 启动健康检查 + 缓存预热开始 ===");

        // 0. 数据库连接健康检查（对齐 V18 OpenConnection() SELECT 1 验证）
        var dbReady = false;
        for (int retry = 0; retry < 3 && !dbReady && !cancellationToken.IsCancellationRequested; retry++)
        {
            try
            {
                using var healthScope = _services.CreateScope();
                var healthDb = healthScope.ServiceProvider.GetRequiredService<PerfumeShopContext>();
                await healthDb.Database.CanConnectAsync(cancellationToken);
                dbReady = true;
                _logger.LogInformation("  数据库健康检查通过 (SELECT 1)");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "  数据库健康检查重试 {Retry}/3", retry + 1);
                if (retry < 2)
                    await Task.Delay(500, cancellationToken);
            }
        }

        if (!dbReady)
        {
            _logger.LogError("! 数据库健康检查失败（3次重试后仍不可用）— 服务启动但数据库不可达");
            return;
        }

        try
        {
            using var scope = _services.CreateScope();
            var cache = scope.ServiceProvider.GetRequiredService<ICacheService>();
            var db = scope.ServiceProvider.GetRequiredService<PerfumeShopContext>();

            // 1. 预热首页商品分类列表（5分钟缓存）
            var categories = await Task.Run(() =>
                db.Products
                    .Where(p => p.IsActive == true)
                    .Select(p => p.Category)
                    .Distinct()
                    .OrderBy(c => c)
                    .ToList(), cancellationToken);

            if (categories.Count > 0)
            {
                await cache.SetAsync("home_categories", categories,
                    TimeSpan.FromMinutes(5));
                _logger.LogInformation("  预热首页分类: {Count} 个", categories.Count);
            }

            // 2. 预热活跃产品类型配置（10分钟缓存）
            var productTypes = await Task.Run(() =>
                db.ProductTypeConfigs
                    .Where(pt => pt.IsActive == true)
                    .OrderBy(pt => pt.DisplayOrder)
                    .ToList(), cancellationToken);

            if (productTypes.Count > 0)
            {
                await cache.SetAsync("active_product_types", productTypes,
                    TimeSpan.FromMinutes(10));
                _logger.LogInformation("  预热产品类型: {Count} 个", productTypes.Count);
            }

            // 3. 预热最新10个活跃商品（2分钟缓存）
            var hotProducts = await Task.Run(() =>
                db.Products
                    .Where(p => p.IsActive == true)
                    .OrderByDescending(p => p.CreatedAt)
                    .Take(10)
                    .ToList(), cancellationToken);

            if (hotProducts.Count > 0)
            {
                await cache.SetAsync("home_hot_products", hotProducts,
                    TimeSpan.FromMinutes(2));
                _logger.LogInformation("  预热热门商品: {Count} 个", hotProducts.Count);
            }

            _logger.LogInformation("=== V19 缓存预热完成 ({Count} 键) ===", 3);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "V19 缓存预热失败（不影响服务启动）");
        }
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }
}
