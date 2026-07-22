using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// RecommendationEngine V18 完整算法单元测试
/// 使用 EngineTestContext，每个测试创建独立 InMemory DB
/// </summary>
public class RecommendationEngineV18Tests : IDisposable
{
    private readonly EngineTestContext _db;
    private readonly RecommendationEngine _engine;

    public RecommendationEngineV18Tests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"RecV18_{Guid.NewGuid()}")
            .Options;
        _db = new EngineTestContext(options);
        _engine = new RecommendationEngine(_db);
    }

    public void Dispose()
    {
        _db.Database.EnsureDeleted();
        _db.Dispose();
    }

    // ==================== 辅助 ====================

    private async Task SeedProductAsync(int id, string name, string category, string? productType = null,
        bool active = true, DateTime? createdAt = null)
    {
        _db.Products.Add(new Product
        {
            ProductId = id, ProductName = name, Category = category,
            ProductType = productType ?? category,
            IsActive = active, CreatedAt = createdAt ?? DateTime.Now.AddDays(-30)
        });
        await _db.SaveChangesAsync();
    }

    private async Task SeedOrderAsync(int orderId, int userId, string status = "Paid")
    {
        _db.Orders.Add(new Order
        {
            OrderId = orderId, OrderNo = $"ORD{orderId}", UserId = userId,
            TotalAmount = 100, Status = status,
            CreatedAt = DateTime.Now.AddDays(-5), UpdatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();
    }

    private async Task SeedOrderDetailAsync(int orderId, int productId, int qty = 1)
    {
        _db.OrderDetails.Add(new OrderDetail
        {
            OrderId = orderId, ProductId = productId, Quantity = qty,
            UnitPrice = 100m, Subtotal = 100m * qty, ProductName = $"P{productId}"
        });
        await _db.SaveChangesAsync();
    }

    // ==================== 1. GetPersonalizedAsync ====================

    /// <summary>协同过滤 — 相似用户购买的商品应被推荐</summary>
    [Fact]
    public async Task Personalized_CollaborativeFilter_RecommendsSimilarUserProducts()
    {
        // 商品: 1,2,3,4,5（同类）
        for (int i = 1; i <= 5; i++)
            await SeedProductAsync(i, $"Perfume{i}", "女士香水", "香水");

        // 用户101 购买 产品1, 产品2
        _db.Users.Add(new User { UserId = 101, Username = "u101", Email = "u101@t.com", Password = "h" });
        await _db.SaveChangesAsync();
        await SeedOrderAsync(1, 101);
        await SeedOrderDetailAsync(1, 1);
        await SeedOrderDetailAsync(1, 2);

        // 用户102 购买 产品1, 产品2, 产品3（与101相似，多买了3）
        _db.Users.Add(new User { UserId = 102, Username = "u102", Email = "u102@t.com", Password = "h" });
        await _db.SaveChangesAsync();
        await SeedOrderAsync(2, 102);
        await SeedOrderDetailAsync(2, 1);
        await SeedOrderDetailAsync(2, 2);
        await SeedOrderDetailAsync(2, 3);

        var result = (await _engine.GetPersonalizedAsync(101, 6)).ToList();

        // 产品3 应该被推荐（相似用户102买过，101没买过）
        Assert.Contains(result, r => r.ProductId == 3);
    }

    /// <summary>个性化推荐 — 排除用户已购买商品</summary>
    [Fact]
    public async Task Personalized_ExcludesPurchasedProducts()
    {
        for (int i = 1; i <= 3; i++)
            await SeedProductAsync(i, $"Perfume{i}", "女士香水", "香水");

        _db.Users.Add(new User { UserId = 1, Username = "u1", Email = "u1@t.com", Password = "h" });
        await _db.SaveChangesAsync();
        await SeedOrderAsync(1, 1);
        await SeedOrderDetailAsync(1, 1);
        await SeedOrderDetailAsync(1, 2);

        var result = (await _engine.GetPersonalizedAsync(1, 6)).ToList();

        Assert.DoesNotContain(result, r => r.ProductId == 1);
        Assert.DoesNotContain(result, r => r.ProductId == 2);
    }

    /// <summary>无购买历史 — 回退到热门/活跃商品</summary>
    [Fact]
    public async Task Personalized_NoHistory_FallsBackToPopular()
    {
        for (int i = 1; i <= 3; i++)
            await SeedProductAsync(i, $"Perfume{i}", "女士香水", "香水",
                createdAt: DateTime.Now.AddDays(-i));

        _db.Users.Add(new User { UserId = 99, Username = "new", Email = "new@t.com", Password = "h" });
        await _db.SaveChangesAsync();

        var result = (await _engine.GetPersonalizedAsync(99, 6)).ToList();

        // 应返回活跃商品作为回退
        Assert.True(result.Count > 0);
    }

    // ==================== 2. GetPopularProductsAsync ====================

    /// <summary>热门商品 — 按销量排序</summary>
    [Fact]
    public async Task PopularProducts_SortedBySales()
    {
        for (int i = 1; i <= 3; i++)
            await SeedProductAsync(i, $"P{i}", "香水");

        // 产品1 卖5件, 产品2 卖10件, 产品3 卖1件
        await SeedOrderAsync(1, 1);
        for (int i = 0; i < 5; i++) await SeedOrderDetailAsync(1, 1);

        await SeedOrderAsync(2, 1);
        for (int i = 0; i < 10; i++) await SeedOrderDetailAsync(2, 2);

        await SeedOrderAsync(3, 1);
        await SeedOrderDetailAsync(3, 3);

        var result = (await _engine.GetPopularProductsAsync(3)).ToList();

        // 产品2 应排第一
        Assert.Equal(2, result[0]);
        Assert.Equal(3, result.Count);
    }

    /// <summary>热门商品不足 — 用活跃商品补充</summary>
    [Fact]
    public async Task PopularProducts_InsufficientSales_FillsWithActive()
    {
        // 只有1个商品有销量，另2个活跃但无销量
        await SeedProductAsync(1, "Hot", "香水", createdAt: DateTime.Now.AddDays(-10));
        await SeedProductAsync(2, "New1", "香水", createdAt: DateTime.Now.AddDays(-2));
        await SeedProductAsync(3, "New2", "香水", createdAt: DateTime.Now.AddDays(-1));

        await SeedOrderAsync(1, 1);
        await SeedOrderDetailAsync(1, 1, 5);

        var result = (await _engine.GetPopularProductsAsync(3)).ToList();

        Assert.Equal(3, result.Count);
        Assert.Contains(1, result);
    }

    // ==================== 3. GetRelatedAsync ====================

    /// <summary>相关产品 — 同品类应排前面</summary>
    [Fact]
    public async Task RelatedProducts_SameCategory_RankedHigher()
    {
        await SeedProductAsync(1, "Rose", "女士香水", "香水");
        await SeedProductAsync(2, "Jasmine", "女士香水", "香水");
        await SeedProductAsync(3, "Cologne", "男士香水", "古龙水");
        await SeedProductAsync(4, "Lavender", "女士香水", "精油");

        var result = (await _engine.GetRelatedAsync(1, 4)).ToList();

        // 产品2（同品类+同类型）应该在列表中
        Assert.Contains(result, r => r.ProductId == 2);
        // 不包含自身
        Assert.DoesNotContain(result, r => r.ProductId == 1);
    }

    /// <summary>不存在的产品 — 返回空</summary>
    [Fact]
    public async Task RelatedProducts_NonExistent_ReturnsEmpty()
    {
        var result = await _engine.GetRelatedAsync(9999, 4);
        Assert.Empty(result);
    }

    // ==================== 4. GetTrendingAsync ====================

    /// <summary>趋势分析 — 近期有销量的商品应有趋势评分</summary>
    [Fact]
    public async Task Trending_RecentSales_HasScores()
    {
        await SeedProductAsync(1, "Trending1", "香水");
        await SeedProductAsync(2, "Trending2", "香水");

        // 近期订单（5天前）
        _db.Orders.Add(new Order
        {
            OrderId = 100, OrderNo = "TREND1", UserId = 1,
            TotalAmount = 200, Status = "Paid",
            CreatedAt = DateTime.Now.AddDays(-5), UpdatedAt = DateTime.Now
        });
        _db.OrderDetails.Add(new OrderDetail
        {
            OrderId = 100, ProductId = 1, Quantity = 10,
            UnitPrice = 100m, Subtotal = 1000m, ProductName = "Trending1"
        });
        _db.OrderDetails.Add(new OrderDetail
        {
            OrderId = 100, ProductId = 2, Quantity = 3,
            UnitPrice = 100m, Subtotal = 300m, ProductName = "Trending2"
        });
        await _db.SaveChangesAsync();

        var result = (await _engine.GetTrendingAsync(5)).ToList();

        Assert.True(result.Count >= 1);
        Assert.All(result, r => Assert.True(r.Score > 0));
    }

    // ==================== 5. GetNewArrivalsAsync ====================

    /// <summary>新品推荐 — 按创建时间降序</summary>
    [Fact]
    public async Task NewArrivals_SortedByCreatedAt()
    {
        await SeedProductAsync(1, "Old", "香水", createdAt: DateTime.Now.AddDays(-30));
        await SeedProductAsync(2, "Medium", "香水", createdAt: DateTime.Now.AddDays(-10));
        await SeedProductAsync(3, "Newest", "香水", createdAt: DateTime.Now.AddDays(-1));

        var result = (await _engine.GetNewArrivalsAsync(3)).ToList();

        Assert.Equal(3, result[0]); // 最新排第一
        Assert.Equal(2, result[1]);
        Assert.Equal(1, result[2]);
    }

    // ==================== 6. 接口覆盖检查 ====================

    /// <summary>IRecommendationEngine 覆盖 V18 所有 RE_* 函数</summary>
    [Fact]
    public void IRecommendationEngine_CoversAllV18Functions()
    {
        var methods = typeof(IRecommendationEngine).GetMethods().Select(m => m.Name).ToHashSet();

        Assert.Contains("GetPersonalizedAsync", methods);
        Assert.Contains("GetPersonalizedRecommendationsAsync", methods);
        Assert.Contains("GetPopularProductsAsync", methods);
        Assert.Contains("GetRelatedAsync", methods);
        Assert.Contains("GetRelatedProductsAsync", methods);
        Assert.Contains("GetTrendingAsync", methods);
        Assert.Contains("GetNewArrivalsAsync", methods);
    }
}
