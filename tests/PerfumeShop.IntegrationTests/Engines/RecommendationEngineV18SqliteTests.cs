using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// RecommendationEngine V18 关键路径 SQLite 测试
/// 使用真实 SQLite 内存数据库验证（非 InMemory Provider）
/// 重点验证：GROUP BY + ORDER BY 在真实 SQL、JOIN 查询、趋势评分
/// </summary>
public class RecommendationEngineV18SqliteTests : SqliteTestBase
{
    private RecommendationEngine CreateEngine() => new(Db);

    // ==================== 辅助 ====================

    private async Task SeedProductAsync(int id, string name, string category, string? productType = null,
        bool active = true, DateTime? createdAt = null)
    {
        Db.Products.Add(new Product
        {
            ProductId = id, ProductName = name, Category = category,
            ProductType = productType ?? category,
            IsActive = active, CreatedAt = createdAt ?? DateTime.Now.AddDays(-30)
        });
        await Db.SaveChangesAsync();
    }

    private async Task SeedOrderAsync(int orderId, int userId, string status = "Paid", DateTime? createdAt = null)
    {
        Db.Orders.Add(new Order
        {
            OrderId = orderId, OrderNo = $"ORD{orderId}", UserId = userId,
            TotalAmount = 100, Status = status,
            CreatedAt = createdAt ?? DateTime.Now.AddDays(-5), UpdatedAt = DateTime.Now
        });
        await Db.SaveChangesAsync();
    }

    private async Task SeedOrderDetailAsync(int orderId, int productId, int qty = 1)
    {
        Db.OrderDetails.Add(new OrderDetail
        {
            OrderId = orderId, ProductId = productId, Quantity = qty,
            UnitPrice = 100m, Subtotal = 100m * qty, ProductName = $"P{productId}"
        });
        await Db.SaveChangesAsync();
    }

    // ==================== 1. GetPersonalizedAsync — 协同过滤 SQLite ====================

    /// <summary>协同过滤 SQLite — 验证 Jaccard 相似度 + JOIN 在真实 SQL</summary>
    [Fact]
    public async Task Personalized_Sqlite_CollaborativeFilter()
    {
        // 商品: 1-5（同类）
        for (int i = 1; i <= 5; i++)
            await SeedProductAsync(i, $"Perfume{i}", "女士香水", "香水");

        // 用户101 购买 1, 2
        Db.Users.Add(new User { UserId = 101, Username = "u101", Email = "u101@t.com", Password = "h" });
        await Db.SaveChangesAsync();
        await SeedOrderAsync(1, 101);
        await SeedOrderDetailAsync(1, 1);
        await SeedOrderDetailAsync(1, 2);

        // 用户102 购买 1, 2, 3（与101相似）
        Db.Users.Add(new User { UserId = 102, Username = "u102", Email = "u102@t.com", Password = "h" });
        await Db.SaveChangesAsync();
        await SeedOrderAsync(2, 102);
        await SeedOrderDetailAsync(2, 1);
        await SeedOrderDetailAsync(2, 2);
        await SeedOrderDetailAsync(2, 3);

        var result = (await CreateEngine().GetPersonalizedAsync(101, 6)).ToList();

        // 产品3 应被推荐
        Assert.Contains(result, r => r.ProductId == 3);
        // 产品1,2 不在推荐中（已购买）
        Assert.DoesNotContain(result, r => r.ProductId == 1);
        Assert.DoesNotContain(result, r => r.ProductId == 2);
    }

    // ==================== 2. GetPopularProductsAsync — GROUP BY SQLite ====================

    /// <summary>热门商品 SQLite — 验证 GROUP BY + ORDER BY SUM(Quantity) 在真实 SQL</summary>
    [Fact]
    public async Task PopularProducts_Sqlite_GroupBySorted()
    {
        for (int i = 1; i <= 3; i++)
            await SeedProductAsync(i, $"P{i}", "香水");

        // 产品1=5件, 产品2=10件, 产品3=1件
        await SeedOrderAsync(1, 1);
        for (int i = 0; i < 5; i++) await SeedOrderDetailAsync(1, 1);

        await SeedOrderAsync(2, 1);
        for (int i = 0; i < 10; i++) await SeedOrderDetailAsync(2, 2);

        await SeedOrderAsync(3, 1);
        await SeedOrderDetailAsync(3, 3);

        var result = (await CreateEngine().GetPopularProductsAsync(3)).ToList();

        Assert.Equal(2, result[0]); // 产品2 销量最高
    }

    // ==================== 3. GetRelatedAsync — 余弦相似度 SQLite ====================

    /// <summary>相关产品 SQLite — 验证同品类查询 + 余弦相似度在真实 DB</summary>
    [Fact]
    public async Task RelatedProducts_Sqlite_SameCategoryRanked()
    {
        await SeedProductAsync(1, "Rose", "女士香水", "香水");
        await SeedProductAsync(2, "Jasmine", "女士香水", "香水");
        await SeedProductAsync(3, "Cologne", "男士香水", "古龙水");
        await SeedProductAsync(4, "Lavender", "女士香水", "精油");

        var result = (await CreateEngine().GetRelatedAsync(1, 4)).ToList();

        Assert.Contains(result, r => r.ProductId == 2);
        Assert.DoesNotContain(result, r => r.ProductId == 1); // 不包含自身
    }

    // ==================== 4. GetTrendingAsync — 趋势评分 SQLite ====================

    /// <summary>趋势分析 SQLite — 验证近期/历史比率计算在真实 SQL</summary>
    [Fact]
    public async Task Trending_Sqlite_RecentSalesHaveScores()
    {
        await SeedProductAsync(1, "Trending1", "香水");
        await SeedProductAsync(2, "Trending2", "香水");

        // 近期订单
        Db.Orders.Add(new Order
        {
            OrderId = 100, OrderNo = "TREND1", UserId = 1,
            TotalAmount = 200, Status = "Paid",
            CreatedAt = DateTime.Now.AddDays(-5), UpdatedAt = DateTime.Now
        });
        Db.OrderDetails.Add(new OrderDetail
        {
            OrderId = 100, ProductId = 1, Quantity = 10,
            UnitPrice = 100m, Subtotal = 1000m, ProductName = "Trending1"
        });
        Db.OrderDetails.Add(new OrderDetail
        {
            OrderId = 100, ProductId = 2, Quantity = 3,
            UnitPrice = 100m, Subtotal = 300m, ProductName = "Trending2"
        });
        await Db.SaveChangesAsync();

        var result = (await CreateEngine().GetTrendingAsync(5)).ToList();

        Assert.True(result.Count >= 1);
        Assert.All(result, r => Assert.True(r.Score > 0));
    }

    // ==================== 5. GetNewArrivalsAsync ====================

    /// <summary>新品推荐 SQLite — ORDER BY CreatedAt DESC 在真实 DB</summary>
    [Fact]
    public async Task NewArrivals_Sqlite_SortedByCreatedAt()
    {
        await SeedProductAsync(1, "Old", "香水", createdAt: DateTime.Now.AddDays(-30));
        await SeedProductAsync(2, "Medium", "香水", createdAt: DateTime.Now.AddDays(-10));
        await SeedProductAsync(3, "Newest", "香水", createdAt: DateTime.Now.AddDays(-1));

        var result = (await CreateEngine().GetNewArrivalsAsync(3)).ToList();

        Assert.Equal(3, result[0]); // 最新排第一
        Assert.Equal(1, result[2]); // 最旧排最后
    }
}
