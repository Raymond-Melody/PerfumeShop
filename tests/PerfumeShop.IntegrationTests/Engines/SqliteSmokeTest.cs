using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// SQLite 冒烟测试 — 验证 SQLite EnsureCreated 能正常创建 schema
/// </summary>
public class SqliteSmokeTest : SqliteTestBase
{
    [Fact]
    public void EnsureCreated_TablesExist()
    {
        _ = Db.Orders.Count();
        _ = Db.Products.Count();
        _ = Db.Coupons.Count();
        _ = Db.PaymentRecords.Count();
        _ = Db.ProductionOrders.Count();
    }

    [Fact]
    public async Task CanInsertAndQuery()
    {
        Db.Users.Add(new User { UserId = 1, Username = "test", Email = "t@t.com", Password = "h" });
        Db.Products.Add(new Product { ProductId = 1, ProductName = "Test", Category = "香水", IsActive = true });
        await Db.SaveChangesAsync();

        Assert.Equal(1, Db.Users.Count());
        Assert.Equal(1, Db.Products.Count());
    }
}
