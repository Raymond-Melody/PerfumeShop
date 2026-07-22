using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// PromotionEngine V18 关键路径 SQLite 测试
/// 使用真实 SQLite 内存数据库验证（非 InMemory Provider）
/// 重点验证：ExecuteUpdateAsync 在真实 DB 的行为、事务一致性、JOIN 查询正确性
/// </summary>
public class PromotionEngineV18SqliteTests : SqliteTestBase
{
    private PromotionEngine CreateEngine() => new(Db);

    // ==================== 辅助 ====================

    private async Task SeedCouponAsync(string code, string type, decimal value, decimal minSpend = 0,
        decimal maxDiscount = 0, bool firstOrderOnly = false, string? category = null,
        int totalQty = 100, int usedQty = 0)
    {
        Db.Coupons.Add(new Coupon
        {
            CouponCode = code, CouponName = $"Test-{code}", CouponType = type,
            DiscountValue = value, MinSpend = minSpend, MaxDiscount = maxDiscount,
            FirstOrderOnly = firstOrderOnly, ApplicableCategory = category,
            TotalQty = totalQty, UsedQty = usedQty, IsActive = true,
            ValidFrom = DateTime.Now.AddDays(-30), ValidTo = DateTime.Now.AddDays(30),
            IsPublic = false, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await Db.SaveChangesAsync();
    }

    private async Task SeedUserAsync(int userId)
    {
        Db.Users.Add(new User
        {
            UserId = userId, Username = $"user{userId}", Email = $"u{userId}@t.com",
            Password = "h", CreatedAt = DateTime.Now
        });
        await Db.SaveChangesAsync();
    }

    // ==================== 1. ValidateCoupon 完整校验链（SQLite） ====================

    /// <summary>Fixed 满减券 SQLite — 验证 AsNoTracking + JOIN 在真实 DB 正确</summary>
    [Fact]
    public async Task ValidateCoupon_Fixed_Sqlite_ReturnsCorrectDiscount()
    {
        await SeedCouponAsync("FIX50", "fixed", 50, minSpend: 200);
        await SeedUserAsync(1);

        var result = await CreateEngine().ValidateCouponAsync("FIX50", 1, 300);

        Assert.True(result.Valid);
        Assert.Equal(50m, result.Discount);
        Assert.Equal("fixed", result.Type);
    }

    /// <summary>Percentage 折扣 SQLite — 验证 maxDiscount 封顶逻辑在真实 DB</summary>
    [Fact]
    public async Task ValidateCoupon_Percentage_Sqlite_CappedCorrectly()
    {
        await SeedCouponAsync("PCT20", "percentage", 20, minSpend: 100, maxDiscount: 50);
        await SeedUserAsync(1);

        // 500 * 20% = 100, 但 maxDiscount=50
        var result = await CreateEngine().ValidateCouponAsync("PCT20", 1, 500);

        Assert.True(result.Valid);
        Assert.Equal(50m, result.Discount);
    }

    /// <summary>过期优惠券 SQLite — AsNoTracking 查询在真实 DB 的时间比较</summary>
    [Fact]
    public async Task ValidateCoupon_Expired_Sqlite_ReturnsInvalid()
    {
        Db.Coupons.Add(new Coupon
        {
            CouponCode = "EXPIRED", CouponName = "Expired", CouponType = "fixed",
            DiscountValue = 10, IsActive = true,
            ValidFrom = DateTime.Now.AddDays(-60), ValidTo = DateTime.Now.AddDays(-1),
            TotalQty = 100, UsedQty = 0, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await Db.SaveChangesAsync();
        await SeedUserAsync(1);

        var result = await CreateEngine().ValidateCouponAsync("EXPIRED", 1, 200);
        Assert.False(result.Valid);
        Assert.Contains("有效期", result.Message);
    }

    /// <summary>首单限制 SQLite — Orders COUNT 在真实 DB</summary>
    [Fact]
    public async Task ValidateCoupon_FirstOrderOnly_Sqlite_Rejected()
    {
        await SeedCouponAsync("FIRSTONLY", "fixed", 30, firstOrderOnly: true);
        await SeedUserAsync(1);
        Db.Orders.Add(new Order
        {
            OrderNo = "ORD001", UserId = 1, TotalAmount = 100,
            Status = "Paid", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await Db.SaveChangesAsync();

        var result = await CreateEngine().ValidateCouponAsync("FIRSTONLY", 1, 200);
        Assert.False(result.Valid);
        Assert.Contains("首单", result.Message);
    }

    /// <summary>已使用检查 SQLite — UserCoupons COUNT(status='used') 在真实 DB</summary>
    [Fact]
    public async Task ValidateCoupon_AlreadyUsed_Sqlite_Rejected()
    {
        await SeedCouponAsync("USED", "fixed", 20);
        await SeedUserAsync(1);
        Db.UserCoupons.Add(new UserCoupon
        {
            UserId = 1, CouponId = 1, CouponCode = "USED",
            Source = "test", Status = "used", ObtainedAt = DateTime.Now.AddDays(-5),
            UsedAt = DateTime.Now.AddDays(-1)
        });
        await Db.SaveChangesAsync();

        var result = await CreateEngine().ValidateCouponAsync("USED", 1, 200);
        Assert.False(result.Valid);
        Assert.Contains("已使用", result.Message);
    }

    // ==================== 2. IssueCoupon + UseCoupon 事务一致性（SQLite） ====================

    /// <summary>IssueCoupon SQLite — 验证跟踪实体更新 + SaveChanges 事务在真实 DB</summary>
    [Fact]
    public async Task IssueCoupon_Sqlite_UpdatesUsedQtyTransactionally()
    {
        await SeedCouponAsync("WELCOME", "fixed", 10);
        await SeedUserAsync(1);

        var engine = CreateEngine();
        var issued = await engine.IssueCouponAsync(1, "WELCOME", "new_user");
        Assert.True(issued);

        // 验证 Coupon.UsedQty 增加
        var coupon = await Db.Coupons.FirstAsync(c => c.CouponCode == "WELCOME");
        Assert.Equal(1, coupon.UsedQty);

        // 验证 UserCoupon 已创建
        var uc = await Db.UserCoupons.FirstAsync(x => x.UserId == 1 && x.CouponCode == "WELCOME");
        Assert.Equal("available", uc.Status);
    }

    /// <summary>UseCoupon SQLite — 验证 ExecuteUpdateAsync 在真实 DB（用 AsNoTracking 避免缓存）</summary>
    [Fact]
    public async Task UseCoupon_Sqlite_ExecuteUpdateWorks()
    {
        await SeedCouponAsync("USEME", "fixed", 20);
        await SeedUserAsync(1);

        // 先发放
        var engine = CreateEngine();
        await engine.IssueCouponAsync(1, "USEME", "test");

        // 使用
        var used = await engine.UseCouponAsync("USEME", 1, 99);
        Assert.True(used);

        // 验证 UserCoupon 状态变更 — AsNoTracking 绕过 change tracker 缓存
        var uc = await Db.UserCoupons.AsNoTracking().FirstAsync(x => x.UserId == 1 && x.CouponCode == "USEME");
        Assert.Equal("used", uc.Status);
        Assert.NotNull(uc.UsedAt);

        // 验证 Coupon.UsedQty 再次增加
        var coupon = await Db.Coupons.AsNoTracking().FirstAsync(c => c.CouponCode == "USEME");
        Assert.Equal(2, coupon.UsedQty); // 1 from issue + 1 from use
    }

    // ==================== 3. ApplyDiscount + SiteSetting（SQLite） ====================

    /// <summary>ApplyDiscount SQLite — 验证 SiteSetting 查询 + 满减 + 优惠券叠加在真实 DB</summary>
    [Fact]
    public async Task ApplyDiscount_Sqlite_SiteSettingAndCouponCombined()
    {
        Db.SiteSettings.Add(new SiteSetting
        {
            SettingKey = "Promotion_Threshold", SettingValue = "100|80|满100减80"
        });
        await SeedCouponAsync("BIG50", "fixed", 50);
        await SeedUserAsync(1);
        await Db.SaveChangesAsync();

        var engine = CreateEngine();
        var result = await engine.ApplyDiscountAsync(100, 1, "BIG50");

        // 总折扣不超过 50%（cartTotal=100, max=50）
        Assert.True(result.TotalDiscount <= 50m);
        Assert.True(result.FinalAmount >= 50m);
    }

    // ==================== 4. GetPromotionHistory JOIN 查询（SQLite） ====================

    /// <summary>GetPromotionHistory SQLite — 验证 UserCoupons JOIN Coupons 在真实 DB</summary>
    [Fact]
    public async Task GetPromotionHistory_Sqlite_JoinWorks()
    {
        await SeedCouponAsync("HIST01", "fixed", 10);
        await SeedUserAsync(1);
        Db.UserCoupons.Add(new UserCoupon
        {
            UserId = 1, CouponId = 1, CouponCode = "HIST01",
            Source = "test", Status = "used", ObtainedAt = DateTime.Now.AddDays(-10),
            UsedAt = DateTime.Now.AddDays(-5), UsedOrderId = 99
        });
        await Db.SaveChangesAsync();

        var history = (await CreateEngine().GetPromotionHistoryByUserAsync(1)).ToList();
        Assert.Single(history);
        Assert.Equal("HIST01", history[0].CouponCode);
        Assert.Equal("used", history[0].Status);
    }
}
