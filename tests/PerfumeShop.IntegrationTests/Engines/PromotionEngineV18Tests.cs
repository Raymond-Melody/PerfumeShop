using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// PromotionEngine V18 完整算法单元测试
/// 使用 EngineTestContext（带 SiteSetting 主键）
/// </summary>
public class PromotionEngineV18Tests : IDisposable
{
    private readonly EngineTestContext _db;
    private readonly PromotionEngine _engine;

    public PromotionEngineV18Tests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"PromoV18_{Guid.NewGuid()}")
            .Options;
        _db = new EngineTestContext(options);
        _engine = new PromotionEngine(_db);
    }

    public void Dispose()
    {
        _db.Database.EnsureDeleted();
        _db.Dispose();
    }

    // ==================== 辅助 ====================

    private async Task SeedCouponAsync(string code, string type, decimal value, decimal minSpend = 0,
        decimal maxDiscount = 0, bool firstOrderOnly = false, string? category = null,
        int totalQty = 100, int usedQty = 0)
    {
        _db.Coupons.Add(new Coupon
        {
            CouponCode = code, CouponName = $"Test-{code}", CouponType = type,
            DiscountValue = value, MinSpend = minSpend, MaxDiscount = maxDiscount,
            FirstOrderOnly = firstOrderOnly, ApplicableCategory = category,
            TotalQty = totalQty, UsedQty = usedQty, IsActive = true,
            ValidFrom = DateTime.Now.AddDays(-30), ValidTo = DateTime.Now.AddDays(30),
            IsPublic = false, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();
    }

    private async Task SeedUserAsync(int userId, string? tier = null)
    {
        _db.Users.Add(new User
        {
            UserId = userId, Username = $"user{userId}", Email = $"u{userId}@t.com",
            Password = "h", CustomerTier = tier, CreatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();
    }

    // ==================== 1. 四种促销类型 ====================

    /// <summary>Fixed 满减券 — 减固定金额</summary>
    [Fact]
    public async Task ValidateCoupon_Fixed_ReturnsFixedDiscount()
    {
        await SeedCouponAsync("FIX50", "fixed", 50, minSpend: 200);
        await SeedUserAsync(1);

        var result = await _engine.ValidateCouponAsync("FIX50", 1, 300);

        Assert.True(result.Valid);
        Assert.Equal(50m, result.Discount);
        Assert.Equal("fixed", result.Type);
    }

    /// <summary>Percentage 折扣券 — 按百分比折扣，受 maxDiscount 限制</summary>
    [Fact]
    public async Task ValidateCoupon_Percentage_CappedAtMaxDiscount()
    {
        await SeedCouponAsync("PCT20", "percentage", 20, minSpend: 100, maxDiscount: 50);
        await SeedUserAsync(1);

        // cartTotal=500, 20%=100, 但 maxDiscount=50
        var result = await _engine.ValidateCouponAsync("PCT20", 1, 500);

        Assert.True(result.Valid);
        Assert.Equal(50m, result.Discount);
    }

    /// <summary>FreeShipping 免邮券 — 折扣金额为0</summary>
    [Fact]
    public async Task ValidateCoupon_FreeShipping_ReturnsZeroDiscount()
    {
        await SeedCouponAsync("FREESHIP", "free_shipping", 0);
        await SeedUserAsync(1);

        var result = await _engine.ValidateCouponAsync("FREESHIP", 1, 150);

        Assert.True(result.Valid);
        Assert.Equal(0m, result.Discount);
        Assert.Equal("free_shipping", result.Type);
    }

    /// <summary>Gift 礼品券 — 折扣金额为0</summary>
    [Fact]
    public async Task ValidateCoupon_Gift_ReturnsZeroDiscount()
    {
        await SeedCouponAsync("GIFT01", "gift", 0);
        await SeedUserAsync(1);

        var result = await _engine.ValidateCouponAsync("GIFT01", 1, 200);

        Assert.True(result.Valid);
        Assert.Equal(0m, result.Discount);
        Assert.Equal("gift", result.Type);
    }

    // ==================== 2. 校验链 ====================

    /// <summary>过期优惠券 — 应返回无效</summary>
    [Fact]
    public async Task ValidateCoupon_Expired_ReturnsInvalid()
    {
        _db.Coupons.Add(new Coupon
        {
            CouponCode = "EXPIRED", CouponName = "Expired", CouponType = "fixed",
            DiscountValue = 10, IsActive = true,
            ValidFrom = DateTime.Now.AddDays(-60), ValidTo = DateTime.Now.AddDays(-1),
            TotalQty = 100, UsedQty = 0, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();
        await SeedUserAsync(1);

        var result = await _engine.ValidateCouponAsync("EXPIRED", 1, 200);

        Assert.False(result.Valid);
        Assert.Contains("有效期", result.Message);
    }

    /// <summary>首单限制 — 已有订单的用户不能使用</summary>
    [Fact]
    public async Task ValidateCoupon_FirstOrderOnly_ExistingUserRejected()
    {
        await SeedCouponAsync("FIRSTONLY", "fixed", 30, firstOrderOnly: true);
        await SeedUserAsync(1);
        _db.Orders.Add(new Order
        {
            OrderNo = "ORD001", UserId = 1, TotalAmount = 100,
            Status = "Paid", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        var result = await _engine.ValidateCouponAsync("FIRSTONLY", 1, 200);

        Assert.False(result.Valid);
        Assert.Contains("首单", result.Message);
    }

    /// <summary>库存耗尽 — 已领完的优惠券不可使用</summary>
    [Fact]
    public async Task ValidateCoupon_OutOfStock_ReturnsInvalid()
    {
        await SeedCouponAsync("NOSTOCK", "fixed", 20, totalQty: 10, usedQty: 10);
        await SeedUserAsync(1);

        var result = await _engine.ValidateCouponAsync("NOSTOCK", 1, 200);

        Assert.False(result.Valid);
        Assert.Contains("领完", result.Message);
    }

    /// <summary>品类限制 — 不匹配的品类应被拒绝</summary>
    [Fact]
    public async Task ValidateCoupon_CategoryMismatch_ReturnsInvalid()
    {
        await SeedCouponAsync("CATCOUPON", "percentage", 15, category: "女士香水");
        await SeedUserAsync(1);

        var result = await _engine.ValidateCouponAsync("CATCOUPON", 1, 300, cartCategory: "男士香水");

        Assert.False(result.Valid);
        Assert.Contains("品类", result.Message);
    }

    /// <summary>最低消费未达 — 应返回无效</summary>
    [Fact]
    public async Task ValidateCoupon_BelowMinSpend_ReturnsInvalid()
    {
        await SeedCouponAsync("MIN200", "fixed", 30, minSpend: 200);
        await SeedUserAsync(1);

        var result = await _engine.ValidateCouponAsync("MIN200", 1, 100);

        Assert.False(result.Valid);
        Assert.Contains("最低消费", result.Message);
    }

    /// <summary>已使用过的优惠券 — 不可重复使用</summary>
    [Fact]
    public async Task ValidateCoupon_AlreadyUsed_ReturnsInvalid()
    {
        await SeedCouponAsync("USED", "fixed", 20);
        await SeedUserAsync(1);
        _db.UserCoupons.Add(new UserCoupon
        {
            UserId = 1, CouponId = 1, CouponCode = "USED",
            Source = "test", Status = "used", ObtainedAt = DateTime.Now.AddDays(-5),
            UsedAt = DateTime.Now.AddDays(-1)
        });
        await _db.SaveChangesAsync();

        var result = await _engine.ValidateCouponAsync("USED", 1, 200);

        Assert.False(result.Valid);
        Assert.Contains("已使用", result.Message);
    }

    // ==================== 3. 折扣计算与满减（需要 SiteSetting） ====================

    /// <summary>满减促销 — 达到门槛应享受折扣</summary>
    [Fact]
    public async Task CalculateDiscount_ThresholdReached_AppliesDiscount()
    {
        _db.SiteSettings.Add(new SiteSetting
        {
            SettingKey = "Promotion_Threshold", SettingValue = "299|50|满299减50"
        });
        await SeedUserAsync(1);
        await _db.SaveChangesAsync();

        var discount = await _engine.CalculateDiscountAsync(350, 1);
        Assert.Equal(50m, discount);
    }

    /// <summary>免运费检查</summary>
    [Fact]
    public void CheckFreeShipping_AboveThreshold_ReturnsTrue()
    {
        Assert.True(_engine.CheckFreeShipping(299));
        Assert.True(_engine.CheckFreeShipping(500));
        Assert.False(_engine.CheckFreeShipping(100));
    }

    // ==================== 4. ApplyDiscount 综合 ====================

    /// <summary>ApplyDiscount — 促销+优惠券叠加，不超过50%上限</summary>
    [Fact]
    public async Task ApplyDiscount_CombinedRespectsMaxRatio()
    {
        _db.SiteSettings.Add(new SiteSetting
        {
            SettingKey = "Promotion_Threshold", SettingValue = "100|80|满100减80"
        });
        await SeedCouponAsync("BIG50", "fixed", 50);
        await SeedUserAsync(1);
        await _db.SaveChangesAsync();

        // cartTotal=100, promo=80, coupon=50, total=130 > 100*0.5=50
        var result = await _engine.ApplyDiscountAsync(100, 1, "BIG50");

        Assert.True(result.TotalDiscount <= 50m);
        Assert.True(result.FinalAmount >= 50m);
    }

    // ==================== 5. 发放与查询 ====================

    /// <summary>IssueCoupon + GetUserCoupons — 发放后可查询</summary>
    [Fact]
    public async Task IssueCoupon_ThenGetUserCoupons_ReturnsIssued()
    {
        await SeedCouponAsync("WELCOME10", "fixed", 10);
        await SeedUserAsync(1);

        var issued = await _engine.IssueCouponAsync(1, "WELCOME10", "new_user");
        Assert.True(issued);

        var coupons = await _engine.GetUserCouponsAsync(1, "available");
        Assert.Single(coupons);
    }

    /// <summary>GetPromotionHistoryByUser — 返回用户历史</summary>
    [Fact]
    public async Task GetPromotionHistory_ReturnsUserRecords()
    {
        await SeedCouponAsync("HIST01", "fixed", 10);
        await SeedUserAsync(1);
        _db.UserCoupons.Add(new UserCoupon
        {
            UserId = 1, CouponId = 1, CouponCode = "HIST01",
            Source = "test", Status = "used", ObtainedAt = DateTime.Now.AddDays(-10),
            UsedAt = DateTime.Now.AddDays(-5), UsedOrderId = 99
        });
        await _db.SaveChangesAsync();

        var history = await _engine.GetPromotionHistoryByUserAsync(1);
        Assert.Single(history);
        Assert.Equal("used", history.First().Status);
    }

    // ==================== 6. 接口覆盖检查 ====================

    /// <summary>IPromotionEngine 覆盖 V18 所有 PE_* 函数</summary>
    [Fact]
    public void IPromotionEngine_CoversAllV18Functions()
    {
        var methods = typeof(IPromotionEngine).GetMethods().Select(m => m.Name).ToHashSet();

        Assert.Contains("CheckPromotionEligibilityAsync", methods);
        Assert.Contains("ValidateCouponAsync", methods);
        Assert.Contains("CalculateDiscountAsync", methods);
        Assert.Contains("ApplyDiscountAsync", methods);
        Assert.Contains("CheckFreeShipping", methods);
        Assert.Contains("UseCouponAsync", methods);
        Assert.Contains("IssueCouponAsync", methods);
        Assert.Contains("IssueWelcomeCouponsAsync", methods);
        Assert.Contains("IssueTierUpgradeCouponAsync", methods);
        Assert.Contains("GetUserCouponCountAsync", methods);
        Assert.Contains("GetPromotionHistoryByUserAsync", methods);
        Assert.Contains("GetUserCouponsAsync", methods);
        Assert.Contains("GetApplicableCouponsAsync", methods);
        Assert.Contains("GetCouponStatsAsync", methods);
    }
}
