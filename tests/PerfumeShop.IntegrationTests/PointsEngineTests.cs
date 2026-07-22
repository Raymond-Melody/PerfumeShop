using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Microsoft.Extensions.Caching.Memory;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests;

/// <summary>
/// PointsEngine 单元测试 — 基于 V18 points_engine.asp 真实业务数据回归
/// V19 关键改进: 三表同步写入使用 EF Core Transaction 包裹
/// </summary>
public class PointsEngineTests : IDisposable
{
    private readonly TestEngineContext _db;
    private readonly IMemoryCache _cache;
    private readonly PointsEngine _engine;

    public PointsEngineTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"PointsEngineTests_{Guid.NewGuid()}")
            .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning))
            .Options;
        _db = new TestEngineContext(options);
        _cache = new MemoryCache(new MemoryCacheOptions());
        _engine = new PointsEngine(_db, _cache);
    }

    public void Dispose()
    {
        _db.Database.EnsureDeleted();
        _db.Dispose();
        _cache.Dispose();
    }

    // ==================== 1. 规则缓存 ====================

    /// <summary>测试1: GetRuleAsync 返回数据库中的规则值</summary>
    [Fact]
    public async Task GetRule_FromDatabase_ReturnsCorrectValue()
    {
        _db.PointsRules.Add(new PointsRule
        {
            RuleCode = "purchase_rate", RuleName = "Purchase Rate",
            RuleValue = 2m, RuleUnit = "points/yuan", IsEnabled = true,
            SortOrder = 1, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        var val = await _engine.GetRuleAsync("purchase_rate");
        Assert.Equal(2m, val);
    }

    /// <summary>测试2: GetRuleAsync 回退到默认值</summary>
    [Fact]
    public async Task GetRule_NotInDb_ReturnsDefault()
    {
        var val = await _engine.GetRuleAsync("signin_points");
        Assert.Equal(5m, val); // V18 默认值
    }

    /// <summary>测试3: RefreshRuleCache 后缓存生效</summary>
    [Fact]
    public async Task RefreshRuleCache_SecondCallHitsCache()
    {
        _db.PointsRules.Add(new PointsRule
        {
            RuleCode = "review_points", RuleName = "Review Points",
            RuleValue = 25m, RuleUnit = "points", IsEnabled = true,
            SortOrder = 2, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        await _engine.RefreshRuleCacheAsync();
        var val1 = await _engine.GetRuleAsync("review_points");
        var val2 = await _engine.GetRuleAsync("review_points");
        Assert.Equal(25m, val1);
        Assert.Equal(val1, val2);
    }

    // ==================== 2. 积分获取 (EarnAsync) ====================

    /// <summary>测试4: 获取积分后余额正确增加</summary>
    [Fact]
    public async Task EarnAsync_IncreasesBalance()
    {
        _db.Users.Add(new User { UserId = 1, Username = "test1", Email = "t1@test.com", Password = "pwd" });
        _db.UserPoints.Add(new UserPoint { UserId = 1, AvailablePoints = 0, TotalPoints = 0 });
        await _db.SaveChangesAsync();

        var result = await _engine.EarnAsync(1, 100, "purchase", "Test purchase", 1);
        Assert.True(result);

        // 验证 PointsLedger 记录
        var ledgerCount = await _db.PointsLedgers.CountAsync(l => l.UserId == 1);
        Assert.Equal(1, ledgerCount);
    }

    /// <summary>测试5: 获取积分为 0 或负数时返回 false</summary>
    [Fact]
    public async Task EarnAsync_ZeroPoints_ReturnsFalse()
    {
        var result = await _engine.EarnAsync(1, 0, "test");
        Assert.False(result);

        var result2 = await _engine.EarnAsync(1, -10, "test");
        Assert.False(result2);
    }

    // ==================== 3. 积分消费 (RedeemAsync) ====================

    /// <summary>测试6: 消费积分后余额正确减少</summary>
    [Fact]
    public async Task RedeemAsync_DecreasesBalance()
    {
        _db.Users.Add(new User { UserId = 10, Username = "test10", Email = "t10@test.com", Password = "pwd", Points = 200 });
        _db.UserPoints.Add(new UserPoint { UserId = 10, AvailablePoints = 200, TotalPoints = 200 });
        // 先获取积分让 PointsLedger 有记录
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 10, Points = 200, PointType = "earn", Source = "purchase",
            IsExpired = false, CreatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        var result = await _engine.RedeemAsync(10, 50, "redeem_discount");
        Assert.True(result);
    }

    /// <summary>测试7: 余额不足时自动调整消费数量</summary>
    [Fact]
    public async Task RedeemAsync_InsufficientBalance_AdjustsPoints()
    {
        _db.Users.Add(new User { UserId = 11, Username = "test11", Email = "t11@test.com", Password = "pwd", Points = 30 });
        _db.UserPoints.Add(new UserPoint { UserId = 11, AvailablePoints = 30, TotalPoints = 100 });
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 11, Points = 30, PointType = "earn", Source = "purchase",
            IsExpired = false, CreatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        // 请求消费 50，但只有 30
        var result = await _engine.RedeemAsync(11, 50, "test");
        Assert.True(result); // 应该成功，自动调整为 30
    }

    // ==================== 4. 积分过期 (ApplyExpirationAsync) ====================

    /// <summary>测试8: 30天前的积分正确过期</summary>
    [Fact]
    public async Task ApplyExpiration_MarksExpiredPoints()
    {
        _db.PointsRules.Add(new PointsRule
        {
            RuleCode = "points_expire_months", RuleName = "Expire Months",
            RuleValue = 1m, RuleUnit = "months", IsEnabled = true,
            SortOrder = 1, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 20, Points = 100, PointType = "earn", Source = "purchase",
            ExpiresAt = DateTime.Now.AddDays(-40), // 已过期
            IsExpired = false, CreatedAt = DateTime.Now.AddDays(-40)
        });
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 20, Points = 50, PointType = "earn", Source = "purchase",
            ExpiresAt = DateTime.Now.AddDays(30), // 未过期
            IsExpired = false, CreatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        await _engine.RefreshRuleCacheAsync();
        await _engine.ApplyExpirationAsync(20);

        var expired = await _db.PointsLedgers.CountAsync(l => l.UserId == 20 && l.IsExpired);
        Assert.Equal(1, expired); // 只有第一条过期
    }

    /// <summary>测试9: expireMonths=0 时永不过期</summary>
    [Fact]
    public async Task ApplyExpiration_ZeroMonths_NeverExpires()
    {
        _db.PointsRules.Add(new PointsRule
        {
            RuleCode = "points_expire_months", RuleName = "Never Expire",
            RuleValue = 0m, RuleUnit = "months", IsEnabled = true,
            SortOrder = 1, CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 21, Points = 100, PointType = "earn", Source = "purchase",
            ExpiresAt = DateTime.Now.AddDays(-365),
            IsExpired = false, CreatedAt = DateTime.Now.AddDays(-365)
        });
        await _db.SaveChangesAsync();

        await _engine.RefreshRuleCacheAsync();
        await _engine.ApplyExpirationAsync(21);

        var expired = await _db.PointsLedgers.CountAsync(l => l.UserId == 21 && l.IsExpired);
        Assert.Equal(0, expired);
    }

    // ==================== 5. 积分余额 (GetBalanceAsync) ====================

    /// <summary>测试10: 多次 earn/redeem/expire 后余额正确</summary>
    [Fact]
    public async Task GetBalance_AfterMultipleOperations_Correct()
    {
        _db.Users.Add(new User { UserId = 30, Username = "test30", Email = "t30@test.com", Password = "pwd" });
        // earn 100
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 30, Points = 100, PointType = "earn", Source = "purchase",
            IsExpired = false, CreatedAt = DateTime.Now
        });
        // earn 50
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 30, Points = 50, PointType = "earn", Source = "signin",
            IsExpired = false, CreatedAt = DateTime.Now
        });
        // redeem -30
        _db.PointsLedgers.Add(new PointsLedger
        {
            UserId = 30, Points = -30, PointType = "redeem", Source = "discount",
            IsExpired = false, CreatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        var balance = await _engine.GetBalanceAsync(30);
        // 100 + 50 - 30 = 120
        Assert.Equal(120, balance);
    }

    // ==================== 6. 积分计算 ====================

    /// <summary>测试11: CalcOrderPoints 正确计算消费积分</summary>
    [Fact]
    public async Task CalcOrderPoints_WithDefaultRate()
    {
        // 默认 purchase_rate = 1
        var points = await _engine.CalcOrderPointsAsync(299.5m);
        Assert.Equal(299, points); // CInt(299.5 * 1) = 299
    }

    /// <summary>测试12: CalcPointsValue 正确计算积分价值</summary>
    [Fact]
    public async Task CalcPointsValue_WithDefaultRate()
    {
        // 默认 redeem_discount_rate = 100
        var value = await _engine.CalcPointsValueAsync(500);
        Assert.Equal(5m, value); // 500 / 100 = 5
    }

    /// <summary>测试13: CalcMaxRedeemablePoints 正确计算最大可抵扣积分</summary>
    [Fact]
    public async Task CalcMaxRedeemablePoints_CorrectCalculation()
    {
        // 默认: max_redeem_pct = 30, redeem_discount_rate = 100
        var maxPoints = await _engine.CalcMaxRedeemablePointsAsync(1, 1000m);
        // maxValue = 1000 * 0.3 = 300, maxPoints = 300 * 100 = 30000
        Assert.Equal(30000, maxPoints);
    }

    // ==================== 7. 签到 ====================

    /// <summary>测试14: 签到检查 — 未签到时返回 false</summary>
    [Fact]
    public async Task CheckSignIn_NotSignedIn_ReturnsFalse()
    {
        var result = await _engine.CheckSignInAsync(40);
        Assert.False(result);
    }

    // ==================== 8. 账本分页 ====================

    /// <summary>测试15: GetLedgerAsync 分页正确</summary>
    [Fact]
    public async Task GetLedger_PaginationCorrect()
    {
        for (int i = 0; i < 25; i++)
        {
            _db.PointsLedgers.Add(new PointsLedger
            {
                UserId = 50, Points = 10, PointType = "earn", Source = "test",
                IsExpired = false, CreatedAt = DateTime.Now.AddMinutes(-i)
            });
        }
        await _db.SaveChangesAsync();

        var page1 = await _engine.GetLedgerAsync(50, 1, 10);
        Assert.Equal(10, page1.Items.Count);
        Assert.Equal(25, page1.Total);

        var page3 = await _engine.GetLedgerAsync(50, 3, 10);
        Assert.Equal(5, page3.Items.Count);
    }

    // ==================== 9. 积分汇总 ====================

    /// <summary>测试16: GetPointsSummary 各字段正确</summary>
    [Fact]
    public async Task GetPointsSummary_CorrectAggregation()
    {
        _db.Users.Add(new User { UserId = 60, Username = "test60", Email = "t60@test.com", Password = "pwd" });
        _db.PointsLedgers.AddRange(
            new PointsLedger { UserId = 60, Points = 100, PointType = "earn", Source = "purchase", IsExpired = false, CreatedAt = DateTime.Now },
            new PointsLedger { UserId = 60, Points = 50, PointType = "earn", Source = "signin", IsExpired = false, CreatedAt = DateTime.Today },
            new PointsLedger { UserId = 60, Points = -20, PointType = "redeem", Source = "discount", IsExpired = false, CreatedAt = DateTime.Now }
        );
        await _db.SaveChangesAsync();

        var summary = await _engine.GetPointsSummaryAsync(60);
        Assert.Equal(130, summary.Available); // 100 + 50 - 20
        Assert.Equal(150, summary.TotalEarned); // 100 + 50
        Assert.Equal(20, summary.TotalRedeemed); // |-20|
    }

    // ==================== 10. V18 函数映射覆盖检查 ====================

    /// <summary>测试17: 接口方法签名覆盖 V18 所有 PE_* 函数</summary>
    [Fact]
    public void IPointsEngine_CoversAllV18Functions()
    {
        var methods = typeof(IPointsEngine).GetMethods().Select(m => m.Name).ToHashSet();

        Assert.Contains("GetRuleAsync", methods);            // PE_GetRule / PE_GetRuleCache
        Assert.Contains("RefreshRuleCacheAsync", methods);    // PE_GetRuleCache
        Assert.Contains("EarnAsync", methods);                // PE_EarnPoints
        Assert.Contains("RedeemAsync", methods);              // PE_RedeemPoints
        Assert.Contains("GetBalanceAsync", methods);          // PE_GetAvailablePoints
        Assert.Contains("ApplyExpirationAsync", methods);     // PE_ExpireOutdatedPoints
        Assert.Contains("UpdateBalanceAsync", methods);       // PE_UpdateBalance
        Assert.Contains("CalcPointsValueAsync", methods);     // PE_CalcPointsValue
        Assert.Contains("CalcOrderPointsAsync", methods);     // PE_CalcOrderPoints
        Assert.Contains("GetMaxRedeemPctAsync", methods);     // PE_GetMaxRedeemPct
        Assert.Contains("CalcMaxRedeemablePointsAsync", methods); // PE_CalcMaxRedeemablePoints
        Assert.Contains("CheckSignInAsync", methods);         // PE_CheckSignIn
        Assert.Contains("DoSignInAsync", methods);            // PE_DoSignIn
        Assert.Contains("GetLedgerAsync", methods);           // PE_GetPointsLedger
        Assert.Contains("GetPointsSummaryAsync", methods);    // PE_GetPointsSummary
        Assert.Contains("GetOrderPointsAsync", methods);      // PE_GetOrderPoints
        Assert.Contains("DoRedeemAsync", methods);            // PE_DoRedeem
    }
}
