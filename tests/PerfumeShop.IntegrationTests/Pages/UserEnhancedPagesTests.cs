using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using PerfumeShop.Api.Controllers;
using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests.Pages;

/// <summary>
/// M3-C 增强/营销/合规类页面 E2E 测试
/// 覆盖: Reviews、Subscription、KolProducts、DataExport、AccountDelete、MyReferrals、Privacy
/// </summary>
public class UserEnhancedPagesTests : IDisposable
{
    private readonly TestEngineContext _db;
    private readonly UserExtendedController _ctrl;
    private const int TestUserId = 1;

    public UserEnhancedPagesTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"UserEnhancedPagesTests_{Guid.NewGuid()}")
            .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning))
            .Options;
        _db = new TestEngineContext(options);
        _ctrl = new UserExtendedController(_db);
        SeedData();
    }

    private void SeedData()
    {
        _db.Users.Add(new User
        {
            UserId = TestUserId,
            Username = "testuser",
            Email = "test@example.com",
            Password = PerfumeShop.Shared.Security.PasswordHasher.Hash("hashedpwd123"),
            FullName = "测试用户",
            Phone = "13800000001",
            Points = 300,
            IsActive = true,
            CreatedAt = DateTime.Now
        });

        _db.Users.Add(new User
        {
            UserId = 2,
            Username = "referred_user",
            Email = "referred@example.com",
            Password = "pwd",
            FullName = "被邀请用户",
            IsActive = true,
            ReferrerUserId = TestUserId,
            CreatedAt = DateTime.Now.AddDays(-10)
        });

        _db.Products.Add(new Product
        {
            ProductId = 1,
            ProductName = "玫瑰香氛",
            BasePrice = 299m,
            Category = "花香调",
            ImageUrl = "/images/products/rose.jpg",
            Description = "浪漫玫瑰",
            IsActive = true,
            CreatedAt = DateTime.Now
        });
        _db.Products.Add(new Product
        {
            ProductId = 2,
            ProductName = "KOL限定香氛",
            BasePrice = 599m,
            Category = "KOL推荐",
            ProductType = "KOL",
            Kolid = 99,
            ImageUrl = "/images/products/kol.jpg",
            Description = "KOL独家推荐",
            IsActive = true,
            CreatedAt = DateTime.Now
        });

        _db.ProductReviews.Add(new ProductReview
        {
            ReviewId = 1,
            UserId = TestUserId,
            ProductId = 1,
            Rating = 5,
            Title = "非常好闻",
            Comment = "玫瑰香气很浓郁，持久度也很好",
            IsVerifiedPurchase = true,
            LikeCount = 10,
            AIFeelingSummary = "用户喜爱玫瑰花香",
            Status = "Active",
            CreatedAt = DateTime.Now.AddDays(-5)
        });
        _db.ProductReviews.Add(new ProductReview
        {
            ReviewId = 2,
            UserId = TestUserId,
            ProductId = 1,
            Rating = 4,
            Title = "第二次购买",
            Comment = "品质稳定",
            IsVerifiedPurchase = true,
            LikeCount = 3,
            Status = "Active",
            CreatedAt = DateTime.Now.AddDays(-2)
        });

        _db.SubscriptionPlans.Add(new SubscriptionPlan
        {
            PlanId = 1,
            PlanName = "月度精选",
            Period = "monthly",
            Price = 199m,
            SampleCount = 3,
            FullSizeCount = 1,
            FreeShipping = true,
            CancellationFee = 0m,
            IsActive = true,
            SortOrder = 1,
            CreatedAt = DateTime.Now
        });

        _db.UserSubscriptions.Add(new UserSubscription
        {
            SubscriptionId = 1,
            UserId = TestUserId,
            PlanId = 1,
            Status = "Active",
            StartDate = DateTime.Now.AddMonths(-2),
            AutoRenew = true,
            CreatedAt = DateTime.Now.AddMonths(-2)
        });

        _db.SubscriptionDeliveries.Add(new SubscriptionDelivery
        {
            DeliveryId = 1,
            SubscriptionId = 1,
            DeliveryDate = DateTime.Now.AddMonths(-1),
            Status = "Delivered",
            TrackingNo = "SF1234567890",
            CreatedAt = DateTime.Now.AddMonths(-1)
        });

        _db.ReferralTokens.Add(new ReferralToken
        {
            TokenId = 1,
            ReferrerUserId = TestUserId,
            TokenHash = "abc123def456",
            OriginalToken = "abc123def456",
            ReferrerType = "user",
            ExpiresAt = DateTime.Now.AddYears(1),
            MaxUses = 100,
            UsedCount = 1,
            IsActive = true,
            CreatedAt = DateTime.Now.AddDays(-30)
        });

        _db.ReferralRelations.Add(new ReferralRelation
        {
            RelationId = 1,
            AncestorUserId = TestUserId,
            DescendantUserId = 2,
            Depth = 1,
            CreatedAt = DateTime.Now.AddDays(-10)
        });

        _db.Orders.Add(new Order
        {
            OrderId = 1,
            OrderNo = "ORD-20250001",
            UserId = TestUserId,
            TotalAmount = 299m,
            Status = "Delivered",
            PaymentMethod = "微信支付",
            CreatedAt = DateTime.Now.AddDays(-20)
        });

        _db.SaveChanges();
    }

    public void Dispose()
    {
        _db.Database.EnsureDeleted();
        _db.Dispose();
    }

    // ==================== Helper ====================

    private static JsonElement ToJson(IActionResult result)
    {
        var ok = Assert.IsType<OkObjectResult>(result);
        var json = JsonSerializer.Serialize(ok.Value);
        return JsonSerializer.Deserialize<JsonElement>(json);
    }

    // ==================== 评价测试 ====================

    [Fact]
    public async Task Reviews_GetList_ReturnsUserReviews()
    {
        var result = await _ctrl.GetUserReviews(TestUserId);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());
        Assert.Equal(2, data.GetProperty("total").GetInt32());
    }

    [Fact]
    public async Task Reviews_Delete_SoftDeletesReview()
    {
        var result = await _ctrl.DeleteReview(TestUserId, 1);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());

        var review = await _db.ProductReviews.FindAsync(1);
        Assert.Equal("Deleted", review!.Status);
    }

    [Fact]
    public async Task Reviews_Delete_NotOwned_Returns404()
    {
        var result = await _ctrl.DeleteReview(999, 1);
        Assert.IsType<NotFoundObjectResult>(result);
    }

    // ==================== 订阅管理测试 ====================

    [Fact]
    public async Task Subscription_Pause_Success()
    {
        var result = await _ctrl.PauseSubscription(TestUserId, 1);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());

        var sub = await _db.UserSubscriptions.FindAsync(1);
        Assert.Equal("Paused", sub!.Status);
    }

    [Fact]
    public async Task Subscription_Pause_WhenPaused_ReturnsBadRequest()
    {
        await _ctrl.PauseSubscription(TestUserId, 1);
        var result = await _ctrl.PauseSubscription(TestUserId, 1);
        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public async Task Subscription_Resume_Success()
    {
        await _ctrl.PauseSubscription(TestUserId, 1);
        var result = await _ctrl.ResumeSubscription(TestUserId, 1);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());

        var sub = await _db.UserSubscriptions.FindAsync(1);
        Assert.Equal("Active", sub!.Status);
    }

    [Fact]
    public async Task Subscription_Cancel_Success()
    {
        var result = await _ctrl.CancelUserSubscription(TestUserId, 1, null);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());

        var sub = await _db.UserSubscriptions.FindAsync(1);
        Assert.Equal("Cancelled", sub!.Status);
        Assert.NotNull(sub.EndDate);
    }

    // ==================== KOL 商品测试 ====================

    [Fact]
    public async Task KolProducts_GetList_ReturnsKolProducts()
    {
        var result = await _ctrl.GetKolProducts(TestUserId);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());
        var items = data.GetProperty("data");
        Assert.True(items.GetArrayLength() >= 1);
    }

    [Fact]
    public async Task KolProducts_FilterByCategory_Works()
    {
        var result = await _ctrl.GetKolProducts(TestUserId, category: "KOL推荐");
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());
    }

    // ==================== 数据导出测试 ====================

    [Fact]
    public async Task DataExport_All_ReturnsFile()
    {
        var result = await _ctrl.ExportUserData(TestUserId, "all");
        var fileResult = Assert.IsType<FileContentResult>(result);
        Assert.Equal("text/csv", fileResult.ContentType);
        Assert.Contains("my_data_1", fileResult.FileDownloadName);
    }

    [Fact]
    public async Task DataExport_UserNotFound_Returns404()
    {
        var result = await _ctrl.ExportUserData(9999, "all");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    // ==================== 账户删除测试 ====================

    [Fact]
    public async Task AccountDelete_WrongPassword_ReturnsBadRequest()
    {
        var result = await _ctrl.RequestAccountDelete(TestUserId, new AccountDeleteRequest { Password = "wrongpwd" });
        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public async Task AccountDelete_Success_MarksUserInactive()
    {
        var result = await _ctrl.RequestAccountDelete(TestUserId, new AccountDeleteRequest { Password = "hashedpwd123" });
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());
        Assert.Equal(30, data.GetProperty("coolingDays").GetInt32());

        var user = await _db.Users.FindAsync(TestUserId);
        Assert.False(user!.IsActive);
    }

    // ==================== 推荐统计测试 ====================

    [Fact]
    public async Task Referrals_GetStats_ReturnsCorrectCounts()
    {
        var result = await _ctrl.GetReferralStats(TestUserId);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());
        Assert.Equal(1, data.GetProperty("totalInvites").GetInt32());
        Assert.Equal(100, data.GetProperty("totalRewardPoints").GetInt32());
    }

    [Fact]
    public async Task Referrals_GenerateLink_Success()
    {
        var result = await _ctrl.GenerateReferralLink(TestUserId);
        var data = ToJson(result);
        Assert.True(data.GetProperty("success").GetBoolean());
        Assert.False(string.IsNullOrEmpty(data.GetProperty("link").GetString()));
        Assert.False(string.IsNullOrEmpty(data.GetProperty("tokenHash").GetString()));
    }
}
