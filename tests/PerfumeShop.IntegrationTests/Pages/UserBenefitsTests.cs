using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using PerfumeShop.Api.Controllers;
using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests.Pages;

/// <summary>
/// M3-B User Benefits E2E Tests
/// Covers: Favorites, Points, Coupons, Addresses, Password Change (Settings)
/// </summary>
public class UserBenefitsTests : IDisposable
{
    private readonly TestEngineContext _db;
    private readonly UserExtendedController _ctrl;
    private const int TestUserId = 1;

    public UserBenefitsTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"UserBenefitsTests_{Guid.NewGuid()}")
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
            UserId = TestUserId, Username = "testuser", Email = "test@example.com",
            Password = PerfumeShop.Shared.Security.PasswordHasher.Hash("oldpwd123"), FullName = "TestUser", Phone = "13800000001",
            Points = 500, CustomerTier = "GoldMember", IsActive = true, CreatedAt = DateTime.Now
        });

        _db.Products.Add(new Product
        {
            ProductId = 1, ProductName = "RosePerfume", BasePrice = 299.00m,
            Category = "Floral", ImageUrl = "/images/products/rose.jpg",
            Description = "Classic rose fragrance", IsActive = true
        });
        _db.Products.Add(new Product
        {
            ProductId = 2, ProductName = "OceanPerfume", BasePrice = 399.00m,
            Category = "Ocean", ImageUrl = "/images/products/ocean.jpg",
            Description = "Fresh ocean scent", IsActive = true
        });

        _db.Coupons.Add(new Coupon
        {
            CouponId = 1, CouponName = "NewUserDiscount", CouponCode = "NEW100",
            CouponType = "fixed", DiscountValue = 100m, MinSpend = 500m, MaxDiscount = 100m,
            ValidFrom = DateTime.Now.AddDays(-1), ValidTo = DateTime.Now.AddDays(30),
            TotalQty = 100, UsedQty = 0, IsPublic = true, Description = "Spend 500 save 100"
        });
        _db.UserCoupons.Add(new UserCoupon
        {
            UserCouponId = 1, UserId = TestUserId, CouponId = 1,
            CouponCode = "NEW100", Source = "activity", Status = "available", ObtainedAt = DateTime.Now
        });

        _db.PointsRedemptions.Add(new PointsRedemption
        {
            RedemptionId = 1, ItemName = "50YuanCoupon", ItemType = "coupon",
            PointsCost = 200, Stock = 10, RedemptionValue = 50m,
            IsEnabled = true, SortOrder = 1, Description = "50 yuan coupon for 500 points"
        });

        _db.SaveChanges();
    }

    public void Dispose()
    {
        _db.Database.EnsureDeleted();
        _db.Dispose();
    }

    private static JsonElement GetJson(IActionResult result)
    {
        var ok = Assert.IsType<OkObjectResult>(result);
        var json = JsonSerializer.Serialize(ok.Value);
        return JsonSerializer.Deserialize<JsonElement>(json);
    }

    // ==================== Favorites Tests ====================

    [Fact]
    public async Task Favorites_AddAndGet_Success()
    {
        var addResult = await _ctrl.AddFavorite(TestUserId, new FavoriteRequest { ProductId = 1 });
        var addJson = GetJson(addResult);
        Assert.True(addJson.GetProperty("success").GetBoolean());

        var listResult = await _ctrl.GetFavorites(TestUserId);
        var listJson = GetJson(listResult);
        Assert.True(listJson.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Favorites_Remove_Success()
    {
        await _ctrl.AddFavorite(TestUserId, new FavoriteRequest { ProductId = 2 });
        var removeResult = await _ctrl.RemoveFavorite(TestUserId, 2);
        var json = GetJson(removeResult);
        Assert.True(json.GetProperty("success").GetBoolean());
    }

    // ==================== Points Tests ====================

    [Fact]
    public async Task Points_GetBalance_Returns500()
    {
        var result = await _ctrl.GetPointsBalance(TestUserId);
        var json = GetJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());
        Assert.Equal(500, json.GetProperty("points").GetInt32());
    }

    [Fact]
    public async Task Points_Redeem_Success()
    {
        var result = await _ctrl.RedeemPoints(TestUserId, new RedeemPointsRequest { RedemptionId = 1 });
        var json = GetJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());
        Assert.Equal(300, json.GetProperty("pointsRemaining").GetInt32());
    }

    [Fact]
    public async Task Points_Redeem_InsufficientPoints()
    {
        var user = await _db.Users.FindAsync(TestUserId);
        user!.Points = 100;
        await _db.SaveChangesAsync();

        var result = await _ctrl.RedeemPoints(TestUserId, new RedeemPointsRequest { RedemptionId = 1 });
        Assert.IsType<BadRequestObjectResult>(result);
    }

    // ==================== Coupons Tests ====================

    [Fact]
    public async Task Coupons_GetList_ReturnsAvailable()
    {
        var result = await _ctrl.GetUserCoupons(TestUserId, "available");
        var json = GetJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Coupons_Validate_ValidCode()
    {
        var result = await _ctrl.ValidateCoupon(TestUserId, new ValidateCouponRequest { Code = "NEW100" });
        var json = GetJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());
    }

    // ==================== Address CRUD Tests ====================

    [Fact]
    public async Task Address_AddAndGet_Success()
    {
        var addResult = await _ctrl.AddAddress(TestUserId, new AddressRequest
        {
            Consignee = "ZhangSan", Phone = "13800138000", Address = "Zhongguancun St No.1",
            Province = "Beijing", City = "Beijing", District = "Haidian", IsDefault = true
        });
        var addJson = GetJson(addResult);
        Assert.True(addJson.GetProperty("success").GetBoolean());

        var listResult = await _ctrl.GetAddresses(TestUserId);
        var listJson = GetJson(listResult);
        Assert.True(listJson.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Address_Delete_Success()
    {
        await _ctrl.AddAddress(TestUserId, new AddressRequest
        {
            Consignee = "LiSi", Phone = "13900139000", Address = "Nanjing Rd No.100",
            Province = "Shanghai", City = "Shanghai", District = "Huangpu"
        });
        var addresses = await _db.UserAddresses.Where(a => a.UserId == TestUserId).ToListAsync();
        var addrId = addresses.Last().AddressId;

        var deleteResult = await _ctrl.DeleteAddress(TestUserId, addrId);
        var json = GetJson(deleteResult);
        Assert.True(json.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Address_SetDefault_Success()
    {
        await _ctrl.AddAddress(TestUserId, new AddressRequest
        {
            Consignee = "WangWu", Phone = "13700137000",
            Address = "Tianhe Rd No.100", Province = "Guangdong", City = "Guangzhou", District = "Tianhe"
        });
        var addresses = await _db.UserAddresses.Where(a => a.UserId == TestUserId).ToListAsync();
        var addrId = addresses.Last().AddressId;

        var result = await _ctrl.SetDefaultAddress(TestUserId, addrId);
        var json = GetJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());
    }

    // ==================== Password Change Tests ====================

    [Fact]
    public async Task Password_Change_Success()
    {
        var result = await _ctrl.ChangePassword(TestUserId, new ChangePasswordRequest
        {
            CurrentPassword = "oldpwd123", NewPassword = "newpwd456"
        });
        var json = GetJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task Password_Change_TooShort()
    {
        var result = await _ctrl.ChangePassword(TestUserId, new ChangePasswordRequest
        {
            CurrentPassword = "oldpwd123", NewPassword = "123"
        });
        Assert.IsType<BadRequestObjectResult>(result);
    }

    // ==================== Profile Tests ====================

    [Fact]
    public async Task Profile_Update_Success()
    {
        var result = await _ctrl.UpdateProfile(TestUserId, new ProfileRequest
        {
            FullName = "NewName", Email = "newemail@example.com", Phone = "13600136000"
        });
        var json = GetJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());

        var user = await _db.Users.FindAsync(TestUserId);
        Assert.Equal("NewName", user!.FullName);
    }
}
