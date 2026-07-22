using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.IntegrationTests.Admin;

/// <summary>
/// 运营模块 29 个 Blazor 页面数据层测试
/// 验证 OperationRepository 各方法在 InMemory 数据库中可正常执行
/// </summary>
public class OperationPagesTests : IDisposable
{
    private readonly PerfumeShopContext _db;
    private readonly OperationRepository _repo;

    public OperationPagesTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"OpTests_{Guid.NewGuid()}")
            .Options;
        _db = new TestEngineContext(options);
        _repo = new OperationRepository(_db);
    }

    public void Dispose() => _db.Dispose();

    // ── Index (运营看板) ──
    [Fact]
    public async Task Index_GetSalesReport_ReturnsDto()
    {
        var result = await _repo.GetSalesReport();
        Assert.NotNull(result);
        Assert.NotNull(result.ChartLabels);
        Assert.NotNull(result.ChartData);
    }

    // ── Orders (订单管理) ──
    [Fact]
    public async Task Orders_GetOrdersPage_ReturnsTuple()
    {
        var (items, total) = await _repo.GetOrdersPage(null, null, 1, 10);
        Assert.NotNull(items);
        Assert.True(total >= 0);
    }

    // ── OrderDetail (订单详情) ──
    [Fact]
    public async Task OrderDetail_GetOrderDetail_ReturnsNullForMissing()
    {
        var result = await _repo.GetOrderDetail(99999);
        Assert.Null(result);
    }

    // ── OrderEdit (编辑订单) ──
    [Fact]
    public async Task OrderEdit_GetOrderItems_ReturnsList()
    {
        var result = await _repo.GetOrderItems(99999);
        Assert.NotNull(result);
    }

    // ── OrderReviews (订单评价) ──
    [Fact]
    public async Task OrderReviews_GetReviews_ReturnsTuple()
    {
        var (items, total) = await _repo.GetReviews(null, 1, 10);
        Assert.NotNull(items);
        Assert.True(total >= 0);
    }

    // ── Customers (客户管理) ──
    [Fact]
    public async Task Customers_GetCustomersPage_ReturnsTuple()
    {
        var (items, total) = await _repo.GetCustomersPage(null, null, 1, 10);
        Assert.NotNull(items);
        Assert.True(total >= 0);
    }

    // ── CustomerDetail (客户详情) ──
    [Fact]
    public async Task CustomerDetail_GetCustomerDetail_ReturnsNullForMissing()
    {
        var result = await _repo.GetCustomerDetail(99999);
        Assert.Null(result);
    }

    // ── Products (商品管理) ──
    [Fact]
    public async Task Products_GetProductsPage_ReturnsTuple()
    {
        var (items, total) = await _repo.GetProductsPage(null, null, null, 1, 10);
        Assert.NotNull(items);
        Assert.True(total >= 0);
    }

    // ── Categories (分类管理) ──
    [Fact]
    public async Task Categories_GetCategories_ReturnsList()
    {
        var result = await _repo.GetCategories();
        Assert.NotNull(result);
    }

    // ── FlashSale (秒杀管理) ──
    [Fact]
    public async Task FlashSale_GetFlashSales_ReturnsList()
    {
        var result = await _repo.GetFlashSales();
        Assert.NotNull(result);
    }

    // ── GroupBuy (拼团管理) ──
    [Fact]
    public async Task GroupBuy_GetGroupBuyPlans_ReturnsList()
    {
        var result = await _repo.GetGroupBuyPlans();
        Assert.NotNull(result);
    }

    // ── Coupons (优惠券) ──
    [Fact]
    public async Task Coupons_GetCoupons_ReturnsList()
    {
        var result = await _repo.GetCoupons();
        Assert.NotNull(result);
    }

    // ── CouponManagement (优惠券发放) ──
    [Fact]
    public async Task CouponManagement_GetUserCoupons_ReturnsList()
    {
        var result = await _repo.GetUserCoupons();
        Assert.NotNull(result);
    }

    // ── TierManagement (会员等级) ──
    [Fact]
    public async Task TierManagement_GetMemberTiers_ReturnsList()
    {
        var result = await _repo.GetMemberTiers();
        Assert.NotNull(result);
    }

    // ── Points (积分管理) ──
    [Fact]
    public async Task Points_GetPointsRules_ReturnsList()
    {
        var result = await _repo.GetPointsRules();
        Assert.NotNull(result);
    }

    // ── SubscriptionPlans (订阅管理) ──
    [Fact]
    public async Task SubscriptionPlans_GetSubscriptionPlans_ReturnsList()
    {
        var result = await _repo.GetSubscriptionPlans();
        Assert.NotNull(result);
    }

    // ── ReviewsManage (评价管理) ──
    [Fact]
    public async Task ReviewsManage_GetReviewsWithFilter_ReturnsTuple()
    {
        var (items, total) = await _repo.GetReviews("pending", 1, 10);
        Assert.NotNull(items);
        Assert.True(total >= 0);
    }

    // ── AfterSales (售后管理) ──
    [Fact]
    public async Task AfterSales_GetAfterSales_ReturnsTuple()
    {
        var (items, total) = await _repo.GetAfterSales(null, 1, 10);
        Assert.NotNull(items);
        Assert.True(total >= 0);
    }

    // ── Marketing (营销活动) ──
    [Fact]
    public async Task Marketing_GetMarketingCampaigns_ReturnsList()
    {
        var result = await _repo.GetMarketingCampaigns();
        Assert.NotNull(result);
    }

    // ── CampaignEdit (活动编辑) ──
    [Fact]
    public async Task CampaignEdit_CreateCampaign_ReturnsId()
    {
        var campaign = new MarketingCampaign { CampaignName = "Test", CampaignType = "test", CreatedAt = DateTime.Now };
        var id = await _repo.CreateCampaign(campaign);
        Assert.True(id > 0);
    }

    // ── ContentPages (内容页管理) ──
    [Fact]
    public async Task ContentPages_GetContentPages_ReturnsList()
    {
        var result = await _repo.GetContentPages();
        Assert.NotNull(result);
    }

    // ── ProductTypes (商品类型) ──
    [Fact]
    public async Task ProductTypes_GetProductTypes_ReturnsList()
    {
        var result = await _repo.GetProductTypes();
        Assert.NotNull(result);
    }

    // ── Fragrances (香料管理) ──
    [Fact]
    public async Task Fragrances_GetFragrances_ReturnsList()
    {
        var result = await _repo.GetFragrances();
        Assert.NotNull(result);
    }

    // ── BaseNotes (基调管理) ──
    [Fact]
    public async Task BaseNotes_GetBaseNotes_ReturnsList()
    {
        var result = await _repo.GetBaseNotes();
        Assert.NotNull(result);
    }

    // ── RecipeEdit (配方编辑) ──
    [Fact]
    public async Task RecipeEdit_GetRecipes_ReturnsList()
    {
        var result = await _repo.GetRecipes();
        Assert.NotNull(result);
    }

    // ── ReferralCodes (推荐码) ──
    [Fact]
    public async Task ReferralCodes_GetReferralCodes_ReturnsList()
    {
        var result = await _repo.GetReferralCodes();
        Assert.NotNull(result);
    }

    // ── PaymentSwitch (支付开关) ──
    [Fact]
    public async Task PaymentSwitch_GetPaymentRecords_ReturnsList()
    {
        var result = await _repo.GetPaymentRecords();
        Assert.NotNull(result);
    }

    // ── PerformanceDashboard (性能仪表板) ──
    [Fact]
    public async Task PerformanceDashboard_GetMetrics_ReturnsDto()
    {
        var result = await _repo.GetPerformanceMetrics();
        Assert.NotNull(result);
        Assert.True(result.TotalOrders >= 0);
        Assert.True(result.TotalRevenue >= 0);
    }

    // ── FixTierData (修复等级数据) ──
    [Fact]
    public async Task FixTierData_GetMemberBenefits_ReturnsList()
    {
        var result = await _repo.GetMemberBenefits();
        Assert.NotNull(result);
    }
}
