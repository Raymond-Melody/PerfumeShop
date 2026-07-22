using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 运营模块仓储 — 封装管理后台运营中心 29 个页面所需的数据查询
/// </summary>
public class OperationRepository
{
    private readonly PerfumeShopContext _context;

    public OperationRepository(PerfumeShopContext context)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
    }

    // ===================== 订单管理 =====================

    public async Task<(List<Order> Items, int Total)> GetOrdersPage(
        string? search, string? status, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.Orders.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(search))
            query = query.Where(o => o.OrderNo.Contains(search) || o.ShippingName!.Contains(search));
        if (!string.IsNullOrWhiteSpace(status))
            query = query.Where(o => o.Status == status);
        var total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(o => o.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task<Order?> GetOrderDetail(int orderId, CancellationToken ct = default)
        => await _context.Orders.FirstOrDefaultAsync(o => o.OrderId == orderId, ct);

    public async Task<List<OrderItem>> GetOrderItems(int orderId, CancellationToken ct = default)
        => await _context.OrderItems.Where(i => i.OrderId == orderId).ToListAsync(ct);

    public async Task UpdateOrderStatus(int orderId, string status, CancellationToken ct = default)
    {
        await _context.Orders.Where(o => o.OrderId == orderId)
            .ExecuteUpdateAsync(s => s.SetProperty(o => o.Status, status)
                .SetProperty(o => o.UpdatedAt, DateTime.Now), ct);
    }

    // ===================== 客户管理 =====================

    public async Task<(List<User> Items, int Total)> GetCustomersPage(
        string? search, bool? activeOnly, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.Users.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(search))
            query = query.Where(u => u.Username.Contains(search) || u.Email.Contains(search) || (u.FullName ?? "").Contains(search));
        if (activeOnly == true)
            query = query.Where(u => u.IsActive == true);
        var total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(u => u.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task<User?> GetCustomerDetail(int userId, CancellationToken ct = default)
        => await _context.Users.FirstOrDefaultAsync(u => u.UserId == userId, ct);

    // ===================== 商品管理 =====================

    public async Task<(List<Product> Items, int Total)> GetProductsPage(
        string? search, string? category, bool? activeOnly, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.Products.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(search))
            query = query.Where(p => p.ProductName.Contains(search));
        if (!string.IsNullOrWhiteSpace(category))
            query = query.Where(p => p.Category == category);
        if (activeOnly == true)
            query = query.Where(p => p.IsActive == true);
        var total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(p => p.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    // ===================== 分类管理 =====================

    public async Task<List<Product>> GetCategories(CancellationToken ct = default)
        => await _context.Products.AsNoTracking().Where(p => p.Category != null)
            .Select(p => new Product { Category = p.Category }).Distinct().ToListAsync(ct);

    // ===================== 秒杀管理 =====================

    public async Task<List<FlashSale>> GetFlashSales(CancellationToken ct = default)
        => await _context.FlashSales.AsNoTracking().OrderByDescending(f => f.CreatedAt).ToListAsync(ct);

    public async Task<FlashSale?> GetFlashSaleById(int id, CancellationToken ct = default)
        => await _context.FlashSales.FirstOrDefaultAsync(f => f.FlashSaleId == id, ct);

    public async Task<int> CreateFlashSale(FlashSale entity, CancellationToken ct = default)
    {
        entity.CreatedAt = DateTime.Now;
        _context.FlashSales.Add(entity);
        await _context.SaveChangesAsync(ct);
        return entity.FlashSaleId;
    }

    public async Task UpdateFlashSale(FlashSale entity, CancellationToken ct = default)
    {
        _context.FlashSales.Update(entity);
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeleteFlashSale(int id, CancellationToken ct = default)
    {
        var entity = await _context.FlashSales.FirstOrDefaultAsync(f => f.FlashSaleId == id, ct);
        if (entity != null) { _context.FlashSales.Remove(entity); await _context.SaveChangesAsync(ct); }
    }

    // ===================== 拼团管理 =====================

    public async Task<List<GroupBuyPlan>> GetGroupBuyPlans(CancellationToken ct = default)
        => await _context.GroupBuyPlans.AsNoTracking().OrderByDescending(p => p.CreatedAt).ToListAsync(ct);

    public async Task<List<GroupBuyOrder>> GetGroupBuyOrders(CancellationToken ct = default)
        => await _context.GroupBuyOrders.AsNoTracking().OrderByDescending(o => o.CreatedAt).Take(100).ToListAsync(ct);

    public async Task<List<GroupBuyParticipant>> GetGroupBuyParticipants(CancellationToken ct = default)
        => await _context.GroupBuyParticipants.AsNoTracking().OrderByDescending(p => p.JoinedAt).Take(100).ToListAsync(ct);

    // ===================== 优惠券管理 =====================

    public async Task<List<Coupon>> GetCoupons(CancellationToken ct = default)
        => await _context.Coupons.AsNoTracking().OrderByDescending(c => c.CreatedAt).ToListAsync(ct);

    public async Task<int> CreateCoupon(Coupon entity, CancellationToken ct = default)
    {
        entity.CreatedAt = DateTime.Now;
        _context.Coupons.Add(entity);
        await _context.SaveChangesAsync(ct);
        return entity.CouponId;
    }

    public async Task UpdateCoupon(Coupon entity, CancellationToken ct = default)
    {
        entity.UpdatedAt = DateTime.Now;
        _context.Coupons.Update(entity);
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeleteCoupon(int id, CancellationToken ct = default)
    {
        var entity = await _context.Coupons.FirstOrDefaultAsync(c => c.CouponId == id, ct);
        if (entity != null) { _context.Coupons.Remove(entity); await _context.SaveChangesAsync(ct); }
    }

    public async Task<List<UserCoupon>> GetUserCoupons(CancellationToken ct = default)
        => await _context.UserCoupons.AsNoTracking().OrderByDescending(uc => uc.ObtainedAt).Take(200).ToListAsync(ct);

    // ===================== 会员等级 =====================

    public async Task<List<MemberTier>> GetMemberTiers(CancellationToken ct = default)
        => await _context.MemberTiers.AsNoTracking().OrderBy(t => t.SortOrder).ToListAsync(ct);

    public async Task UpdateTier(MemberTier entity, CancellationToken ct = default)
    {
        entity.UpdatedAt = DateTime.Now;
        _context.MemberTiers.Update(entity);
        await _context.SaveChangesAsync(ct);
    }

    public async Task<List<MemberBenefit>> GetMemberBenefits(CancellationToken ct = default)
        => await _context.MemberBenefits.AsNoTracking().ToListAsync(ct);

    public async Task<List<UserTierHistory>> GetTierHistory(CancellationToken ct = default)
        => await _context.UserTierHistories.AsNoTracking().OrderByDescending(h => h.ChangedAt).Take(100).ToListAsync(ct);

    // ===================== 积分管理 =====================

    public async Task<List<PointsRule>> GetPointsRules(CancellationToken ct = default)
        => await _context.PointsRules.AsNoTracking().OrderBy(r => r.SortOrder).ToListAsync(ct);

    public async Task<List<PointTransaction>> GetPointsLedger(int? userId, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.PointTransactions.AsNoTracking().AsQueryable();
        if (userId.HasValue) query = query.Where(t => t.UserId == userId.Value);
        return await query.OrderByDescending(t => t.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
    }

    // ===================== 订阅管理 =====================

    public async Task<List<SubscriptionPlan>> GetSubscriptionPlans(CancellationToken ct = default)
        => await _context.SubscriptionPlans.AsNoTracking().OrderBy(s => s.SortOrder).ToListAsync(ct);

    public async Task<List<UserSubscription>> GetSubscriptions(CancellationToken ct = default)
        => await _context.UserSubscriptions.AsNoTracking().OrderByDescending(s => s.CreatedAt).Take(200).ToListAsync(ct);

    // ===================== 评价管理 =====================

    public async Task<(List<ProductReview> Items, int Total)> GetReviews(
        string? status, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.ProductReviews.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) query = query.Where(r => r.Status == status);
        var total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(r => r.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task UpdateReviewStatus(int reviewId, string status, CancellationToken ct = default)
    {
        await _context.ProductReviews.Where(r => r.ReviewId == reviewId)
            .ExecuteUpdateAsync(s => s.SetProperty(r => r.Status, status)
                .SetProperty(r => r.UpdatedAt, DateTime.Now), ct);
    }

    // ===================== 售后管理 =====================

    public async Task<(List<AfterSale> Items, int Total)> GetAfterSales(
        string? status, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.AfterSales.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) query = query.Where(a => a.Status == status);
        var total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(a => a.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task ProcessRefund(int afterSaleId, string adminNotes, decimal refundAmount, CancellationToken ct = default)
    {
        await _context.AfterSales.Where(a => a.AfterSalesId == afterSaleId)
            .ExecuteUpdateAsync(s => s.SetProperty(a => a.Status, "refunded")
                .SetProperty(a => a.AdminNotes, adminNotes)
                .SetProperty(a => a.ProcessedAt, DateTime.Now), ct);
    }

    // ===================== 营销活动 =====================

    public async Task<List<MarketingCampaign>> GetMarketingCampaigns(CancellationToken ct = default)
        => await _context.MarketingCampaigns.AsNoTracking().OrderByDescending(c => c.CreatedAt).ToListAsync(ct);

    public async Task<int> CreateCampaign(MarketingCampaign entity, CancellationToken ct = default)
    {
        entity.CreatedAt = DateTime.Now;
        _context.MarketingCampaigns.Add(entity);
        await _context.SaveChangesAsync(ct);
        return entity.CampaignId;
    }

    public async Task UpdateCampaign(MarketingCampaign entity, CancellationToken ct = default)
    {
        _context.MarketingCampaigns.Update(entity);
        await _context.SaveChangesAsync(ct);
    }

    // ===================== 内容页管理 =====================

    public async Task<List<ContentPage>> GetContentPages(CancellationToken ct = default)
        => await _context.ContentPages.AsNoTracking().OrderBy(c => c.SortOrder).ToListAsync(ct);

    // ===================== 商品类型 =====================

    public async Task<List<ProductTypeConfig>> GetProductTypes(CancellationToken ct = default)
        => await _context.ProductTypeConfigs.AsNoTracking().OrderBy(t => t.DisplayOrder).ToListAsync(ct);

    // ===================== 香料 / 基调 =====================

    public async Task<List<FragranceNote>> GetFragrances(CancellationToken ct = default)
        => await _context.FragranceNotes.AsNoTracking().ToListAsync(ct);

    public async Task<List<BaseNote>> GetBaseNotes(CancellationToken ct = default)
        => await _context.BaseNotes.AsNoTracking().ToListAsync(ct);

    // ===================== 推荐码 =====================

    public async Task<List<ReferralToken>> GetReferralCodes(CancellationToken ct = default)
        => await _context.ReferralTokens.AsNoTracking().OrderByDescending(t => t.CreatedAt).ToListAsync(ct);

    // ===================== 配方 =====================

    public async Task<List<Recipe>> GetRecipes(CancellationToken ct = default)
        => await _context.Recipes.AsNoTracking().OrderByDescending(r => r.CreatedAt).ToListAsync(ct);

    // ===================== 支付记录 =====================

    public async Task<List<PaymentRecord>> GetPaymentRecords(CancellationToken ct = default)
        => await _context.PaymentRecords.AsNoTracking().OrderByDescending(p => p.CreatedAt).Take(100).ToListAsync(ct);

    // ===================== 报表/看板 =====================

    public async Task<SalesReportDto> GetSalesReport(CancellationToken ct = default)
    {
        var today = DateTime.Today;
        var todayOrders = await _context.Orders.CountAsync(o => o.CreatedAt >= today, ct);
        var todayRevenue = await _context.Orders
            .Where(o => o.CreatedAt >= today && o.Status != "cancelled")
            .SumAsync(o => (decimal?)o.TotalAmount, ct) ?? 0;
        var todayNewUsers = await _context.Users.CountAsync(u => u.CreatedAt >= today, ct);

        var stats = await _context.DailyStatistics
            .Where(s => s.StatDate >= today.AddDays(-6))
            .OrderBy(s => s.StatDate).ToListAsync(ct);

        return new SalesReportDto
        {
            TodayOrders = todayOrders,
            TodayRevenue = todayRevenue,
            TodayNewUsers = todayNewUsers,
            ChartLabels = stats.Select(s => s.StatDate.ToString("MM-dd")).ToArray(),
            ChartData = stats.Select(s => (double)(s.TotalRevenue ?? 0)).ToArray()
        };
    }

    public async Task<PerformanceMetricsDto> GetPerformanceMetrics(CancellationToken ct = default)
    {
        var totalOrders = await _context.Orders.CountAsync(ct);
        var totalRevenue = await _context.Orders.Where(o => o.Status != "cancelled")
            .SumAsync(o => (decimal?)o.TotalAmount, ct) ?? 0;
        var totalUsers = await _context.Users.CountAsync(ct);
        var totalProducts = await _context.Products.CountAsync(p => p.IsActive == true, ct);

        return new PerformanceMetricsDto
        {
            TotalOrders = totalOrders,
            TotalRevenue = totalRevenue,
            TotalUsers = totalUsers,
            TotalProducts = totalProducts
        };
    }
}

// ===================== DTOs =====================

public class SalesReportDto
{
    public int TodayOrders { get; set; }
    public decimal TodayRevenue { get; set; }
    public int TodayNewUsers { get; set; }
    public string[] ChartLabels { get; set; } = Array.Empty<string>();
    public double[] ChartData { get; set; } = Array.Empty<double>();
}

public class PerformanceMetricsDto
{
    public int TotalOrders { get; set; }
    public decimal TotalRevenue { get; set; }
    public int TotalUsers { get; set; }
    public int TotalProducts { get; set; }
}
