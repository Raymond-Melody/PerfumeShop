using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// 引擎测试专用 DbContext — 在 TestEngineContext 基础上增加 SiteSetting 主键
/// 解决 SiteSetting 作为 keyless 实体在 InMemory Provider 下无法 Add/Query 的问题
/// </summary>
public class EngineTestContext : PerfumeShopContext
{
    public EngineTestContext(DbContextOptions<PerfumeShopContext> options) : base(options) { }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // ====== 核心实体主键 ======
        modelBuilder.Entity<Product>().HasKey(e => e.ProductId);
        modelBuilder.Entity<Order>().HasKey(e => e.OrderId);
        modelBuilder.Entity<OrderDetail>().HasKey(e => e.DetailId);
        modelBuilder.Entity<User>().HasKey(e => e.UserId);
        modelBuilder.Entity<Coupon>().HasKey(e => e.CouponId);
        modelBuilder.Entity<UserCoupon>().HasKey(e => e.UserCouponId);
        modelBuilder.Entity<MemberTier>().HasKey(e => e.TierId);
        modelBuilder.Entity<PaymentRecord>().HasKey(e => e.RecordId);
        modelBuilder.Entity<RefundRecord>().HasKey(e => e.RefundId);
        modelBuilder.Entity<ProductionOrder>().HasKey(e => e.ProductionId);
        modelBuilder.Entity<ProductionLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<Recipe>().HasKey(e => e.RecipeId);
        modelBuilder.Entity<UserFavorite>().HasKey(e => e.FavoriteId);
        modelBuilder.Entity<AdminLog>().HasKey(e => e.LogId);

        // ====== SiteSetting — 解决 keyless 实体在 InMemory 下的主键问题 ======
        modelBuilder.Entity<SiteSetting>().HasKey(e => e.SettingKey);

        // ====== E2E: additional keyless entities ======
        modelBuilder.Entity<CommunityPost>().HasKey(e => e.PostId);
        modelBuilder.Entity<OrderItem>().HasKey(e => e.OrderItemId);
        modelBuilder.Entity<Supplier>().HasKey(e => e.SupplierId);
        modelBuilder.Entity<UserAddress>().HasKey(e => e.AddressId);
        modelBuilder.Entity<AccountsPayable>().HasKey(e => e.PayableId);
        modelBuilder.Entity<AccountsReceivable>().HasKey(e => e.ReceivableId);
        modelBuilder.Entity<ReconciliationLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<PurchaseOrder>().HasKey(e => e.PurchaseId);
        modelBuilder.Entity<PurchaseOrderDetail>().HasKey(e => e.DetailId);
        modelBuilder.Entity<PurchaseOrderStatusLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<PurchaseReceipt>().HasKey(e => e.ReceiptId);
        modelBuilder.Entity<PurchaseReceiptDetail>().HasKey(e => e.ReceiptDetailId);
        modelBuilder.Entity<ProductReview>().HasKey(e => e.ReviewId);
        modelBuilder.Entity<ReviewImage>().HasKey(e => e.ImageId);
        modelBuilder.Entity<UserSubscription>().HasKey(e => e.SubscriptionId);
        modelBuilder.Entity<SubscriptionPlan>().HasKey(e => e.PlanId);
        modelBuilder.Entity<SubscriptionDelivery>().HasKey(e => e.DeliveryId);
        modelBuilder.Entity<FlashSale>().HasKey(e => e.FlashSaleId);
        modelBuilder.Entity<GroupBuyPlan>().HasKey(e => e.PlanId);
        modelBuilder.Entity<GroupBuyOrder>().HasKey(e => e.GroupId);
        modelBuilder.Entity<GroupBuyParticipant>().HasKey(e => e.ParticipantId);
        modelBuilder.Entity<PointsRule>().HasKey(e => e.RuleId);
        modelBuilder.Entity<UserPoint>().HasKey(e => e.PointId);
        modelBuilder.Entity<PointsLedger>().HasKey(e => e.LedgerId);
        modelBuilder.Entity<PointsRedemption>().HasKey(e => e.RedemptionId);
        modelBuilder.Entity<PointTransaction>().HasKey(e => e.TransactionId);
        modelBuilder.Entity<DailyStatistic>().HasKey(e => e.StatId);
        modelBuilder.Entity<AppLog>().HasKey(e => e.LogId);
    }
}
