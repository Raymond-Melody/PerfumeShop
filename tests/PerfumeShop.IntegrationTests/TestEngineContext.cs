using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests;

/// <summary>
/// 测试专用 DbContext — 为 InMemory Provider 重新配置 keyless 实体的主键
/// 覆盖 PromotionEngine / PaymentHandler / RecommendationEngine 测试所需的所有实体
/// </summary>
public class TestEngineContext : PerfumeShopContext
{
    public TestEngineContext(DbContextOptions<PerfumeShopContext> options) : base(options) { }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // 通用主键配置
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

        // CostEngine 相关
        modelBuilder.Entity<RawMaterialInventory>().HasKey(e => e.MaterialId);
        modelBuilder.Entity<SupplierPrice>().HasKey(e => e.PriceId);
        modelBuilder.Entity<FragranceNote>().HasKey(e => e.NoteId);
        modelBuilder.Entity<BottleStyle>().HasKey(e => e.BottleId);
        modelBuilder.Entity<ProductNoteRatio>().HasKey(e => new { e.ProductId, e.NoteId });
        modelBuilder.Entity<ProductBottleStyle>().HasKey(e => new { e.ProductId, e.BottleId });
        modelBuilder.Entity<ProductCost>().HasKey(e => new { e.ProductId, e.CostType });
        modelBuilder.Entity<RecipeAccord>().HasKey(e => e.AccordRecipeId);
        modelBuilder.Entity<RecipeAccordMaterial>().HasKey(e => e.DetailId);
        modelBuilder.Entity<NoteIngredient>().HasKey(e => new { e.NoteId, e.BaseNoteId });
        modelBuilder.Entity<BaseNote>().HasKey(e => e.BaseNoteId);
        modelBuilder.Entity<OrderCostAllocation>().HasKey(e => e.AllocationId);
        modelBuilder.Entity<FixedBrandProduct>().HasKey(e => e.FixedProductId);
        modelBuilder.Entity<FixedBrandInventory>().HasKey(e => e.InventoryId);

        // PointsEngine 相关
        modelBuilder.Entity<UserPoint>().HasKey(e => e.PointId);
        modelBuilder.Entity<PointsRule>().HasKey(e => e.RuleId);
        modelBuilder.Entity<PointsLedger>().HasKey(e => e.LedgerId);
        modelBuilder.Entity<PointsRedemption>().HasKey(e => e.RedemptionId);

        // M3-B: 地址管理
        modelBuilder.Entity<UserAddress>().HasKey(e => e.AddressId);

        // M3-A: PasswordResetToken
        modelBuilder.Entity<PasswordResetToken>().HasKey(e => e.TokenId);

        // M3-C: Reviews / Subscription / Referrals / Enhanced pages
        modelBuilder.Entity<ProductReview>().HasKey(e => e.ReviewId);
        modelBuilder.Entity<ReviewImage>().HasKey(e => e.ImageId);
        modelBuilder.Entity<UserSubscription>().HasKey(e => e.SubscriptionId);
        modelBuilder.Entity<SubscriptionPlan>().HasKey(e => e.PlanId);
        modelBuilder.Entity<SubscriptionDelivery>().HasKey(e => e.DeliveryId);
        modelBuilder.Entity<ReferralToken>().HasKey(e => e.TokenId);
        modelBuilder.Entity<ReferralRelation>().HasKey(e => e.RelationId);

        // M4-A: Operation module
        modelBuilder.Entity<MarketingCampaign>().HasKey(e => e.CampaignId);
        modelBuilder.Entity<ContentPage>().HasKey(e => e.PageId);
        modelBuilder.Entity<AfterSale>().HasKey(e => e.AfterSalesId);
        modelBuilder.Entity<PointTransaction>().HasKey(e => e.TransactionId);
        modelBuilder.Entity<DailyStatistic>().HasKey(e => e.StatId);
        modelBuilder.Entity<PaymentRecord>().HasKey(e => e.RecordId);
        modelBuilder.Entity<GroupBuyPlan>().HasKey(e => e.PlanId);
        modelBuilder.Entity<GroupBuyOrder>().HasKey(e => e.GroupId);
        modelBuilder.Entity<GroupBuyParticipant>().HasKey(e => e.ParticipantId);
        modelBuilder.Entity<FlashSale>().HasKey(e => e.FlashSaleId);
        modelBuilder.Entity<UserTierHistory>().HasKey(e => e.HistoryId);
        modelBuilder.Entity<MemberBenefit>().HasKey(e => e.BenefitId);

        // M4-B: 采购模块
        modelBuilder.Entity<PurchaseOrder>().HasKey(e => e.PurchaseId);
        modelBuilder.Entity<PurchaseOrderDetail>().HasKey(e => e.DetailId);
        modelBuilder.Entity<PurchaseOrderStatusLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<PurchaseReceipt>().HasKey(e => e.ReceiptId);
        modelBuilder.Entity<PurchaseReceiptDetail>().HasKey(e => e.ReceiptDetailId);
        modelBuilder.Entity<Supplier>().HasKey(e => e.SupplierId);
        modelBuilder.Entity<SupplierEvaluation>().HasKey(e => e.EvaluationId);
        modelBuilder.Entity<SupplierContract>().HasKey(e => e.ContractId);
        modelBuilder.Entity<PurchaseBatch>().HasKey(e => e.BatchId);
        modelBuilder.Entity<FixedBrandPurchaseOrder>().HasKey(e => e.PurchaseId);
        modelBuilder.Entity<FixedBrandCostAllocation>().HasKey(e => e.AllocationId);
        modelBuilder.Entity<PurchaseCostReview>().HasKey(e => e.ReviewId);
        modelBuilder.Entity<PurchaseHistoryStat>().HasKey(e => e.StatId);
        modelBuilder.Entity<PurchaseCategory>().HasKey(e => e.CategoryId);
        modelBuilder.Entity<PackagingInventory>().HasKey(e => e.PackagingId);
        modelBuilder.Entity<BottleInventory>().HasKey(e => e.BottleId);

        // M4-D: TechCenter 相关
        modelBuilder.Entity<RecipeNote>().HasKey(e => e.Id);
        modelBuilder.Entity<RecipeIngredient>().HasKey(e => e.Id);
        modelBuilder.Entity<NoteInventory>().HasKey(e => e.InventoryId);
        modelBuilder.Entity<Ingredient>().HasKey(e => e.IngredientId);
        modelBuilder.Entity<RecipeProduct>().HasKey(e => e.ProductRecipeId);
        modelBuilder.Entity<RecipeProductNote>().HasKey(e => e.DetailId);
        modelBuilder.Entity<RecipePublishLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<InventoryTransaction>().HasKey(e => e.TransactionId);
        modelBuilder.Entity<ProductTypeConfig>().HasKey(e => e.ConfigId);
        modelBuilder.Entity<Volume>().HasKey(e => e.VolumeId);
        modelBuilder.Entity<ProductVolumePrice>().HasKey(e => e.PvpriceId);

        // M5-A: System / Inventory / Logistics / Finance entities
        modelBuilder.Entity<AdminUser>().HasKey(e => e.AdminId);
        modelBuilder.Entity<AdminRole>().HasKey(e => e.RoleId);
        modelBuilder.Entity<AdminAuditLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<AppLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<LoginAlert>().HasKey(e => e.AlertId);
        modelBuilder.Entity<ModulePermission>().HasKey(e => e.PermissionId);
        modelBuilder.Entity<RolePermission>().HasKey(e => e.PermId);
        modelBuilder.Entity<AccountsPayable>().HasKey(e => e.PayableId);
        modelBuilder.Entity<AccountsReceivable>().HasKey(e => e.ReceivableId);
        modelBuilder.Entity<BudgetPlan>().HasKey(e => e.BudgetId);
        modelBuilder.Entity<ReconciliationLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<ExpenseRecord>().HasKey(e => e.ExpenseId);
        modelBuilder.Entity<CostCenter>().HasKey(e => e.CenterId);
        modelBuilder.Entity<Gltransaction>().HasKey(e => e.Glid);
        modelBuilder.Entity<ShippingCompany>().HasKey(e => e.CompanyId);
        modelBuilder.Entity<AfterSale>().HasKey(e => e.AfterSalesId);
        modelBuilder.Entity<RawMaterialInventory>().HasKey(e => e.MaterialId);
        modelBuilder.Entity<StockMovement>().HasKey(e => e.MovementId);
        modelBuilder.Entity<DailyStatistic>().HasKey(e => e.StatId);
        modelBuilder.Entity<ProductInventory>().HasKey(e => e.InventoryId);
        modelBuilder.Entity<InventoryBatch>().HasKey(e => e.BatchId);
        modelBuilder.Entity<SprayHeadInventory>().HasKey(e => e.SprayHeadId);
        modelBuilder.Entity<PrintingInventory>().HasKey(e => e.PrintingId);
        modelBuilder.Entity<SiteSetting>().HasKey(e => e.SettingKey);

        // E2E: CommunityPost / OrderItem
        modelBuilder.Entity<CommunityPost>().HasKey(e => e.PostId);
        modelBuilder.Entity<OrderItem>().HasKey(e => e.OrderItemId);
    }
}
