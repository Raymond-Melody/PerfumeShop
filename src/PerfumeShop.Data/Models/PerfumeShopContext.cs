using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace PerfumeShop.Data.Models;

public partial class PerfumeShopContext : DbContext
{
    public PerfumeShopContext(DbContextOptions<PerfumeShopContext> options)
        : base(options)
    {
    }

    public virtual DbSet<AccordProduction> AccordProductions { get; set; }

    public virtual DbSet<AccordProductionDetail> AccordProductionDetails { get; set; }

    public virtual DbSet<AccordQcreport> AccordQcreports { get; set; }

    public virtual DbSet<AccountsPayable> AccountsPayables { get; set; }

    public virtual DbSet<AccountsReceivable> AccountsReceivables { get; set; }

    public virtual DbSet<AdminAuditLog> AdminAuditLogs { get; set; }

    public virtual DbSet<AdminLog> AdminLogs { get; set; }

    public virtual DbSet<AdminRole> AdminRoles { get; set; }

    public virtual DbSet<AdminUser> AdminUsers { get; set; }

    public virtual DbSet<AfterSale> AfterSales { get; set; }

    public virtual DbSet<AppLog> AppLogs { get; set; }

    public virtual DbSet<Area> Areas { get; set; }

    public virtual DbSet<AuthToken> AuthTokens { get; set; }

    public virtual DbSet<BaseNote> BaseNotes { get; set; }

    public virtual DbSet<BottleInventory> BottleInventories { get; set; }

    public virtual DbSet<BottleStyle> BottleStyles { get; set; }

    public virtual DbSet<BudgetPlan> BudgetPlans { get; set; }

    public virtual DbSet<Cart> Carts { get; set; }

    public virtual DbSet<CartNoteSelection> CartNoteSelections { get; set; }

    public virtual DbSet<Category> Categories { get; set; }

    public virtual DbSet<CommunityPost> CommunityPosts { get; set; }

    public virtual DbSet<ContentPage> ContentPages { get; set; }

    public virtual DbSet<CostCenter> CostCenters { get; set; }

    public virtual DbSet<Coupon> Coupons { get; set; }

    public virtual DbSet<DailyStatistic> DailyStatistics { get; set; }

    public virtual DbSet<ExpenseRecord> ExpenseRecords { get; set; }

    public virtual DbSet<FixedBrandCostAllocation> FixedBrandCostAllocations { get; set; }

    public virtual DbSet<FixedBrandInventory> FixedBrandInventories { get; set; }

    public virtual DbSet<FixedBrandProduct> FixedBrandProducts { get; set; }

    public virtual DbSet<FixedBrandPurchaseDetail> FixedBrandPurchaseDetails { get; set; }

    public virtual DbSet<FixedBrandPurchaseOrder> FixedBrandPurchaseOrders { get; set; }

    public virtual DbSet<FixedBrandReceipt> FixedBrandReceipts { get; set; }

    public virtual DbSet<FixedBrandReceiptDetail> FixedBrandReceiptDetails { get; set; }

    public virtual DbSet<FlashSale> FlashSales { get; set; }

    public virtual DbSet<Formula> Formulas { get; set; }

    public virtual DbSet<FormulaNote> FormulaNotes { get; set; }

    public virtual DbSet<FragranceIngredient> FragranceIngredients { get; set; }

    public virtual DbSet<FragranceNote> FragranceNotes { get; set; }

    public virtual DbSet<FundAccount> FundAccounts { get; set; }

    public virtual DbSet<Gltransaction> Gltransactions { get; set; }

    public virtual DbSet<GroupBuyOrder> GroupBuyOrders { get; set; }

    public virtual DbSet<GroupBuyParticipant> GroupBuyParticipants { get; set; }

    public virtual DbSet<GroupBuyPlan> GroupBuyPlans { get; set; }

    public virtual DbSet<Ingredient> Ingredients { get; set; }

    public virtual DbSet<InventoryBatch> InventoryBatches { get; set; }

    public virtual DbSet<InventoryTransaction> InventoryTransactions { get; set; }

    public virtual DbSet<Ipblacklist> Ipblacklists { get; set; }

    public virtual DbSet<LoginAlert> LoginAlerts { get; set; }

    public virtual DbSet<MarketingCampaign> MarketingCampaigns { get; set; }

    public virtual DbSet<MaterialOutbound> MaterialOutbounds { get; set; }

    public virtual DbSet<MaterialOutboundDetail> MaterialOutboundDetails { get; set; }

    public virtual DbSet<MemberBenefit> MemberBenefits { get; set; }

    public virtual DbSet<MemberTier> MemberTiers { get; set; }

    public virtual DbSet<ModulePermission> ModulePermissions { get; set; }

    public virtual DbSet<NoteIngredient> NoteIngredients { get; set; }

    public virtual DbSet<NoteInventory> NoteInventories { get; set; }

    public virtual DbSet<Order> Orders { get; set; }

    public virtual DbSet<OrderCostAllocation> OrderCostAllocations { get; set; }

    public virtual DbSet<OrderDetail> OrderDetails { get; set; }

    public virtual DbSet<OrderDetailNoteSelection> OrderDetailNoteSelections { get; set; }

    public virtual DbSet<OrderIngredient> OrderIngredients { get; set; }

    public virtual DbSet<OrderItem> OrderItems { get; set; }

    public virtual DbSet<PackagingInventory> PackagingInventories { get; set; }

    public virtual DbSet<PaymentRecord> PaymentRecords { get; set; }

    public virtual DbSet<PointTransaction> PointTransactions { get; set; }

    public virtual DbSet<PointsLedger> PointsLedgers { get; set; }

    public virtual DbSet<PointsRedemption> PointsRedemptions { get; set; }

    public virtual DbSet<PointsRule> PointsRules { get; set; }

    public virtual DbSet<PostComment> PostComments { get; set; }

    public virtual DbSet<PostLike> PostLikes { get; set; }

    public virtual DbSet<PriceChangeLog> PriceChangeLogs { get; set; }

    public virtual DbSet<PrintingInventory> PrintingInventories { get; set; }

    public virtual DbSet<Product> Products { get; set; }

    public virtual DbSet<ProductBottleStyle> ProductBottleStyles { get; set; }

    public virtual DbSet<ProductCost> ProductCosts { get; set; }

    public virtual DbSet<ProductImage> ProductImages { get; set; }

    public virtual DbSet<ProductInventory> ProductInventories { get; set; }

    public virtual DbSet<ProductManufacturing> ProductManufacturings { get; set; }

    public virtual DbSet<ProductManufacturingDetail> ProductManufacturingDetails { get; set; }

    public virtual DbSet<ProductNote> ProductNotes { get; set; }

    public virtual DbSet<ProductNoteRatio> ProductNoteRatios { get; set; }

    public virtual DbSet<ProductReview> ProductReviews { get; set; }

    public virtual DbSet<ProductTypeConfig> ProductTypeConfigs { get; set; }

    public virtual DbSet<ProductVolumePrice> ProductVolumePrices { get; set; }

    public virtual DbSet<ProductionLog> ProductionLogs { get; set; }

    public virtual DbSet<ProductionOrder> ProductionOrders { get; set; }

    public virtual DbSet<PurchaseBatch> PurchaseBatches { get; set; }

    public virtual DbSet<PurchaseCategory> PurchaseCategories { get; set; }

    public virtual DbSet<PurchaseCostReview> PurchaseCostReviews { get; set; }

    public virtual DbSet<PurchaseHistoryStat> PurchaseHistoryStats { get; set; }

    public virtual DbSet<PurchaseOrder> PurchaseOrders { get; set; }

    public virtual DbSet<PurchaseOrderDetail> PurchaseOrderDetails { get; set; }

    public virtual DbSet<PurchaseOrderStatusLog> PurchaseOrderStatusLogs { get; set; }

    public virtual DbSet<PurchaseReceipt> PurchaseReceipts { get; set; }

    public virtual DbSet<PurchaseReceiptDetail> PurchaseReceiptDetails { get; set; }

    public virtual DbSet<RawMaterialInventory> RawMaterialInventories { get; set; }

    public virtual DbSet<Recipe> Recipes { get; set; }

    public virtual DbSet<RecipeAccord> RecipeAccords { get; set; }

    public virtual DbSet<RecipeAccordMaterial> RecipeAccordMaterials { get; set; }

    public virtual DbSet<RecipeIngredient> RecipeIngredients { get; set; }

    public virtual DbSet<RecipeNote> RecipeNotes { get; set; }

    public virtual DbSet<RecipePopularity> RecipePopularities { get; set; }

    public virtual DbSet<RecipeProduct> RecipeProducts { get; set; }

    public virtual DbSet<RecipeProductNote> RecipeProductNotes { get; set; }

    public virtual DbSet<RecipePublishLog> RecipePublishLogs { get; set; }

    public virtual DbSet<RecommendedRecipe> RecommendedRecipes { get; set; }

    public virtual DbSet<ReconciliationLog> ReconciliationLogs { get; set; }

    public virtual DbSet<ReferralRelation> ReferralRelations { get; set; }

    public virtual DbSet<ReferralToken> ReferralTokens { get; set; }

    public virtual DbSet<RefundRecord> RefundRecords { get; set; }

    public virtual DbSet<RegistrationAttempt> RegistrationAttempts { get; set; }

    public virtual DbSet<ReviewLike> ReviewLikes { get; set; }
    
    public virtual DbSet<ReviewImage> ReviewImages { get; set; }

    public virtual DbSet<RolePermission> RolePermissions { get; set; }

    public virtual DbSet<ShippingCompany> ShippingCompanies { get; set; }

    public virtual DbSet<SiteSetting> SiteSettings { get; set; }

    public virtual DbSet<SprayHeadInventory> SprayHeadInventories { get; set; }

    public virtual DbSet<StockMovement> StockMovements { get; set; }

    public virtual DbSet<SubscriptionDelivery> SubscriptionDeliveries { get; set; }

    public virtual DbSet<SubscriptionPlan> SubscriptionPlans { get; set; }

    public virtual DbSet<Supplier> Suppliers { get; set; }

    public virtual DbSet<SupplierContract> SupplierContracts { get; set; }

    public virtual DbSet<SupplierEvaluation> SupplierEvaluations { get; set; }

    public virtual DbSet<SupplierPrice> SupplierPrices { get; set; }

    public virtual DbSet<TierConfigLog> TierConfigLogs { get; set; }

    public virtual DbSet<User> Users { get; set; }

    public virtual DbSet<UserAddress> UserAddresses { get; set; }

    public virtual DbSet<UserCoupon> UserCoupons { get; set; }

    public virtual DbSet<UserFavorite> UserFavorites { get; set; }

    public virtual DbSet<UserPoint> UserPoints { get; set; }

    public virtual DbSet<UserPreference> UserPreferences { get; set; }

    public virtual DbSet<UserSubscription> UserSubscriptions { get; set; }

    public virtual DbSet<UserTierHistory> UserTierHistories { get; set; }

    public virtual DbSet<Volume> Volumes { get; set; }

    public virtual DbSet<WorkshopTransfer> WorkshopTransfers { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<AccordProduction>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.AccordRecipeId).HasColumnName("AccordRecipeID");
            entity.Property(e => e.ApprovedBy).HasMaxLength(50);
            entity.Property(e => e.BatchNo).HasMaxLength(30);
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.NoteName).HasMaxLength(100);
            entity.Property(e => e.ProductionId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ProductionID");
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.WorkCenter).HasMaxLength(20);
        });

        modelBuilder.Entity<AccordProductionDetail>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.DetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("DetailID");
            entity.Property(e => e.MaterialId).HasColumnName("MaterialID");
            entity.Property(e => e.MaterialName).HasMaxLength(100);
            entity.Property(e => e.ProductionId).HasColumnName("ProductionID");
            entity.Property(e => e.TotalCost).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitCost).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<AccordQcreport>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("AccordQCReports");

            entity.Property(e => e.BatchNo).HasMaxLength(30);
            entity.Property(e => e.ProductionId).HasColumnName("ProductionID");
            entity.Property(e => e.QcreportId)
                .ValueGeneratedOnAdd()
                .HasColumnName("QCReportID");
            entity.Property(e => e.Qcresult)
                .HasMaxLength(20)
                .HasColumnName("QCResult");
            entity.Property(e => e.TesterId).HasColumnName("TesterID");
            entity.Property(e => e.TesterName).HasMaxLength(50);
        });

        modelBuilder.Entity<AccountsPayable>(entity =>
        {
            entity.HasKey(e => e.PayableId).HasName("PK__Accounts__97CCDB3A2D183470");

            entity.ToTable("AccountsPayable");

            entity.HasIndex(e => e.DueDate, "IX_AccountsPayable_DueDate");

            entity.HasIndex(e => e.Status, "IX_AccountsPayable_Status");

            entity.Property(e => e.PayableId).HasColumnName("PayableID");
            entity.Property(e => e.Amount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 2)");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.DueDate).HasColumnType("datetime");
            entity.Property(e => e.PaidAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 2)");
            entity.Property(e => e.PayableNo).HasMaxLength(50);
            entity.Property(e => e.Status)
                .HasMaxLength(20)
                .HasDefaultValue("Pending");
            entity.Property(e => e.SupplierName).HasMaxLength(200);
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<AccountsReceivable>(entity =>
        {
            entity.HasKey(e => e.ReceivableId).HasName("PK__Accounts__AEA43F3C1985BA06");

            entity.ToTable("AccountsReceivable");

            entity.HasIndex(e => e.DueDate, "IX_AccountsReceivable_DueDate");

            entity.HasIndex(e => e.Status, "IX_AccountsReceivable_Status");

            entity.Property(e => e.ReceivableId).HasColumnName("ReceivableID");
            entity.Property(e => e.Amount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 2)");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.CustomerName).HasMaxLength(200);
            entity.Property(e => e.DueDate).HasColumnType("datetime");
            entity.Property(e => e.ReceivableNo).HasMaxLength(50);
            entity.Property(e => e.ReceivedAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Status)
                .HasMaxLength(20)
                .HasDefaultValue("Pending");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<AdminAuditLog>(entity =>
        {
            entity.HasKey(e => e.LogId).HasName("PK__AdminAud__5E5499A8E200B508");

            entity.ToTable("AdminAuditLog");

            entity.Property(e => e.LogId).HasColumnName("LogID");
            entity.Property(e => e.ActionType).HasMaxLength(50);
            entity.Property(e => e.AdminId).HasColumnName("AdminID");
            entity.Property(e => e.AdminName).HasMaxLength(100);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Ipaddress)
                .HasMaxLength(50)
                .HasColumnName("IPAddress");
            entity.Property(e => e.TargetId).HasColumnName("TargetID");
            entity.Property(e => e.TargetName).HasMaxLength(200);
            entity.Property(e => e.TargetType).HasMaxLength(50);
            entity.Property(e => e.UserAgent).HasMaxLength(500);
        });

        modelBuilder.Entity<AdminLog>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.CreatedAt, "IX_AdminLogs_CreatedAt");

            entity.HasIndex(e => new { e.ModuleCode, e.CreatedAt }, "IX_AdminLogs_ModuleCode");

            entity.Property(e => e.ActionType).HasMaxLength(100);
            entity.Property(e => e.AdminId).HasColumnName("AdminID");
            entity.Property(e => e.Ipaddress)
                .HasMaxLength(50)
                .HasColumnName("IPAddress");
            entity.Property(e => e.LogId)
                .ValueGeneratedOnAdd()
                .HasColumnName("LogID");
            entity.Property(e => e.ModuleCode).HasMaxLength(50);
            entity.Property(e => e.RecordId)
                .HasMaxLength(50)
                .HasColumnName("RecordID");
            entity.Property(e => e.TableName).HasMaxLength(50);
        });

        modelBuilder.Entity<AdminRole>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.ModuleAccess).HasMaxLength(500);
            entity.Property(e => e.RoleCode).HasMaxLength(20);
            entity.Property(e => e.RoleId)
                .ValueGeneratedOnAdd()
                .HasColumnName("RoleID");
            entity.Property(e => e.RoleName).HasMaxLength(50);
        });

        modelBuilder.Entity<AdminUser>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.AdminId)
                .ValueGeneratedOnAdd()
                .HasColumnName("AdminID");
            entity.Property(e => e.Department).HasMaxLength(50);
            entity.Property(e => e.Email).HasMaxLength(100);
            entity.Property(e => e.FullName).HasMaxLength(100);
            entity.Property(e => e.ResetToken).HasMaxLength(255);
            entity.Property(e => e.RoleId).HasColumnName("RoleID");
            entity.Property(e => e.Username).HasMaxLength(50);
        });

        modelBuilder.Entity<AfterSale>(entity =>
        {
            entity.HasKey(e => e.AfterSalesId).HasName("PK__AfterSal__30D0362A1164FDE8");

            entity.Property(e => e.AfterSalesId).HasColumnName("AfterSalesID");
            entity.Property(e => e.AdminNotes).HasMaxLength(500);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.ProcessedAt).HasColumnType("datetime");
            entity.Property(e => e.Reason).HasMaxLength(500);
            entity.Property(e => e.RefundAmount).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.RequestType).HasMaxLength(20);
            entity.Property(e => e.Status)
                .HasMaxLength(20)
                .HasDefaultValue("pending");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<AppLog>(entity =>
        {
            entity.HasKey(e => e.LogId);

            entity.HasIndex(e => e.CreatedAt, "IX_AppLogs_CreatedAt").IsDescending();

            entity.HasIndex(e => new { e.CreatedAt, e.LogType }, "IX_AppLogs_CreatedAt_LogType").IsDescending(true, false);

            entity.HasIndex(e => new { e.LogLevel, e.CreatedAt }, "IX_AppLogs_Level").IsDescending(false, true);

            // V19: v18_perf_indexes.sql — 过滤索引（仅ERROR/WARN级别）
            entity.HasIndex(e => new { e.LogLevel, e.CreatedAt }, "IX_AppLogs_Level_Date")
                .IsDescending(false, true)
                .HasFilter("([LogLevel]='ERROR' OR [LogLevel]='WARN')");

            entity.Property(e => e.LogId).HasColumnName("LogID");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.Ipaddress)
                .HasMaxLength(50)
                .HasColumnName("IPAddress");
            entity.Property(e => e.LogLevel).HasMaxLength(10);
            entity.Property(e => e.LogMessage).HasMaxLength(500);
            entity.Property(e => e.LogSource).HasMaxLength(100);
            entity.Property(e => e.LogType).HasMaxLength(50);
            entity.Property(e => e.PageUrl)
                .HasMaxLength(200)
                .HasColumnName("PageURL");
            entity.Property(e => e.UserName).HasMaxLength(100);
        });

        modelBuilder.Entity<Area>(entity =>
        {
            entity.HasKey(e => e.AreaId);
            entity.ToTable("Areas");
            entity.Property(e => e.AreaId).HasColumnName("AreaID");
            entity.Property(e => e.AreaName).HasMaxLength(100);
            entity.Property(e => e.ParentId).HasColumnName("ParentID");
        });

        modelBuilder.Entity<AuthToken>(entity =>
        {
            entity.HasKey(e => e.TokenId);

            entity.ToTable("AuthTokens");

            entity.HasIndex(e => new { e.Token, e.ExpiresAt }, "IX_AuthTokens_Token_ExpiresAt");
            entity.HasIndex(e => e.UserId, "IX_AuthTokens_UserId");

            entity.Property(e => e.TokenId).HasColumnName("TokenID");
            entity.Property(e => e.UserId).HasColumnName("UserID");
            entity.Property(e => e.Token).HasMaxLength(128).IsRequired();
            entity.Property(e => e.Source).HasMaxLength(20).IsRequired();
            entity.Property(e => e.IpAddress).HasMaxLength(50);
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(GETUTCDATE())");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
        });

        modelBuilder.Entity<BaseNote>(entity =>
        {
            entity.HasKey(e => e.BaseNoteId);

            entity.Property(e => e.BaseNoteId)
                .ValueGeneratedOnAdd()
                .HasColumnName("BaseNoteID");
            entity.Property(e => e.BaseNoteName).HasMaxLength(100);
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<BottleInventory>(entity =>
        {
            entity.HasKey(e => e.BottleId).HasName("PK__BottleIn__05EC40A1AC8E518B");

            entity.ToTable("BottleInventory");

            entity.Property(e => e.BottleId).HasColumnName("BottleID");
            entity.Property(e => e.BottleName).HasMaxLength(100);
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.SafetyStock)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.StockQty)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<BottleStyle>(entity =>
        {
            entity.HasKey(e => e.BottleId);

            entity.HasIndex(e => e.BottleName, "IX_BottleStyles_Name");

            entity.HasIndex(e => new { e.StockQty, e.SafetyStock }, "IX_BottleStyles_StockSafety");

            entity.Property(e => e.AvgDailyUsage)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 6)");
            entity.Property(e => e.BottleId)
                .ValueGeneratedOnAdd()
                .HasColumnName("BottleID");
            entity.Property(e => e.BottleName).HasMaxLength(50);
            entity.Property(e => e.BottleType)
                .HasMaxLength(30)
                .HasDefaultValue("Standard");
            entity.Property(e => e.CapacityMl)
                .HasDefaultValue(50)
                .HasColumnName("CapacityML");
            entity.Property(e => e.ImageUrl)
                .HasMaxLength(200)
                .HasColumnName("ImageURL");
            entity.Property(e => e.LastPurchaseDate).HasColumnType("datetime");
            entity.Property(e => e.LastReplenishDate).HasColumnType("datetime");
            entity.Property(e => e.LeadTimeDays).HasDefaultValue(7);
            entity.Property(e => e.PriceAddition).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ReorderPoint)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.SafetyStock)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.StockQty)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UnitCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UpdatedAt).HasColumnType("datetime");
        });

        modelBuilder.Entity<BudgetPlan>(entity =>
        {
            entity.HasKey(e => e.BudgetId);

            entity.HasIndex(e => new { e.Category, e.Period }, "IX_BudgetPlans_Category_Period");

            entity.Property(e => e.ActualAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.AlertRoi).HasColumnName("AlertROI");
            entity.Property(e => e.BudgetAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.BudgetId)
                .ValueGeneratedOnAdd()
                .HasColumnName("BudgetID");
            entity.Property(e => e.BudgetName).HasMaxLength(100);
            entity.Property(e => e.Category).HasMaxLength(50);
            entity.Property(e => e.CreatedBy).HasMaxLength(50);
            entity.Property(e => e.Gmvamount)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("GMVAmount");
            entity.Property(e => e.Period).HasMaxLength(10);
            entity.Property(e => e.Roi).HasColumnName("ROI");
            entity.Property(e => e.Status).HasMaxLength(20);
        });

        modelBuilder.Entity<Cart>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("Cart");

            entity.Property(e => e.BaseNoteId).HasColumnName("BaseNoteID");
            entity.Property(e => e.BottleId).HasColumnName("BottleID");
            entity.Property(e => e.CartId)
                .ValueGeneratedOnAdd()
                .HasColumnName("CartID");
            entity.Property(e => e.CustomLabel).HasMaxLength(200);
            entity.Property(e => e.MiddleNoteId).HasColumnName("MiddleNoteID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.SessionId)
                .HasMaxLength(100)
                .HasColumnName("SessionID");
            entity.Property(e => e.TopNoteId).HasColumnName("TopNoteID");
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UserId).HasColumnName("UserID");
            entity.Property(e => e.VolumeId).HasColumnName("VolumeID");
        });

        modelBuilder.Entity<CartNoteSelection>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CartId).HasColumnName("CartID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.NoteType).HasMaxLength(20);
            entity.Property(e => e.SelectionId)
                .ValueGeneratedOnAdd()
                .HasColumnName("SelectionID");
        });

        modelBuilder.Entity<Category>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CategoryId)
                .ValueGeneratedOnAdd()
                .HasColumnName("CategoryID");
            entity.Property(e => e.CategoryName).HasMaxLength(100);
        });

        modelBuilder.Entity<CommunityPost>(entity =>
        {
            entity.HasKey(e => e.PostId).HasName("PK__Communit__AA1260384F5C1F3F");

            entity.Property(e => e.PostId).HasColumnName("PostID");
            entity.Property(e => e.Content).HasMaxLength(4000);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.FragranceNotes).HasMaxLength(500);
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.IsPublic).HasDefaultValue(true);
            entity.Property(e => e.PostType)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasDefaultValue("discussion");
            entity.Property(e => e.Tags).HasMaxLength(300);
            entity.Property(e => e.Title).HasMaxLength(200);
            entity.Property(e => e.UpdatedAt).HasColumnType("datetime");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<ContentPage>(entity =>
        {
            entity.HasKey(e => e.PageId).HasName("PK__ContentP__C565B1243BA4D2EC");

            entity.HasIndex(e => e.Slug, "UQ__ContentP__BC7B5FB6F65C15C5").IsUnique();

            entity.Property(e => e.PageId).HasColumnName("PageID");
            entity.Property(e => e.Content).HasColumnType("ntext");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.IsPublished).HasDefaultValue(false);
            entity.Property(e => e.MetaDescription).HasMaxLength(500);
            entity.Property(e => e.Slug).HasMaxLength(200);
            entity.Property(e => e.SortOrder).HasDefaultValue(0);
            entity.Property(e => e.Title).HasMaxLength(200);
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<CostCenter>(entity =>
        {
            entity.HasKey(e => e.CenterId).HasName("PK__CostCent__398FC7D76BD060FA");

            entity.HasIndex(e => e.CenterCode, "IX_CostCenters_Code");

            entity.HasIndex(e => e.CenterCode, "UQ__CostCent__55D5E3C64DD1CD27").IsUnique();

            entity.Property(e => e.CenterId).HasColumnName("CenterID");
            entity.Property(e => e.BudgetAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.CenterCode).HasMaxLength(50);
            entity.Property(e => e.CenterName).HasMaxLength(200);
            entity.Property(e => e.CenterType)
                .HasMaxLength(50)
                .HasDefaultValue("Department");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.ParentId).HasColumnName("ParentID");
            entity.Property(e => e.UpdatedAt).HasDefaultValueSql("(getdate())");
        });

        modelBuilder.Entity<Coupon>(entity =>
        {
            entity.Property(e => e.CouponId).HasColumnName("CouponID");
            entity.Property(e => e.ApplicableCategory).HasMaxLength(50);
            entity.Property(e => e.ApplicableProductId).HasColumnName("ApplicableProductID");
            entity.Property(e => e.CouponCode).HasMaxLength(50);
            entity.Property(e => e.CouponName)
                .HasMaxLength(100)
                .HasDefaultValue("");
            entity.Property(e => e.CouponType)
                .HasMaxLength(20)
                .HasDefaultValue("fixed");
            entity.Property(e => e.Description)
                .HasMaxLength(500)
                .HasDefaultValue("");
            entity.Property(e => e.DiscountType).HasMaxLength(20);
            entity.Property(e => e.DiscountValue).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.IsPublic).HasDefaultValue(true);
            entity.Property(e => e.MaxDiscount).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.MinPurchase).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.MinSpend).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.Terms)
                .HasMaxLength(500)
                .HasDefaultValue("");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ValidFrom)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ValidTo)
                .HasDefaultValueSql("(dateadd(year,(1),getdate()))")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<DailyStatistic>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.DataJson).HasColumnName("DataJSON");
            entity.Property(e => e.StatId)
                .ValueGeneratedOnAdd()
                .HasColumnName("StatID");
            entity.Property(e => e.TopNoteId).HasColumnName("TopNoteID");
            entity.Property(e => e.TopProductId).HasColumnName("TopProductID");
            entity.Property(e => e.TotalRevenue).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<ExpenseRecord>(entity =>
        {
            // V20: ExpenseID为IDENTITY列，必须配置主键以支持Add/ExecuteDelete写操作（HasNoKey导致写入抛异常）
            entity.HasKey(e => e.ExpenseId);

            entity.HasIndex(e => e.Period, "IX_ExpenseRecords_Period");

            entity.Property(e => e.AllocationMethod).HasMaxLength(20);
            entity.Property(e => e.Amount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.CenterId).HasColumnName("CenterID");
            entity.Property(e => e.ExpenseId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ExpenseID");
            entity.Property(e => e.ExpenseName).HasMaxLength(100);
            entity.Property(e => e.ExpenseType).HasMaxLength(30);
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.Period).HasMaxLength(10);
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.SourceOrderId).HasColumnName("SourceOrderID");
        });

        modelBuilder.Entity<FixedBrandCostAllocation>(entity =>
        {
            entity.HasKey(e => e.AllocationId).HasName("PK__FixedBra__B3C6D6AB3AD7F053");

            entity.ToTable("FixedBrandCostAllocation");

            entity.Property(e => e.AllocationId).HasColumnName("AllocationID");
            entity.Property(e => e.AllocatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.CostPerUnit)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.FixedProductId).HasColumnName("FixedProductID");
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.OrderNo).HasMaxLength(100);
            entity.Property(e => e.ProductName).HasMaxLength(200);
            entity.Property(e => e.ProfitAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ProfitRate)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 4)");
            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.PurchaseNo).HasMaxLength(50);
            entity.Property(e => e.Quantity).HasDefaultValue(0);
            entity.Property(e => e.SalePrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.TotalCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<FixedBrandInventory>(entity =>
        {
            entity.HasKey(e => e.InventoryId).HasName("PK__FixedBra__F5FDE6D3B1538B39");

            entity.ToTable("FixedBrandInventory");

            entity.Property(e => e.InventoryId).HasColumnName("InventoryID");
            entity.Property(e => e.AvgUnitCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ConsecutiveDataMonths).HasDefaultValue(0);
            entity.Property(e => e.DailySalesAvg)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.FixedProductId).HasColumnName("FixedProductID");
            entity.Property(e => e.LastAutoCalcDate).HasColumnType("datetime");
            entity.Property(e => e.LastPurchaseDate).HasColumnType("datetime");
            entity.Property(e => e.LastPurchaseId).HasColumnName("LastPurchaseID");
            entity.Property(e => e.LastPurchasePrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.MinOrderQty).HasDefaultValue(1);
            entity.Property(e => e.ParamMode)
                .HasMaxLength(20)
                .HasDefaultValue("Manual");
            entity.Property(e => e.ProductCode).HasMaxLength(50);
            entity.Property(e => e.ProductName).HasMaxLength(200);
            entity.Property(e => e.SafetyStock).HasDefaultValue(10);
            entity.Property(e => e.Specification).HasMaxLength(100);
            entity.Property(e => e.StockQty).HasDefaultValue(0);
            entity.Property(e => e.TotalPurchased).HasDefaultValue(0);
            entity.Property(e => e.TotalSold).HasDefaultValue(0);
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<FixedBrandProduct>(entity =>
        {
            entity.HasKey(e => e.FixedProductId).HasName("PK__FixedBra__FD210BCBB087D7E4");

            entity.Property(e => e.FixedProductId).HasColumnName("FixedProductID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ImageUrl)
                .HasMaxLength(500)
                .HasColumnName("ImageURL");
            entity.Property(e => e.LeadTimeDays).HasDefaultValue(7);
            entity.Property(e => e.LeadTimeDaysManual).HasDefaultValue(7);
            entity.Property(e => e.MinOrderQty).HasDefaultValue(1);
            entity.Property(e => e.ProductCode).HasMaxLength(50);
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.ProductName).HasMaxLength(200);
            entity.Property(e => e.SafetyStockManual).HasDefaultValue(10);
            entity.Property(e => e.SalePrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.Specification).HasMaxLength(100);
            entity.Property(e => e.Status)
                .HasMaxLength(20)
                .HasDefaultValue("Active");
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.SupplierName).HasMaxLength(200);
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<FixedBrandPurchaseDetail>(entity =>
        {
            entity.HasKey(e => e.DetailId).HasName("PK__FixedBra__135C314D095CCE6A");

            entity.Property(e => e.DetailId).HasColumnName("DetailID");
            entity.Property(e => e.ExpectedDate).HasColumnType("datetime");
            entity.Property(e => e.FixedProductId).HasColumnName("FixedProductID");
            entity.Property(e => e.ProductName).HasMaxLength(200);
            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.Quantity).HasDefaultValue(0);
            entity.Property(e => e.ReceivedQty).HasDefaultValue(0);
            entity.Property(e => e.Remarks).HasMaxLength(500);
            entity.Property(e => e.Specification).HasMaxLength(100);
            entity.Property(e => e.SubTotal)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<FixedBrandPurchaseOrder>(entity =>
        {
            entity.HasKey(e => e.PurchaseId).HasName("PK__FixedBra__6B0A6BDE38FDAACF");

            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.ApprovedAt).HasColumnType("datetime");
            entity.Property(e => e.ApprovedBy).HasMaxLength(100);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.CreatedBy).HasMaxLength(100);
            entity.Property(e => e.ExpectedDate).HasColumnType("datetime");
            entity.Property(e => e.OrderDate)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.PurchaseNo).HasMaxLength(50);
            entity.Property(e => e.Remarks).HasMaxLength(500);
            entity.Property(e => e.Status)
                .HasMaxLength(30)
                .HasDefaultValue("Draft");
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.SupplierName).HasMaxLength(200);
            entity.Property(e => e.TotalAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<FixedBrandReceipt>(entity =>
        {
            entity.HasKey(e => e.ReceiptId).HasName("PK__FixedBra__CC08C4005A6CC60C");

            entity.Property(e => e.ReceiptId).HasColumnName("ReceiptID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Notes).HasMaxLength(500);
            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.ReceiptDate)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ReceiptNo).HasMaxLength(50);
            entity.Property(e => e.ReceivedBy).HasMaxLength(100);
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.TotalReceivedQty).HasDefaultValue(0);
        });

        modelBuilder.Entity<FixedBrandReceiptDetail>(entity =>
        {
            entity.HasKey(e => e.ReceiptDetailId).HasName("PK__FixedBra__82FADEDB88AA4833");

            entity.Property(e => e.ReceiptDetailId).HasColumnName("ReceiptDetailID");
            entity.Property(e => e.AcceptedQty).HasDefaultValue(0);
            entity.Property(e => e.DetailId).HasColumnName("DetailID");
            entity.Property(e => e.FixedProductId).HasColumnName("FixedProductID");
            entity.Property(e => e.ReceiptId).HasColumnName("ReceiptID");
            entity.Property(e => e.RejectReason).HasMaxLength(500);
            entity.Property(e => e.RejectedQty).HasDefaultValue(0);
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<FlashSale>(entity =>
        {
            entity.HasKey(e => e.FlashSaleId).HasName("PK__FlashSal__D603A204D04A3E0F");

            entity.ToTable("FlashSale");

            entity.Property(e => e.FlashSaleId).HasColumnName("FlashSaleID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.EndTime).HasColumnType("datetime");
            entity.Property(e => e.FlashPrice).HasColumnType("decimal(12, 2)");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.LimitPerUser).HasDefaultValue(1);
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.StartTime).HasColumnType("datetime");
        });

        modelBuilder.Entity<Formula>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.FormulaId)
                .ValueGeneratedOnAdd()
                .HasColumnName("FormulaID");
            entity.Property(e => e.FormulaName).HasMaxLength(100);
        });

        modelBuilder.Entity<FormulaNote>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.FormulaId).HasColumnName("FormulaID");
            entity.Property(e => e.Id)
                .ValueGeneratedOnAdd()
                .HasColumnName("ID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
        });

        modelBuilder.Entity<FragranceIngredient>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.FragranceIngredientId)
                .ValueGeneratedOnAdd()
                .HasColumnName("FragranceIngredientID");
            entity.Property(e => e.IngredientId).HasColumnName("IngredientID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
        });

        modelBuilder.Entity<FragranceNote>(entity =>
        {
            entity.HasKey(e => e.NoteId);

            entity.HasIndex(e => e.NoteType, "IX_FragranceNotes_NoteType");

            entity.Property(e => e.BaseNoteId).HasColumnName("BaseNoteID");
            entity.Property(e => e.ImageUrl)
                .HasMaxLength(200)
                .HasColumnName("ImageURL");
            entity.Property(e => e.NoteId)
                .ValueGeneratedOnAdd()
                .HasColumnName("NoteID");
            entity.Property(e => e.NoteName).HasMaxLength(50);
            entity.Property(e => e.NoteType).HasMaxLength(20);
            entity.Property(e => e.PriceAddition).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<FundAccount>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.AccountId)
                .ValueGeneratedOnAdd()
                .HasColumnName("AccountID");
            entity.Property(e => e.AccountName).HasMaxLength(100);
            entity.Property(e => e.AccountType).HasMaxLength(30);
            entity.Property(e => e.AlertThreshold).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.AvailableBalance).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.FrozenAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.PendingSettlement).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.TotalBalance).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<Gltransaction>(entity =>
        {
            entity.HasKey(e => e.Glid).HasName("PK__GLTransa__5F756C57CED9CC41");

            entity.ToTable("GLTransactions");

            entity.HasIndex(e => new { e.CenterId, e.TransactionDate }, "IX_GLTransactions_CenterID");

            entity.HasIndex(e => e.TransactionDate, "IX_GLTransactions_Date");

            entity.HasIndex(e => e.Glno, "IX_GLTransactions_GLNo");

            entity.Property(e => e.Glid).HasColumnName("GLID");
            entity.Property(e => e.AccountCode).HasMaxLength(50);
            entity.Property(e => e.AccountName).HasMaxLength(100);
            entity.Property(e => e.CenterId).HasColumnName("CenterID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.CreatedBy).HasMaxLength(50);
            entity.Property(e => e.CreditAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 2)");
            entity.Property(e => e.DebitAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Description).HasMaxLength(500);
            entity.Property(e => e.Glno)
                .HasMaxLength(50)
                .HasColumnName("GLNo");
            entity.Property(e => e.RefId).HasColumnName("RefID");
            entity.Property(e => e.RefNo).HasMaxLength(50);
            entity.Property(e => e.RefType).HasMaxLength(50);
            entity.Property(e => e.TransactionDate).HasColumnType("datetime");
        });

        modelBuilder.Entity<GroupBuyOrder>(entity =>
        {
            entity.HasKey(e => e.GroupId).HasName("PK__GroupBuy__149AF30A8A34754C");

            entity.Property(e => e.GroupId).HasColumnName("GroupID");
            entity.Property(e => e.CompletedAt).HasColumnType("datetime");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.CurrentSize).HasDefaultValue(1);
            entity.Property(e => e.GroupSn)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasColumnName("GroupSN");
            entity.Property(e => e.InitiatorId).HasColumnName("InitiatorID");
            entity.Property(e => e.PlanId).HasColumnName("PlanID");
        });

        modelBuilder.Entity<GroupBuyParticipant>(entity =>
        {
            entity.HasKey(e => e.ParticipantId).HasName("PK__GroupBuy__7227997EE4800C4E");

            entity.Property(e => e.ParticipantId).HasColumnName("ParticipantID");
            entity.Property(e => e.GroupId).HasColumnName("GroupID");
            entity.Property(e => e.JoinedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<GroupBuyPlan>(entity =>
        {
            entity.HasKey(e => e.PlanId).HasName("PK__GroupBuy__755C22D7536B16D4");

            entity.Property(e => e.PlanId).HasColumnName("PlanID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.DurationHours).HasDefaultValue(24);
            entity.Property(e => e.EndTime).HasColumnType("datetime");
            entity.Property(e => e.GroupPrice).HasColumnType("decimal(12, 2)");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.MinUnit).HasDefaultValue(1);
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.StartTime).HasColumnType("datetime");
            entity.Property(e => e.TeamSize).HasDefaultValue(2);
        });

        modelBuilder.Entity<Ingredient>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.Casnumber)
                .HasMaxLength(50)
                .HasColumnName("CASNumber");
            entity.Property(e => e.Description).HasMaxLength(255);
            entity.Property(e => e.IngredientId)
                .ValueGeneratedOnAdd()
                .HasColumnName("IngredientID");
            entity.Property(e => e.IngredientName).HasMaxLength(100);
        });

        modelBuilder.Entity<InventoryBatch>(entity =>
        {
            entity.HasKey(e => e.BatchId).HasName("PK__Inventor__5D55CE3860F1FB58");

            entity.Property(e => e.BatchId).HasColumnName("BatchID");
            entity.Property(e => e.BatchNo).HasMaxLength(100);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemId).HasColumnName("ItemID");
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.ItemType).HasMaxLength(30);
            entity.Property(e => e.StockQty).HasDefaultValue(0.0);
            entity.Property(e => e.UnitCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UpdatedAt).HasColumnType("datetime");
        });

        modelBuilder.Entity<InventoryTransaction>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CreatedBy).HasMaxLength(50);
            entity.Property(e => e.MaterialId).HasColumnName("MaterialID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.ReferenceOrderId).HasColumnName("ReferenceOrderID");
            entity.Property(e => e.ReferenceType).HasMaxLength(50);
            entity.Property(e => e.TransactionDirection).HasMaxLength(10);
            entity.Property(e => e.TransactionId)
                .ValueGeneratedOnAdd()
                .HasColumnName("TransactionID");
            entity.Property(e => e.TransactionType).HasMaxLength(20);
            entity.Property(e => e.UnitCost).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<Ipblacklist>(entity =>
        {
            entity.HasKey(e => e.Ipid).HasName("PK__IPBlackl__8FB9622A26229E2B");

            entity.ToTable("IPBlacklist");

            entity.HasIndex(e => e.BlockedAt, "IX_IPBlacklist_BlockedAt");

            entity.HasIndex(e => new { e.Ipaddress, e.IsActive }, "IX_IPBlacklist_IPAddress");

            entity.Property(e => e.Ipid).HasColumnName("IPID");
            entity.Property(e => e.BlockedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.HitCount).HasDefaultValue(0);
            entity.Property(e => e.Ipaddress)
                .HasMaxLength(50)
                .HasColumnName("IPAddress");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.Reason).HasMaxLength(255);
        });

        modelBuilder.Entity<LoginAlert>(entity =>
        {
            entity.HasKey(e => e.AlertId).HasName("PK__LoginAle__EBB16AED52C7B7B8");

            entity.HasIndex(e => e.CreatedAt, "IX_LoginAlerts_CreatedAt");

            entity.HasIndex(e => new { e.IsRead, e.CreatedAt }, "IX_LoginAlerts_IsRead");

            entity.Property(e => e.AlertId).HasColumnName("AlertID");
            entity.Property(e => e.AdminId).HasColumnName("AdminID");
            entity.Property(e => e.AlertLevel)
                .HasMaxLength(20)
                .HasDefaultValue("info");
            entity.Property(e => e.AlertMessage).HasMaxLength(500);
            entity.Property(e => e.AlertType).HasMaxLength(50);
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.Ipaddress)
                .HasMaxLength(50)
                .HasColumnName("IPAddress");
            entity.Property(e => e.IsRead).HasDefaultValue(false);
        });

        modelBuilder.Entity<MarketingCampaign>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CampaignId)
                .ValueGeneratedOnAdd()
                .HasColumnName("CampaignID");
            entity.Property(e => e.CampaignName).HasMaxLength(200);
            entity.Property(e => e.CampaignType).HasMaxLength(50);
            entity.Property(e => e.DiscountValue).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.MinPurchase).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.TotalSales).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<MaterialOutbound>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("MaterialOutbound");

            entity.Property(e => e.ApprovedBy).HasMaxLength(50);
            entity.Property(e => e.OutboundId)
                .ValueGeneratedOnAdd()
                .HasColumnName("OutboundID");
            entity.Property(e => e.OutboundNo).HasMaxLength(50);
            entity.Property(e => e.OutboundType).HasMaxLength(20);
            entity.Property(e => e.ReferenceId).HasColumnName("ReferenceID");
            entity.Property(e => e.ReferenceType).HasMaxLength(50);
            entity.Property(e => e.RequestedBy).HasMaxLength(50);
            entity.Property(e => e.Status).HasMaxLength(20);
        });

        modelBuilder.Entity<MaterialOutboundDetail>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.MaterialId).HasColumnName("MaterialID");
            entity.Property(e => e.OutboundDetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("OutboundDetailID");
            entity.Property(e => e.OutboundId).HasColumnName("OutboundID");
            entity.Property(e => e.TotalAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<MemberBenefit>(entity =>
        {
            entity.HasKey(e => e.BenefitId).HasName("PK__MemberBe__5754C53A7C5C867A");

            entity.Property(e => e.BenefitId).HasColumnName("BenefitID");
            entity.Property(e => e.BenefitDesc).HasMaxLength(500);
            entity.Property(e => e.BenefitIcon)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasDefaultValue("fa-check-circle");
            entity.Property(e => e.BenefitName).HasMaxLength(100);
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.TierCode)
                .HasMaxLength(20)
                .IsUnicode(false);
        });

        modelBuilder.Entity<MemberTier>(entity =>
        {
            entity.HasKey(e => e.TierId).HasName("PK__MemberTi__362F55FD01056D59");

            entity.HasIndex(e => e.TierCode, "UQ__MemberTi__07886BFECE658FD3").IsUnique();

            entity.Property(e => e.TierId).HasColumnName("TierID");
            entity.Property(e => e.BadgeBg)
                .HasMaxLength(20)
                .IsUnicode(false);
            entity.Property(e => e.Color)
                .HasMaxLength(20)
                .IsUnicode(false);
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.DiscountRate)
                .HasDefaultValue(1.000m)
                .HasColumnType("decimal(4, 3)");
            entity.Property(e => e.IconClass)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.MaxSpent).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.MinSpent).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.TierCode)
                .HasMaxLength(20)
                .IsUnicode(false);
            entity.Property(e => e.TierName).HasMaxLength(50);
            entity.Property(e => e.TierNameEn)
                .HasMaxLength(50)
                .HasColumnName("TierNameEN");
        });

        modelBuilder.Entity<ModulePermission>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.ModuleCode).HasMaxLength(50);
            entity.Property(e => e.ModuleName).HasMaxLength(100);
            entity.Property(e => e.ParentModule).HasMaxLength(50);
            entity.Property(e => e.PermissionId)
                .ValueGeneratedOnAdd()
                .HasColumnName("PermissionID");
            entity.Property(e => e.RequiredRole).HasMaxLength(20);
            entity.Property(e => e.Urlpattern)
                .HasMaxLength(200)
                .HasColumnName("URLPattern");
        });

        modelBuilder.Entity<NoteIngredient>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.BaseNoteId).HasColumnName("BaseNoteID");
            entity.Property(e => e.Id)
                .ValueGeneratedOnAdd()
                .HasColumnName("ID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
        });

        modelBuilder.Entity<NoteInventory>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("NoteInventory");

            entity.HasIndex(e => e.NoteId, "IX_NoteInventory_NoteID");

            entity.Property(e => e.InventoryId)
                .ValueGeneratedOnAdd()
                .HasColumnName("InventoryID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.WeightedUnitCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<Order>(entity =>
        {
            entity.HasKey(e => e.OrderId).HasName("PK_Orders");

            entity.HasIndex(e => e.CreatedAt, "IX_Orders_CreatedAt");

            entity.HasIndex(e => e.OrderNo, "IX_Orders_OrderNo");

            entity.HasIndex(e => e.Status, "IX_Orders_ProfitAmount_Status");

            entity.HasIndex(e => e.Status, "IX_Orders_Status");

            entity.HasIndex(e => e.UserId, "IX_Orders_UserID");

            // V19: v18_perf_indexes.sql 复合索引 — 用户订单列表 + 后台状态管理
            entity.HasIndex(e => new { e.UserId, e.CreatedAt }, "IX_Orders_UserID_CreatedAt").IsDescending(false, true);
            entity.HasIndex(e => new { e.Status, e.CreatedAt }, "IX_Orders_Status_CreatedAt").IsDescending(false, true);

            entity.Property(e => e.ChannelSource).HasMaxLength(50);
            entity.Property(e => e.CostAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ExpenseAmount)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.OrderId)
                .ValueGeneratedOnAdd()
                .HasColumnName("OrderID");
            entity.Property(e => e.OrderNo).HasMaxLength(50);
            entity.Property(e => e.PaymentMethod).HasMaxLength(50);
            entity.Property(e => e.CouponCode).HasMaxLength(50);
            entity.Property(e => e.CouponDiscount).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.PointsDiscount).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.ProfitAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.RefundAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ShippingAddress).HasMaxLength(200);
            entity.Property(e => e.ShippingCity).HasMaxLength(50);
            entity.Property(e => e.ShippingCompany).HasMaxLength(50);
            entity.Property(e => e.ShippingFee).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ShippingName).HasMaxLength(100);
            entity.Property(e => e.ShippingPhone).HasMaxLength(20);
            entity.Property(e => e.ShippingPostalCode).HasMaxLength(20);
            entity.Property(e => e.ShippingStatus).HasMaxLength(20);
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.TotalAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.TrackingNumber).HasMaxLength(100);
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<OrderCostAllocation>(entity =>
        {
            entity.HasKey(e => e.AllocationId).HasName("PK__OrderCos__B3C6D6ABD312B48B");

            entity.ToTable("OrderCostAllocation");

            entity.Property(e => e.AllocationId).HasColumnName("AllocationID");
            entity.Property(e => e.AllocatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.BatchId).HasColumnName("BatchID");
            entity.Property(e => e.CostType)
                .HasMaxLength(30)
                .HasDefaultValue("Material");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.InvBatchId).HasColumnName("InvBatchID");
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.OrderNo).HasMaxLength(100);
            entity.Property(e => e.Quantity).HasDefaultValue(0.0);
            entity.Property(e => e.TotalCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<OrderDetail>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.OrderId, "IX_OrderDetails_OrderID");

            entity.HasIndex(e => e.ProductId, "IX_OrderDetails_ProductID");

            entity.Property(e => e.BaseNoteName).HasMaxLength(100);
            entity.Property(e => e.BottleName).HasMaxLength(100);
            entity.Property(e => e.CustomLabel).HasMaxLength(200);
            entity.Property(e => e.DetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("DetailID");
            entity.Property(e => e.MiddleNoteName).HasMaxLength(100);
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.ProductName).HasMaxLength(200);
            entity.Property(e => e.Subtotal).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.TopNoteName).HasMaxLength(100);
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.VolumeMl).HasColumnName("VolumeML");
            entity.Property(e => e.VolumeName).HasMaxLength(50);
        });

        modelBuilder.Entity<OrderDetailNoteSelection>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.DetailId).HasColumnName("DetailID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.NoteType).HasMaxLength(20);
            entity.Property(e => e.SelectionId)
                .ValueGeneratedOnAdd()
                .HasColumnName("SelectionID");
        });

        modelBuilder.Entity<OrderIngredient>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.DetailId).HasColumnName("DetailID");
            entity.Property(e => e.IngredientId)
                .ValueGeneratedOnAdd()
                .HasColumnName("IngredientID");
            entity.Property(e => e.IngredientName).HasMaxLength(100);
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
        });

        modelBuilder.Entity<OrderItem>(entity =>
        {
            entity.HasKey(e => e.OrderItemId).HasName("PK__OrderIte__57ED06A1CA81CDDE");

            entity.Property(e => e.OrderItemId).HasColumnName("OrderItemID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.Quantity).HasDefaultValue(1);
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");

            // V19: v18_perf_indexes.sql — OrderID+ProductID 复合索引
            entity.HasIndex(e => new { e.OrderId, e.ProductId }, "IX_OrderItems_OrderID_ProductID");
        });

        modelBuilder.Entity<PackagingInventory>(entity =>
        {
            entity.HasKey(e => e.PackagingId).HasName("PK__Packagin__BD507F584A3A9121");

            entity.ToTable("PackagingInventory");

            entity.HasIndex(e => new { e.ItemName, e.ItemCode }, "IX_Packaging_Name_Code");

            entity.HasIndex(e => new { e.StockQty, e.SafetyStock }, "IX_Packaging_StockSafety");

            entity.Property(e => e.PackagingId).HasColumnName("PackagingID");
            entity.Property(e => e.AvgDailyUsage)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 6)");
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(100);
            entity.Property(e => e.LastReplenishDate).HasColumnType("datetime");
            entity.Property(e => e.LeadTimeDays).HasDefaultValue(7);
            entity.Property(e => e.ReorderPoint)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.SafetyStock)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.StockQty)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.Unit)
                .HasMaxLength(20)
                .HasDefaultValue("pcs");
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<PaymentRecord>(entity =>
        {
            entity.HasKey(e => e.RecordId);

            entity.HasIndex(e => e.CreatedAt, "IX_PaymentRecords_CreatedAt");

            entity.HasIndex(e => e.TransactionType, "IX_PaymentRecords_TransactionType");

            entity.Property(e => e.Amount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.Category).HasMaxLength(50);
            entity.Property(e => e.CenterId).HasColumnName("CenterID");
            entity.Property(e => e.Fee).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.NetAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.OrderNo).HasMaxLength(50);
            entity.Property(e => e.PayableId).HasColumnName("PayableID");
            entity.Property(e => e.PaymentMethod).HasMaxLength(50);
            entity.Property(e => e.PaymentType)
                .HasMaxLength(30)
                .HasDefaultValue("Receipt");
            entity.Property(e => e.ReceivableId).HasColumnName("ReceivableID");
            entity.Property(e => e.ReconcileStatus).HasMaxLength(20);
            entity.Property(e => e.RecordId)
                .ValueGeneratedOnAdd()
                .HasColumnName("RecordID");
            entity.Property(e => e.Remark).HasMaxLength(200);
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.TransactionNo).HasMaxLength(100);
            entity.Property(e => e.TransactionType).HasMaxLength(20);
            entity.Property(e => e.VoucherNo).HasMaxLength(50);
        });

        modelBuilder.Entity<PointTransaction>(entity =>
        {
            entity.HasKey(e => e.TransactionId);

            entity.Property(e => e.CreatedBy).HasMaxLength(50);
            entity.Property(e => e.Description).HasMaxLength(255);
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.Reason).HasMaxLength(200);
            entity.Property(e => e.TransactionId)
                .ValueGeneratedOnAdd()
                .HasColumnName("TransactionID");
            entity.Property(e => e.TransactionType).HasMaxLength(20);
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<PointsLedger>(entity =>
        {
            entity.HasKey(e => e.LedgerId).HasName("PK__PointsLe__AE70E0AFE96D37B5");

            entity.ToTable("PointsLedger");

            entity.HasIndex(e => e.ExpiresAt, "IX_PointsLedger_ExpiresAt").HasFilter("([ExpiresAt] IS NOT NULL AND [IsExpired]=(0))");

            entity.HasIndex(e => new { e.UserId, e.CreatedAt }, "IX_PointsLedger_UserID_CreatedAt").IsDescending(false, true);

            entity.Property(e => e.LedgerId).HasColumnName("LedgerID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Description)
                .HasMaxLength(300)
                .HasDefaultValue("");
            entity.Property(e => e.ExpiresAt).HasColumnType("datetime");
            entity.Property(e => e.PointType).HasMaxLength(20);
            entity.Property(e => e.ReferenceId).HasColumnName("ReferenceID");
            entity.Property(e => e.Source)
                .HasMaxLength(30)
                .HasDefaultValue("");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<PointsRedemption>(entity =>
        {
            entity.HasKey(e => e.RedemptionId).HasName("PK__PointsRe__410680D168BAF4DA");

            entity.ToTable("PointsRedemption");

            entity.Property(e => e.RedemptionId).HasColumnName("RedemptionID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Description)
                .HasMaxLength(500)
                .HasDefaultValue("");
            entity.Property(e => e.ImageUrl)
                .HasMaxLength(300)
                .HasDefaultValue("")
                .HasColumnName("ImageURL");
            entity.Property(e => e.IsEnabled).HasDefaultValue(true);
            entity.Property(e => e.ItemName).HasMaxLength(100);
            entity.Property(e => e.ItemType).HasMaxLength(30);
            entity.Property(e => e.RedemptionValue).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.Terms)
                .HasMaxLength(500)
                .HasDefaultValue("");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<PointsRule>(entity =>
        {
            entity.HasKey(e => e.RuleId).HasName("PK__PointsRu__110458C2E90F9E0A");

            entity.HasIndex(e => e.RuleCode, "UQ__PointsRu__D618C1EE6E819AA2").IsUnique();

            entity.Property(e => e.RuleId).HasColumnName("RuleID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Description)
                .HasMaxLength(200)
                .HasDefaultValue("");
            entity.Property(e => e.IsEnabled).HasDefaultValue(true);
            entity.Property(e => e.RuleCode).HasMaxLength(50);
            entity.Property(e => e.RuleName).HasMaxLength(100);
            entity.Property(e => e.RuleUnit)
                .HasMaxLength(20)
                .HasDefaultValue("");
            entity.Property(e => e.RuleValue).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<PostComment>(entity =>
        {
            entity.HasKey(e => e.CommentId).HasName("PK__PostComm__C3B4DFAA9B116B66");

            entity.Property(e => e.CommentId).HasColumnName("CommentID");
            entity.Property(e => e.Content).HasMaxLength(1000);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.ParentCommentId).HasColumnName("ParentCommentID");
            entity.Property(e => e.PostId).HasColumnName("PostID");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<PostLike>(entity =>
        {
            entity.HasKey(e => e.LikeId).HasName("PK__PostLike__A2922CF409E0B118");

            entity.HasIndex(e => new { e.PostId, e.UserId }, "UQ_PostLikes").IsUnique();

            entity.Property(e => e.LikeId).HasColumnName("LikeID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.PostId).HasColumnName("PostID");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<PriceChangeLog>(entity =>
        {
            entity.HasKey(e => e.LogId).HasName("PK__PriceCha__5E5499A87EA9023F");

            entity.ToTable("PriceChangeLog");

            entity.Property(e => e.LogId).HasColumnName("LogID");
            entity.Property(e => e.ChangedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ChangedBy).HasMaxLength(100);
            entity.Property(e => e.FieldChanged).HasMaxLength(50);
            entity.Property(e => e.NewValue).HasMaxLength(200);
            entity.Property(e => e.OldValue).HasMaxLength(200);
            entity.Property(e => e.PriceId).HasColumnName("PriceID");
        });

        modelBuilder.Entity<PrintingInventory>(entity =>
        {
            entity.HasKey(e => e.PrintingId).HasName("PK__Printing__E79AFD769D0BD6B0");

            entity.ToTable("PrintingInventory");

            entity.HasIndex(e => new { e.ItemName, e.ItemCode }, "IX_Printing_Name_Code");

            entity.HasIndex(e => new { e.StockQty, e.SafetyStock }, "IX_Printing_StockSafety");

            entity.Property(e => e.PrintingId).HasColumnName("PrintingID");
            entity.Property(e => e.AvgDailyUsage)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 6)");
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(100);
            entity.Property(e => e.LastReplenishDate).HasColumnType("datetime");
            entity.Property(e => e.LeadTimeDays).HasDefaultValue(7);
            entity.Property(e => e.ReorderPoint)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.SafetyStock)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.StockQty)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.Unit)
                .HasMaxLength(20)
                .HasDefaultValue("sheets");
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<Product>(entity =>
        {
            entity.HasKey(e => e.ProductId).HasName("PK_Products");

            entity.HasIndex(e => e.IsActive, "IX_Products_IsActive");

            entity.HasIndex(e => e.ProductType, "IX_Products_ProductType");

            // V19: v18_perf_indexes.sql 复合索引 — 对齐 Classic ASP 版本关键索引策略
            entity.HasIndex(e => new { e.ProductType, e.IsActive }, "IX_Products_ProductType_IsActive");
            entity.HasIndex(e => new { e.Category, e.IsActive }, "IX_Products_Category_IsActive");
            entity.HasIndex(e => new { e.IsActive, e.CreatedAt }, "IX_Products_CreatedAt_Active").IsDescending(false, true);
            entity.HasIndex(e => new { e.ProductType, e.IsActive, e.Kolid }, "IX_Products_KOL_Active")
                .HasFilter("([ProductType]='KOL' AND [IsActive]<>(0))");
            entity.HasIndex(e => e.ProductName, "IX_Products_ProductName");

            entity.Property(e => e.BasePrice).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.Bomcost)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("BOMCost");
            entity.Property(e => e.Category).HasMaxLength(50);
            entity.Property(e => e.EngravingPrice).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ImageUrl)
                .HasMaxLength(200)
                .HasColumnName("ImageURL");
            entity.Property(e => e.Kolid).HasColumnName("KOLID");
            entity.Property(e => e.ProductId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ProductID");
            entity.Property(e => e.ProductName).HasMaxLength(100);
            entity.Property(e => e.ProductType).HasMaxLength(50);
            entity.Property(e => e.RecipeId).HasColumnName("RecipeID");
            entity.Property(e => e.ReviewStatus).HasMaxLength(20);
            entity.Property(e => e.UnitCost).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<ProductBottleStyle>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.BottleId).HasColumnName("BottleID");
            entity.Property(e => e.CustomPrice).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.Id)
                .ValueGeneratedOnAdd()
                .HasColumnName("ID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
        });

        modelBuilder.Entity<ProductCost>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CostId)
                .ValueGeneratedOnAdd()
                .HasColumnName("CostID");
            entity.Property(e => e.CostName).HasMaxLength(100);
            entity.Property(e => e.CostType).HasMaxLength(20);
            entity.Property(e => e.CreatedBy).HasMaxLength(50);
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.TotalCost).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitCost).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<ProductImage>(entity =>
        {
            entity.HasKey(e => e.ImageId);

            entity.Property(e => e.ImageId).HasColumnName("ImageID");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.ImageSize).HasDefaultValue(0);
            entity.Property(e => e.ImageUrl)
                .HasMaxLength(500)
                .HasColumnName("ImageURL");
            entity.Property(e => e.IsPrimary).HasDefaultValue(false);
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.SortOrder).HasDefaultValue(0);
        });

        modelBuilder.Entity<ProductInventory>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("ProductInventory");

            entity.HasIndex(e => e.ProductId, "IX_ProductInventory_ProductID");

            entity.Property(e => e.InventoryId)
                .ValueGeneratedOnAdd()
                .HasColumnName("InventoryID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.StockType).HasMaxLength(20);
            entity.Property(e => e.UnitCost).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<ProductManufacturing>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("ProductManufacturing");

            entity.Property(e => e.BatchNo).HasMaxLength(30);
            entity.Property(e => e.ManufacturingId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ManufacturingID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.ProductName).HasMaxLength(100);
            entity.Property(e => e.ProductRecipeId).HasColumnName("ProductRecipeID");
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.TransferRequestId).HasColumnName("TransferRequestID");
            entity.Property(e => e.WorkCenter).HasMaxLength(20);
        });

        modelBuilder.Entity<ProductManufacturingDetail>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.DetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("DetailID");
            entity.Property(e => e.ManufacturingId).HasColumnName("ManufacturingID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.NoteName).HasMaxLength(100);
            entity.Property(e => e.TotalCost).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.UnitCost).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<ProductNote>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.ProductNoteId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ProductNoteID");
        });

        modelBuilder.Entity<ProductNoteRatio>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.ProductId, "IX_ProductNoteRatios_ProductID");

            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.NoteType).HasMaxLength(20);
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.RatioId)
                .ValueGeneratedOnAdd()
                .HasColumnName("RatioID");
        });

        modelBuilder.Entity<ProductReview>(entity =>
        {
            entity.HasKey(e => e.ReviewId);

            entity.HasIndex(e => new { e.ProductId, e.Status }, "IX_ProductReviews_ProductID");

            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.ReviewId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ReviewID");
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.UserId).HasColumnName("UserID");
            entity.Property(e => e.Title).HasMaxLength(200);
            entity.Property(e => e.AIFeelingSummary).HasMaxLength(2000);

            entity.HasMany(e => e.Images)
                .WithOne(e => e.Review)
                .HasForeignKey(e => e.ReviewId)
                .OnDelete(DeleteBehavior.Cascade);
        });

        modelBuilder.Entity<ReviewImage>(entity =>
        {
            entity.HasKey(e => e.ImageId);

            entity.ToTable("ReviewImages");

            entity.HasIndex(e => e.ReviewId, "IX_ReviewImages_ReviewId");

            entity.Property(e => e.ImageId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ImageID");
            entity.Property(e => e.ReviewId).HasColumnName("ReviewID");
            entity.Property(e => e.ImageUrl).HasMaxLength(500);
        });

        modelBuilder.Entity<ProductTypeConfig>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("ProductTypeConfig");

            entity.Property(e => e.ConfigId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ConfigID");
            entity.Property(e => e.DisplayName).HasMaxLength(50);
            entity.Property(e => e.Icon).HasMaxLength(100);
            entity.Property(e => e.NavName).HasMaxLength(50);
            entity.Property(e => e.TypeCode).HasMaxLength(20);
        });

        modelBuilder.Entity<ProductVolumePrice>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.Price).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.PvpriceId)
                .ValueGeneratedOnAdd()
                .HasColumnName("PVPriceID");
            entity.Property(e => e.VolumeId).HasColumnName("VolumeID");
        });

        modelBuilder.Entity<ProductionLog>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CreatedBy).HasMaxLength(100);
            entity.Property(e => e.LogId)
                .ValueGeneratedOnAdd()
                .HasColumnName("LogID");
            entity.Property(e => e.ProductionId).HasColumnName("ProductionID");
            entity.Property(e => e.Status).HasMaxLength(20);
        });

        modelBuilder.Entity<ProductionOrder>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.Status, "IX_ProductionOrders_Status");

            entity.Property(e => e.AssignedTo).HasMaxLength(100);
            entity.Property(e => e.BatchNo).HasMaxLength(50);
            entity.Property(e => e.DetailId).HasColumnName("DetailID");
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.PlannedQty).HasDefaultValue(0);
            entity.Property(e => e.PriorityText).HasMaxLength(10);
            entity.Property(e => e.ProductionId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ProductionID");
            entity.Property(e => e.Qcnotes).HasColumnName("QCNotes");
            entity.Property(e => e.QcpassedAt).HasColumnName("QCPassedAt");
            entity.Property(e => e.RecipeId).HasColumnName("RecipeID");
            entity.Property(e => e.RecipeName).HasMaxLength(100);
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.WorkOrderNo).HasMaxLength(50);
        });

        modelBuilder.Entity<PurchaseBatch>(entity =>
        {
            entity.HasKey(e => e.BatchId).HasName("PK__Purchase__5D55CE388178606B");

            entity.Property(e => e.BatchId).HasColumnName("BatchID");
            entity.Property(e => e.BatchNo).HasMaxLength(50);
            entity.Property(e => e.CostAllocated).HasDefaultValue(false);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.ItemType)
                .HasMaxLength(30)
                .HasDefaultValue("RawMaterial");
            entity.Property(e => e.PurchaseDetailId).HasColumnName("PurchaseDetailID");
            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.Quantity).HasDefaultValue(0.0);
            entity.Property(e => e.ReceivedDate).HasColumnType("datetime");
            entity.Property(e => e.ReceivedQty).HasDefaultValue(0.0);
            entity.Property(e => e.RemainingQty).HasDefaultValue(0.0);
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<PurchaseCategory>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CategoryCode).HasMaxLength(20);
            entity.Property(e => e.CategoryId)
                .ValueGeneratedOnAdd()
                .HasColumnName("CategoryID");
            entity.Property(e => e.CategoryName).HasMaxLength(100);
        });

        modelBuilder.Entity<PurchaseCostReview>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("PurchaseCostReview");

            entity.Property(e => e.CostAllocation).HasMaxLength(20);
            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.ReviewAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ReviewId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ReviewID");
            entity.Property(e => e.ReviewStatus).HasMaxLength(20);
            entity.Property(e => e.ReviewerId).HasColumnName("ReviewerID");
        });

        modelBuilder.Entity<PurchaseHistoryStat>(entity =>
        {
            entity.HasKey(e => e.StatId).HasName("PK__Purchase__3A162D1E6E535A88");

            entity.Property(e => e.StatId).HasColumnName("StatID");
            entity.Property(e => e.Avg30DayUsage)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 6)");
            entity.Property(e => e.Avg90DayUsage)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 6)");
            entity.Property(e => e.ItemCode).HasMaxLength(100);
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.ItemType).HasMaxLength(30);
            entity.Property(e => e.LastOrderDate).HasColumnType("datetime");
            entity.Property(e => e.PreferredSupplierId).HasColumnName("PreferredSupplierID");
            entity.Property(e => e.PreferredUnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.TotalOrders90Days).HasDefaultValue(0);
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<PurchaseOrder>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.Status, "IX_PurchaseOrders_Status");

            entity.HasIndex(e => new { e.OrderType, e.ExpectedDate }, "IX_PurchaseOrders_Type_Date");

            entity.Property(e => e.CategoryCode).HasMaxLength(20);
            entity.Property(e => e.OrderType)
                .HasMaxLength(20)
                .HasDefaultValue("RawMaterial");
            entity.Property(e => e.PurchaseId)
                .ValueGeneratedOnAdd()
                .HasColumnName("PurchaseID");
            entity.Property(e => e.PurchaseNo).HasMaxLength(50);
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.TotalAmount).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<PurchaseOrderDetail>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.ItemCode, "IX_PODetails_ItemCode");

            entity.Property(e => e.DetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("DetailID");
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.Specification).HasMaxLength(200);
            entity.Property(e => e.TotalPrice).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.Unit).HasMaxLength(20);
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<PurchaseOrderStatusLog>(entity =>
        {
            entity.HasKey(e => e.LogId).HasName("PK__Purchase__5E5499A8BF0510F2");

            entity.ToTable("PurchaseOrderStatusLog");

            entity.Property(e => e.LogId).HasColumnName("LogID");
            entity.Property(e => e.ChangedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ChangedBy).HasMaxLength(50);
            entity.Property(e => e.FromStatus).HasMaxLength(30);
            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.Remarks).HasMaxLength(200);
            entity.Property(e => e.ToStatus).HasMaxLength(30);
        });

        modelBuilder.Entity<PurchaseReceipt>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.PurchaseId).HasColumnName("PurchaseID");
            entity.Property(e => e.ReceiptId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ReceiptID");
            entity.Property(e => e.ReceiptNo).HasMaxLength(50);
            entity.Property(e => e.ReceivedBy).HasMaxLength(50);
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
        });

        modelBuilder.Entity<PurchaseReceiptDetail>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.MaterialId).HasColumnName("MaterialID");
            entity.Property(e => e.PurchaseDetailId).HasColumnName("PurchaseDetailID");
            entity.Property(e => e.ReceiptDetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ReceiptDetailID");
            entity.Property(e => e.ReceiptId).HasColumnName("ReceiptID");
            entity.Property(e => e.RejectReason).HasMaxLength(200);
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<RawMaterialInventory>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("RawMaterialInventory");

            entity.HasIndex(e => e.ItemCode, "IX_RawMaterialInventory_ItemCode");

            entity.HasIndex(e => new { e.ItemName, e.ItemCode }, "IX_RawMaterial_Name_Code");

            entity.HasIndex(e => new { e.StockQty, e.SafetyStock }, "IX_RawMaterial_StockSafety");

            entity.Property(e => e.AvgDailyUsage)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 6)");
            entity.Property(e => e.CategoryCode).HasMaxLength(20);
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.LastReplenishDate).HasColumnType("datetime");
            entity.Property(e => e.LeadTimeDays).HasDefaultValue(7);
            entity.Property(e => e.MaterialId)
                .ValueGeneratedOnAdd()
                .HasColumnName("MaterialID");
            entity.Property(e => e.ReorderPoint)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.Unit).HasMaxLength(20);
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.WeightedUnitCost)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 6)");
        });

        modelBuilder.Entity<Recipe>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.CreatedBy).HasMaxLength(100);
            entity.Property(e => e.ProductType).HasMaxLength(20);
            entity.Property(e => e.RecipeCode).HasMaxLength(50);
            entity.Property(e => e.RecipeId)
                .ValueGeneratedOnAdd()
                .HasColumnName("RecipeID");
            entity.Property(e => e.RecipeName).HasMaxLength(100);
            entity.Property(e => e.ReviewStatus).HasMaxLength(20);
        });

        modelBuilder.Entity<RecipeAccord>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.NoteId, "IX_RecipeAccords_NoteID");

            entity.Property(e => e.AccordRecipeId)
                .ValueGeneratedOnAdd()
                .HasColumnName("AccordRecipeID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.PublishedBy).HasMaxLength(50);
            entity.Property(e => e.RecipeId).HasColumnName("RecipeID");
            entity.Property(e => e.RecipeName).HasMaxLength(100);
            entity.Property(e => e.Status).HasMaxLength(20);
        });

        modelBuilder.Entity<RecipeAccordMaterial>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.AccordRecipeId, "IX_RecipeAccordMaterials_AccordRecipeID");

            entity.Property(e => e.AccordRecipeId).HasColumnName("AccordRecipeID");
            entity.Property(e => e.DetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("DetailID");
            entity.Property(e => e.MaterialId).HasColumnName("MaterialID");
            entity.Property(e => e.MaterialName).HasMaxLength(100);
        });

        modelBuilder.Entity<RecipeIngredient>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.Id)
                .ValueGeneratedOnAdd()
                .HasColumnName("ID");
            entity.Property(e => e.IngredientName).HasMaxLength(100);
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.RecipeId).HasColumnName("RecipeID");
        });

        modelBuilder.Entity<RecipeNote>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.Id)
                .ValueGeneratedOnAdd()
                .HasColumnName("ID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.NoteType).HasMaxLength(20);
            entity.Property(e => e.RecipeId).HasColumnName("RecipeID");
        });

        modelBuilder.Entity<RecipePopularity>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("RecipePopularity");

            entity.Property(e => e.PopularityId)
                .ValueGeneratedOnAdd()
                .HasColumnName("PopularityID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
        });

        modelBuilder.Entity<RecipeProduct>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.ProductRecipeId)
                .ValueGeneratedOnAdd()
                .HasColumnName("ProductRecipeID");
            entity.Property(e => e.PublishedBy).HasMaxLength(50);
            entity.Property(e => e.RecipeId).HasColumnName("RecipeID");
            entity.Property(e => e.Status).HasMaxLength(20);
        });

        modelBuilder.Entity<RecipeProductNote>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.DetailId)
                .ValueGeneratedOnAdd()
                .HasColumnName("DetailID");
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.NoteName).HasMaxLength(100);
            entity.Property(e => e.ProductRecipeId).HasColumnName("ProductRecipeID");
        });

        modelBuilder.Entity<RecipePublishLog>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("RecipePublishLog");

            entity.Property(e => e.Ipaddress)
                .HasMaxLength(50)
                .HasColumnName("IPAddress");
            entity.Property(e => e.LogId)
                .ValueGeneratedOnAdd()
                .HasColumnName("LogID");
            entity.Property(e => e.PublishType).HasMaxLength(20);
            entity.Property(e => e.PublishedBy).HasMaxLength(50);
            entity.Property(e => e.RecipeId).HasColumnName("RecipeID");
            entity.Property(e => e.TargetRecipeId).HasColumnName("TargetRecipeID");
        });

        modelBuilder.Entity<RecommendedRecipe>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.RecipeId)
                .ValueGeneratedOnAdd()
                .HasColumnName("RecipeID");
            entity.Property(e => e.RecipeName).HasMaxLength(200);
        });

        modelBuilder.Entity<ReconciliationLog>(entity =>
        {
            entity.HasKey(e => e.LogId);

            entity.Property(e => e.Difference).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.LogId)
                .ValueGeneratedOnAdd()
                .HasColumnName("LogID");
            entity.Property(e => e.OrderAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.OrderNo).HasMaxLength(50);
            entity.Property(e => e.PaymentAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.ResolvedBy).HasMaxLength(50);
            entity.Property(e => e.Status).HasMaxLength(20);
        });

        modelBuilder.Entity<ReferralRelation>(entity =>
        {
            entity.HasKey(e => e.RelationId).HasName("PK__Referral__E2DA1695A250870E");

            entity.HasIndex(e => e.AncestorUserId, "IX_ReferralRelations_Ancestor");

            entity.HasIndex(e => e.DescendantUserId, "IX_ReferralRelations_Descendant");

            entity.Property(e => e.RelationId).HasColumnName("RelationID");
            entity.Property(e => e.AncestorUserId).HasColumnName("AncestorUserID");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.DescendantUserId).HasColumnName("DescendantUserID");
        });

        modelBuilder.Entity<ReferralToken>(entity =>
        {
            entity.HasKey(e => e.TokenId).HasName("PK__Referral__658FEE8A29895946");

            entity.HasIndex(e => e.ReferrerUserId, "IX_ReferralTokens_ReferrerUserID");

            entity.HasIndex(e => e.TokenHash, "IX_ReferralTokens_TokenHash");

            entity.Property(e => e.TokenId).HasColumnName("TokenID");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.MaxUses).HasDefaultValue(1);
            entity.Property(e => e.OriginalToken).HasMaxLength(1000);
            entity.Property(e => e.ReferrerType)
                .HasMaxLength(20)
                .HasDefaultValue("user");
            entity.Property(e => e.ReferrerUserId).HasColumnName("ReferrerUserID");
            entity.Property(e => e.TokenHash).HasMaxLength(255);
            entity.Property(e => e.UsedCount).HasDefaultValue(0);
        });

        modelBuilder.Entity<RefundRecord>(entity =>
        {
            entity.HasKey(e => e.RefundId);

            entity.Property(e => e.ApprovedBy).HasMaxLength(50);
            entity.Property(e => e.OrderId).HasColumnName("OrderID");
            entity.Property(e => e.OrderNo).HasMaxLength(50);
            entity.Property(e => e.RefundAmount).HasColumnType("decimal(19, 4)");
            entity.Property(e => e.RefundId)
                .ValueGeneratedOnAdd()
                .HasColumnName("RefundID");
            entity.Property(e => e.RefundNo).HasMaxLength(50);
            entity.Property(e => e.Status).HasMaxLength(20);
        });

        modelBuilder.Entity<RegistrationAttempt>(entity =>
        {
            entity.HasKey(e => e.AttemptId).HasName("PK__Registra__891A6886C72898EC");

            entity.HasIndex(e => new { e.DeviceFingerprint, e.AttemptedAt }, "IX_RegistrationAttempts_Fingerprint");

            entity.HasIndex(e => new { e.Ipaddress, e.AttemptedAt }, "IX_RegistrationAttempts_IP");

            entity.Property(e => e.AttemptId).HasColumnName("AttemptID");
            entity.Property(e => e.AttemptedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.DeviceFingerprint).HasMaxLength(100);
            entity.Property(e => e.Ipaddress)
                .HasMaxLength(50)
                .HasColumnName("IPAddress");
            entity.Property(e => e.Success).HasDefaultValue(false);
            entity.Property(e => e.TokenHash).HasMaxLength(255);
        });

        modelBuilder.Entity<ReviewLike>(entity =>
        {
            entity.HasKey(e => e.LikeId).HasName("PK__ReviewLi__A2922CF4CA22E7E6");

            entity.HasIndex(e => new { e.ReviewId, e.UserId }, "UQ_ReviewLikes").IsUnique();

            entity.Property(e => e.LikeId).HasColumnName("LikeID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.ReviewId).HasColumnName("ReviewID");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<RolePermission>(entity =>
        {
            entity.HasKey(e => e.PermId).HasName("PK__RolePerm__29C30A7CA6902708");

            entity.Property(e => e.PermId).HasColumnName("PermID");
            entity.Property(e => e.CanApprove).HasDefaultValue(false);
            entity.Property(e => e.CanCreate).HasDefaultValue(false);
            entity.Property(e => e.CanDelete).HasDefaultValue(false);
            entity.Property(e => e.CanEdit).HasDefaultValue(false);
            entity.Property(e => e.CanExport).HasDefaultValue(false);
            entity.Property(e => e.CanView).HasDefaultValue(false);
            entity.Property(e => e.ModuleCode).HasMaxLength(50);
            entity.Property(e => e.RoleId).HasColumnName("RoleID");
        });

        modelBuilder.Entity<ShippingCompany>(entity =>
        {
            entity.HasKey(e => e.CompanyId).HasName("PK__Shipping__2D971C4C8B2429E8");

            entity.Property(e => e.CompanyId).HasColumnName("CompanyID");
            entity.Property(e => e.CompanyName).HasMaxLength(100);
            entity.Property(e => e.ContactPerson).HasMaxLength(50);
            entity.Property(e => e.ContactPhone).HasMaxLength(20);
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.UpdatedAt).HasColumnType("datetime");
            entity.Property(e => e.Website).HasMaxLength(200);
        });

        modelBuilder.Entity<SiteSetting>(entity =>
        {
            entity.HasKey(e => e.SettingKey);

            entity.Property(e => e.Description).HasMaxLength(255);
            entity.Property(e => e.SecurityLockoutMinutes)
                .HasDefaultValue(30)
                .HasColumnName("Security_LockoutMinutes");
            entity.Property(e => e.SecurityLoginMaxAttempts)
                .HasDefaultValue(5)
                .HasColumnName("Security_LoginMaxAttempts");
            entity.Property(e => e.SecurityMfaenabled)
                .HasDefaultValue(false)
                .HasColumnName("Security_MFAEnabled");
            entity.Property(e => e.SecurityPasswordMinLength)
                .HasDefaultValue(8)
                .HasColumnName("Security_PasswordMinLength");
            entity.Property(e => e.SecuritySessionTimeout)
                .HasDefaultValue(30)
                .HasColumnName("Security_SessionTimeout");
            entity.Property(e => e.SettingKey).HasMaxLength(50);
            entity.Property(e => e.SettingName).HasMaxLength(100);
            entity.Property(e => e.SettingValue).HasMaxLength(255);
        });

        modelBuilder.Entity<SprayHeadInventory>(entity =>
        {
            entity.HasKey(e => e.SprayHeadId).HasName("PK__SprayHea__F30C8A265EDC2E4B");

            entity.ToTable("SprayHeadInventory");

            entity.HasIndex(e => new { e.ItemName, e.ItemCode }, "IX_SprayHead_Name_Code");

            entity.HasIndex(e => new { e.StockQty, e.SafetyStock }, "IX_SprayHead_StockSafety");

            entity.Property(e => e.SprayHeadId).HasColumnName("SprayHeadID");
            entity.Property(e => e.AvgDailyUsage)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 6)");
            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(100);
            entity.Property(e => e.LastReplenishDate).HasColumnType("datetime");
            entity.Property(e => e.LeadTimeDays).HasDefaultValue(7);
            entity.Property(e => e.ReorderPoint)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(19, 4)");
            entity.Property(e => e.SafetyStock)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.StockQty)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.Unit)
                .HasMaxLength(20)
                .HasDefaultValue("pcs");
            entity.Property(e => e.UnitPrice)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UpdatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
        });

        modelBuilder.Entity<StockMovement>(entity =>
        {
            entity.HasKey(e => e.MovementId).HasName("PK__StockMov__D182246679AA11F7");

            entity.Property(e => e.MovementId).HasColumnName("MovementID");
            entity.Property(e => e.AfterQty).HasColumnType("decimal(12, 2)");
            entity.Property(e => e.BeforeQty).HasColumnType("decimal(12, 2)");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.CreatedBy).HasMaxLength(50);
            entity.Property(e => e.ItemCode).HasMaxLength(100);
            entity.Property(e => e.ItemId).HasColumnName("ItemID");
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.ItemType).HasMaxLength(30);
            entity.Property(e => e.MovementType).HasMaxLength(20);
            entity.Property(e => e.Notes).HasMaxLength(500);
            entity.Property(e => e.Quantity).HasColumnType("decimal(12, 2)");
            entity.Property(e => e.ReferenceNo).HasMaxLength(100);
            entity.Property(e => e.Unit).HasMaxLength(20);
        });

        modelBuilder.Entity<SubscriptionDelivery>(entity =>
        {
            entity.HasKey(e => e.DeliveryId).HasName("PK__Subscrip__626D8FEEDE75EF41");

            entity.Property(e => e.DeliveryId).HasColumnName("DeliveryID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.DeliveryDate).HasColumnType("datetime");
            entity.Property(e => e.Status)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasDefaultValue("pending");
            entity.Property(e => e.SubscriptionId).HasColumnName("SubscriptionID");
            entity.Property(e => e.TrackingNo)
                .HasMaxLength(100)
                .IsUnicode(false);
        });

        modelBuilder.Entity<SubscriptionPlan>(entity =>
        {
            entity.HasKey(e => e.PlanId).HasName("PK__Subscrip__755C22D79DCFF259");

            entity.Property(e => e.PlanId).HasColumnName("PlanID");
            entity.Property(e => e.CancellationFee).HasColumnType("decimal(12, 2)");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Description).HasMaxLength(500);
            entity.Property(e => e.FeaturedImage)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.FreeShipping).HasDefaultValue(true);
            entity.Property(e => e.FullSizeCount).HasDefaultValue(1);
            entity.Property(e => e.IsActive).HasDefaultValue(true);
            entity.Property(e => e.Period)
                .HasMaxLength(20)
                .IsUnicode(false);
            entity.Property(e => e.PlanName).HasMaxLength(100);
            entity.Property(e => e.Price).HasColumnType("decimal(12, 2)");
            entity.Property(e => e.SampleCount).HasDefaultValue(3);
        });

        modelBuilder.Entity<Supplier>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.Address).HasMaxLength(255);
            entity.Property(e => e.Category).HasMaxLength(50);
            entity.Property(e => e.ContactPerson).HasMaxLength(50);
            entity.Property(e => e.Email).HasMaxLength(100);
            entity.Property(e => e.Phone).HasMaxLength(30);
            entity.Property(e => e.SupplierId)
                .ValueGeneratedOnAdd()
                .HasColumnName("SupplierID");
            entity.Property(e => e.SupplierName).HasMaxLength(100);
        });

        modelBuilder.Entity<SupplierContract>(entity =>
        {
            entity.HasKey(e => e.ContractId).HasName("PK__Supplier__C90D34096DA5D0E5");

            entity.Property(e => e.ContractId).HasColumnName("ContractID");
            entity.Property(e => e.AttachmentUrl)
                .HasMaxLength(500)
                .HasColumnName("AttachmentURL");
            entity.Property(e => e.ContractName).HasMaxLength(200);
            entity.Property(e => e.ContractNo).HasMaxLength(50);
            entity.Property(e => e.ContractType)
                .HasMaxLength(30)
                .HasDefaultValue("Supply");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.PaymentTerms).HasMaxLength(200);
            entity.Property(e => e.Status)
                .HasMaxLength(20)
                .HasDefaultValue("Active");
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.TotalAmount).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<SupplierEvaluation>(entity =>
        {
            entity.HasKey(e => e.EvaluationId).HasName("PK__Supplier__36AE68D3637426B2");

            entity.Property(e => e.EvaluationId).HasColumnName("EvaluationID");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.DeliveryScore).HasDefaultValue(0);
            entity.Property(e => e.EvaluatedBy).HasMaxLength(50);
            entity.Property(e => e.EvaluationDate).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.OverallScore).HasDefaultValue(0);
            entity.Property(e => e.Period).HasMaxLength(20);
            entity.Property(e => e.PriceScore).HasDefaultValue(0);
            entity.Property(e => e.QualityScore).HasDefaultValue(0);
            entity.Property(e => e.Rating)
                .HasMaxLength(10)
                .HasDefaultValue("C");
            entity.Property(e => e.ServiceScore).HasDefaultValue(0);
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
        });

        modelBuilder.Entity<SupplierPrice>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => new { e.ItemCode, e.IsActive }, "IX_SupplierPrices_ItemCode");

            entity.HasIndex(e => new { e.ItemCode, e.IsActive }, "IX_SupplierPrices_ItemCode_Active");

            entity.Property(e => e.ItemCode).HasMaxLength(50);
            entity.Property(e => e.ItemName).HasMaxLength(200);
            entity.Property(e => e.PriceId)
                .ValueGeneratedOnAdd()
                .HasColumnName("PriceID");
            entity.Property(e => e.PriceType)
                .HasMaxLength(30)
                .HasDefaultValue("Standard");
            entity.Property(e => e.SupplierId).HasColumnName("SupplierID");
            entity.Property(e => e.Unit)
                .HasMaxLength(20)
                .HasDefaultValue("kg");
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(19, 4)");
        });

        modelBuilder.Entity<TierConfigLog>(entity =>
        {
            entity.HasKey(e => e.LogId).HasName("PK__TierConf__5E5499A8B77791D0");

            entity.ToTable("TierConfigLog");

            entity.Property(e => e.LogId).HasColumnName("LogID");
            entity.Property(e => e.ChangedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.ChangedBy)
                .HasMaxLength(50)
                .HasDefaultValue("admin");
            entity.Property(e => e.FieldName).HasMaxLength(50);
            entity.Property(e => e.NewValue).HasMaxLength(200);
            entity.Property(e => e.OldValue).HasMaxLength(200);
            entity.Property(e => e.TierCode)
                .HasMaxLength(20)
                .IsUnicode(false);
        });

        modelBuilder.Entity<User>(entity =>
        {
            entity.HasKey(e => e.UserId).HasName("PK_Users");

            entity.HasIndex(e => e.CreatedAt, "IX_Users_CreatedAt");

            entity.HasIndex(e => e.Email, "IX_Users_Email");

            entity.HasIndex(e => e.Username, "IX_Users_Username");

            entity.Property(e => e.UserId).HasColumnName("UserID");
            entity.Property(e => e.Address).HasMaxLength(200);
            entity.Property(e => e.City).HasMaxLength(50);
            entity.Property(e => e.CustomerTier)
                .HasMaxLength(20)
                .HasDefaultValue("bronze");
            entity.Property(e => e.DeviceFingerprint).HasMaxLength(100);
            entity.Property(e => e.Email).HasMaxLength(100);
            entity.Property(e => e.FavoriteCategory).HasMaxLength(100);
            entity.Property(e => e.FullName).HasMaxLength(100);
            entity.Property(e => e.IsVip).HasColumnName("IsVIP");
            entity.Property(e => e.LastOrderDate).HasColumnType("datetime");
            entity.Property(e => e.OrderCount).HasDefaultValue(0);
            entity.Property(e => e.Password).HasMaxLength(255);
            entity.Property(e => e.Phone).HasMaxLength(20);
            entity.Property(e => e.PostalCode).HasMaxLength(20);
            entity.Property(e => e.PreferredNote).HasMaxLength(50);
            entity.Property(e => e.ReferrerUserId).HasColumnName("ReferrerUserID");
            entity.Property(e => e.TotalSpent)
                .HasDefaultValue(0m)
                .HasColumnType("decimal(18, 2)");
            entity.Property(e => e.UserRole).HasMaxLength(20);
            entity.Property(e => e.Username).HasMaxLength(50);
        });

        modelBuilder.Entity<UserAddress>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.Address).HasMaxLength(200);
            entity.Property(e => e.AddressId)
                .ValueGeneratedOnAdd()
                .HasColumnName("AddressID");
            entity.Property(e => e.City).HasMaxLength(50);
            entity.Property(e => e.Consignee).HasMaxLength(50);
            entity.Property(e => e.District).HasMaxLength(50);
            entity.Property(e => e.Phone).HasMaxLength(20);
            entity.Property(e => e.Province).HasMaxLength(50);
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<UserCoupon>(entity =>
        {
            entity.HasKey(e => e.UserCouponId).HasName("PK__UserCoup__22994B73C5E78810");

            entity.HasIndex(e => e.CouponCode, "IX_UserCoupons_Code");

            entity.HasIndex(e => new { e.UserId, e.Status }, "IX_UserCoupons_UserID_Status");

            entity.Property(e => e.UserCouponId).HasColumnName("UserCouponID");
            entity.Property(e => e.CouponCode).HasMaxLength(50);
            entity.Property(e => e.CouponId).HasColumnName("CouponID");
            entity.Property(e => e.ExpiresAt).HasColumnType("datetime");
            entity.Property(e => e.ObtainedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Source)
                .HasMaxLength(30)
                .HasDefaultValue("manual");
            entity.Property(e => e.Status)
                .HasMaxLength(10)
                .HasDefaultValue("available");
            entity.Property(e => e.UsedAt).HasColumnType("datetime");
            entity.Property(e => e.UsedOrderId).HasColumnName("UsedOrderID");
            entity.Property(e => e.UserId).HasColumnName("UserID");

            entity.HasOne(d => d.Coupon).WithMany(p => p.UserCoupons)
                .HasForeignKey(d => d.CouponId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_UserCoupons_CouponID");

            entity.HasOne(d => d.User).WithMany(p => p.UserCoupons)
                .HasForeignKey(d => d.UserId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_UserCoupons_UserID");
        });

        modelBuilder.Entity<UserFavorite>(entity =>
        {
            entity.HasNoKey();

            entity.HasIndex(e => e.ProductId, "IX_UserFavorites_ProductID");

            entity.HasIndex(e => e.UserId, "IX_UserFavorites_UserID");

            entity.Property(e => e.FavoriteId)
                .ValueGeneratedOnAdd()
                .HasColumnName("FavoriteID");
            entity.Property(e => e.ProductId).HasColumnName("ProductID");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<UserPoint>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.PointId)
                .ValueGeneratedOnAdd()
                .HasColumnName("PointID");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<UserPreference>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.PreferenceId)
                .ValueGeneratedOnAdd()
                .HasColumnName("PreferenceID");
            entity.Property(e => e.PreferredBaseNotes).HasMaxLength(255);
            entity.Property(e => e.PreferredCategories).HasMaxLength(255);
            entity.Property(e => e.PreferredMiddleNotes).HasMaxLength(255);
            entity.Property(e => e.PreferredTopNotes).HasMaxLength(255);
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<UserSubscription>(entity =>
        {
            entity.HasKey(e => e.SubscriptionId).HasName("PK__UserSubs__9A2B24BDAB0AF441");

            entity.Property(e => e.SubscriptionId).HasColumnName("SubscriptionID");
            entity.Property(e => e.CreatedAt)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.EndDate).HasColumnType("datetime");
            entity.Property(e => e.PlanId).HasColumnName("PlanID");
            entity.Property(e => e.StartDate)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Status)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasDefaultValue("active");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<UserTierHistory>(entity =>
        {
            entity.HasKey(e => e.HistoryId).HasName("PK__UserTier__4D7B4ADDAA0F788A");

            entity.ToTable("UserTierHistory");

            entity.Property(e => e.HistoryId).HasColumnName("HistoryID");
            entity.Property(e => e.ChangeType)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasDefaultValue("auto");
            entity.Property(e => e.ChangedAt).HasDefaultValueSql("(getdate())");
            entity.Property(e => e.NewTierCode)
                .HasMaxLength(20)
                .IsUnicode(false);
            entity.Property(e => e.OldTierCode)
                .HasMaxLength(20)
                .IsUnicode(false);
            entity.Property(e => e.TotalSpent).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UserId).HasColumnName("UserID");
        });

        modelBuilder.Entity<Volume>(entity =>
        {
            entity.HasNoKey();

            entity.Property(e => e.VolumeId)
                .ValueGeneratedOnAdd()
                .HasColumnName("VolumeID");
            entity.Property(e => e.VolumeMl).HasColumnName("VolumeML");
            entity.Property(e => e.VolumeName).HasMaxLength(50);
        });

        modelBuilder.Entity<WorkshopTransfer>(entity =>
        {
            entity
                .HasNoKey()
                .ToTable("WorkshopTransfer");

            entity.Property(e => e.FromWorkshop).HasMaxLength(20);
            entity.Property(e => e.NoteId).HasColumnName("NoteID");
            entity.Property(e => e.RequestedBy).HasMaxLength(50);
            entity.Property(e => e.Status).HasMaxLength(20);
            entity.Property(e => e.ToWorkshop).HasMaxLength(20);
            entity.Property(e => e.TransferId)
                .ValueGeneratedOnAdd()
                .HasColumnName("TransferID");
            entity.Property(e => e.TransferNo).HasMaxLength(30);
        });

        modelBuilder.Entity<PasswordResetToken>(entity =>
        {
            entity.HasKey(e => e.TokenId);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
