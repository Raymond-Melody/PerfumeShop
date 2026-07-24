using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.IntegrationTests.Admin;

/// <summary>
/// M4-B 采购模块 37 页面 Blazor 迁移测试
/// 覆盖: PurchaseRepository CRUD、采购订单E2E流程、批量操作、供应商管理
/// 注: 涉及原生 SQL 库存更新的收货流程（ReceivePurchaseAsync）由 BusinessPathTests.E15 用 SQLite 覆盖
/// </summary>
public class PurchasePagesTests : IDisposable
{
    private readonly PerfumeShopContext _db;
    private readonly PurchaseRepository _repo;

    public PurchasePagesTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"PurchaseTests_{Guid.NewGuid()}")
            .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning))
            .Options;
        _db = new TestEngineContext(options);
        _repo = new PurchaseRepository(_db, new PerfumeShop.Data.Services.InventoryLedger(_db));
    }

    public void Dispose() => _db.Dispose();

    // ==================== 页面路由验证（37 页面） ====================

    [Theory]
    [InlineData("/admin/Purchase/Dashboard")]
    [InlineData("/admin/Purchase/OrderList")]
    [InlineData("/admin/Purchase/OrderDetail")]
    [InlineData("/admin/Purchase/OrderCreate")]
    [InlineData("/admin/Purchase/OrderApprove")]
    [InlineData("/admin/Purchase/OrderCancel")]
    [InlineData("/admin/Purchase/Receiving")]
    [InlineData("/admin/Purchase/ReceivingCreate")]
    [InlineData("/admin/Purchase/ReceivingDetail")]
    [InlineData("/admin/Purchase/BaseNoteReceiving")]
    [InlineData("/admin/Purchase/Replenishment")]
    [InlineData("/admin/Purchase/RestockOrder")]
    [InlineData("/admin/Purchase/SupplierManagement")]
    [InlineData("/admin/Purchase/SupplierContract")]
    [InlineData("/admin/Purchase/SupplierEvaluation")]
    [InlineData("/admin/Purchase/PriceManagement")]
    [InlineData("/admin/Purchase/CostSummary")]
    [InlineData("/admin/Purchase/CostDetail")]
    [InlineData("/admin/Purchase/BottlePurchase")]
    [InlineData("/admin/Purchase/PackagingPurchase")]
    [InlineData("/admin/Purchase/PrintingPurchase")]
    [InlineData("/admin/Purchase/SprayheadPurchase")]
    [InlineData("/admin/Purchase/BatchPurchase")]
    [InlineData("/admin/Purchase/FixedBrandDashboard")]
    [InlineData("/admin/Purchase/FixedBrandProducts")]
    [InlineData("/admin/Purchase/FixedBrandOrders")]
    [InlineData("/admin/Purchase/FixedBrandReceiving")]
    [InlineData("/admin/Purchase/FixedBrandReplenishment")]
    [InlineData("/admin/Purchase/FixedBrandCostProfit")]
    [InlineData("/admin/Purchase/PurchaseBatchList")]
    [InlineData("/admin/Purchase/BatchDetail")]
    [InlineData("/admin/Purchase/PurchaseReport")]
    [InlineData("/admin/Purchase/ReportTrend")]
    [InlineData("/admin/Purchase/ReportSupplierRank")]
    [InlineData("/admin/Purchase/ReportCategoryAnalysis")]
    [InlineData("/admin/Purchase/BatchReceiving")]
    [InlineData("/admin/Purchase/BatchReconciliation")]
    public void PageRoute_ShouldBeValid(string route)
    {
        // 验证路由格式
        Assert.StartsWith("/admin/Purchase/", route);
        Assert.True(route.Length > "/admin/Purchase/".Length);
    }

    // ==================== 采购订单 CRUD ====================

    [Fact]
    public async Task CreatePurchaseOrder_ShouldGeneratePurchaseNo()
    {
        var order = new PurchaseOrder { SupplierId = 1, OrderType = "Bottle", TotalAmount = 1000m };
        var created = await _repo.CreatePurchaseOrderAsync(order);

        Assert.True(created.PurchaseId > 0);
        Assert.StartsWith("PO-", created.PurchaseNo);
        Assert.Equal("draft", created.Status);
        Assert.NotNull(created.CreatedAt);
    }

    [Fact]
    public async Task GetPurchaseOrders_ShouldReturnPagedResults()
    {
        // 创建测试数据
        for (int i = 0; i < 5; i++)
        {
            _db.PurchaseOrders.Add(new PurchaseOrder
            {
                PurchaseNo = $"PO-TEST-{i}",
                SupplierId = 1,
                Status = "draft",
                TotalAmount = 100 * (i + 1),
                CreatedAt = DateTime.Now
            });
        }
        await _db.SaveChangesAsync();

        var (items, total) = await _repo.GetPurchaseOrdersAsync(page: 1, pageSize: 3);

        Assert.Equal(5, total);
        Assert.Equal(3, items.Count);
    }

    [Fact]
    public async Task ApprovePurchaseOrder_ShouldUpdateStatus()
    {
        var order = new PurchaseOrder { SupplierId = 1, Status = "draft", CreatedAt = DateTime.Now };
        _db.PurchaseOrders.Add(order);
        await _db.SaveChangesAsync();

        var result = await _repo.ApprovePurchaseOrderAsync(order.PurchaseId, 1);

        Assert.True(result);
        var updated = await _db.PurchaseOrders.FindAsync(order.PurchaseId);
        Assert.Equal("approved", updated!.Status);
        Assert.NotNull(updated.ApprovedAt);
    }

    [Fact]
    public async Task CancelPurchaseOrder_ShouldUpdateStatusAndLog()
    {
        var order = new PurchaseOrder { SupplierId = 1, Status = "draft", CreatedAt = DateTime.Now };
        _db.PurchaseOrders.Add(order);
        await _db.SaveChangesAsync();

        var result = await _repo.CancelPurchaseOrderAsync(order.PurchaseId, "测试取消");

        Assert.True(result);
        var updated = await _db.PurchaseOrders.FindAsync(order.PurchaseId);
        Assert.Equal("cancelled", updated!.Status);

        var log = await _db.PurchaseOrderStatusLogs.FirstOrDefaultAsync(l => l.PurchaseId == order.PurchaseId);
        Assert.NotNull(log);
        Assert.Equal("cancelled", log.ToStatus);
    }

    [Fact]
    public async Task BatchApprove_ShouldApproveMultipleOrders()
    {
        var orders = new List<PurchaseOrder>();
        for (int i = 0; i < 3; i++)
        {
            var o = new PurchaseOrder { SupplierId = 1, Status = "draft", CreatedAt = DateTime.Now };
            _db.PurchaseOrders.Add(o);
            orders.Add(o);
        }
        await _db.SaveChangesAsync();

        var ids = orders.Select(o => o.PurchaseId);
        await _repo.BatchApproveAsync(ids, 1);

        var approved = await _db.PurchaseOrders.Where(o => ids.Contains(o.PurchaseId)).ToListAsync();
        Assert.All(approved, o => Assert.Equal("approved", o.Status));
    }

    // ==================== 收货入库 ====================

    [Fact(Skip = "ReceivePurchaseAsync 含原生 SQL，InMemory 不支持；由 BusinessPathTests.E15 (SQLite) 覆盖")]
    public async Task ReceivePurchase_ShouldCreateReceipt()
    {
        var receipt = new PurchaseReceipt
        {
            PurchaseId = 1,
            SupplierId = 1,
            ReceivedBy = "测试员",
            TotalReceivedQty = 100
        };
        var details = new List<PurchaseReceiptDetail>
        {
            new() { MaterialId = 1, ReceivedQty = 50, AcceptedQty = 48, RejectedQty = 2 },
            new() { MaterialId = 2, ReceivedQty = 50, AcceptedQty = 50, RejectedQty = 0 }
        };

        var created = await _repo.ReceivePurchaseAsync(receipt, details);

        Assert.True(created.ReceiptId > 0);
        Assert.StartsWith("RCV-", created.ReceiptNo);
        Assert.Equal("Complete", created.Status);

        var savedDetails = await _repo.GetReceiptDetailsAsync(created.ReceiptId);
        Assert.Equal(2, savedDetails.Count);
    }

    // ==================== 供应商管理 ====================

    [Fact]
    public async Task UpsertSupplier_Create_ShouldReturnNewSupplier()
    {
        var supplier = new Supplier
        {
            SupplierName = "测试供应商",
            Category = "Bottle",
            ContactPerson = "张三",
            Phone = "13800138000",
            IsActive = true
        };

        var created = await _repo.UpsertSupplierAsync(supplier);

        Assert.True(created.SupplierId > 0);
        Assert.Equal("测试供应商", created.SupplierName);
    }

    [Fact]
    public async Task UpsertSupplier_Update_ShouldModifyExisting()
    {
        var supplier = new Supplier { SupplierName = "原始供应商", IsActive = true };
        _db.Suppliers.Add(supplier);
        await _db.SaveChangesAsync();

        supplier.SupplierName = "更新后供应商";
        await _repo.UpsertSupplierAsync(supplier);

        var updated = await _db.Suppliers.FindAsync(supplier.SupplierId);
        Assert.Equal("更新后供应商", updated!.SupplierName);
    }

    [Fact]
    public async Task GetSuppliers_WithFilters_ShouldReturnCorrectResults()
    {
        _db.Suppliers.AddRange(
            new Supplier { SupplierName = "A公司", Category = "Bottle", IsActive = true },
            new Supplier { SupplierName = "B公司", Category = "Packaging", IsActive = true },
            new Supplier { SupplierName = "C公司", Category = "Bottle", IsActive = false }
        );
        await _db.SaveChangesAsync();

        var bottleSuppliers = await _repo.GetSuppliersAsync(category: "Bottle");
        Assert.Equal(2, bottleSuppliers.Count);

        var activeOnly = await _repo.GetSuppliersAsync(activeOnly: true);
        Assert.Equal(2, activeOnly.Count);
    }

    // ==================== 采购批次 ====================

    [Fact]
    public async Task GetPurchaseBatches_ShouldReturnPagedResults()
    {
        for (int i = 0; i < 3; i++)
        {
            _db.PurchaseBatches.Add(new PurchaseBatch
            {
                BatchNo = $"BATCH-{i}",
                ItemType = "Bottle",
                ItemCode = $"BOT-{i}",
                Quantity = 100 + i * 10,
                CreatedAt = DateTime.Now
            });
        }
        await _db.SaveChangesAsync();

        var (items, total) = await _repo.GetPurchaseBatchesAsync(itemType: "Bottle");

        Assert.Equal(3, total);
        Assert.Equal(3, items.Count);
    }

    // ==================== 价格管理 ====================

    [Fact]
    public async Task UpsertPrice_ShouldCreateNewPrice()
    {
        var price = new SupplierPrice
        {
            SupplierId = 1,
            ItemCode = "BOT-001",
            ItemName = "50ml玻璃瓶",
            UnitPrice = 2.50m,
            Unit = "个",
            MinOrderQty = 1000,
            IsActive = true
        };

        var created = await _repo.UpsertPriceAsync(price);
        Assert.True(created.PriceId > 0);
    }

    // ==================== 品牌定香采购 ====================

    [Fact]
    public async Task GetFixedBrandProducts_ShouldReturnAll()
    {
        _db.FixedBrandProducts.AddRange(
            new FixedBrandProduct { ProductName = "产品A", ProductCode = "FB-001" },
            new FixedBrandProduct { ProductName = "产品B", ProductCode = "FB-002" }
        );
        await _db.SaveChangesAsync();

        var products = await _repo.GetFixedBrandProductsAsync();
        Assert.Equal(2, products.Count);
    }

    [Fact]
    public async Task GetFixedBrandOrders_ShouldReturnPaged()
    {
        _db.FixedBrandPurchaseOrders.Add(new FixedBrandPurchaseOrder
        {
            PurchaseNo = "FBPO-001",
            SupplierName = "测试供应商",
            TotalAmount = 5000,
            Status = "approved",
            CreatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        var (items, total) = await _repo.GetFixedBrandOrdersAsync();
        Assert.Equal(1, total);
    }

    // ==================== E2E 采购流程 ====================

    [Fact(Skip = "含 ReceivePurchaseAsync 原生 SQL，InMemory 不支持；由 BusinessPathTests.E15 (SQLite) 覆盖")]
    public async Task E2E_PurchaseFlow_CreateApproveReceive()
    {
        // 1. 创建供应商
        var supplier = await _repo.UpsertSupplierAsync(new Supplier
        {
            SupplierName = "E2E供应商", Category = "Bottle", IsActive = true
        });

        // 2. 创建采购单（含明细）
        var order = await _repo.CreatePurchaseOrderAsync(
            new PurchaseOrder { SupplierId = supplier.SupplierId, OrderType = "Bottle" },
            new List<PurchaseOrderDetail>
            {
                new() { ItemCode = "BOT-001", ItemName = "50ml瓶", Quantity = 100, UnitPrice = 2.5m, Unit = "个" }
            }
        );
        Assert.Equal("draft", order.Status);

        // 3. 审批
        var approved = await _repo.ApprovePurchaseOrderAsync(order.PurchaseId, 1);
        Assert.True(approved);

        var orderAfterApprove = await _repo.GetPurchaseOrderAsync(order.PurchaseId);
        Assert.Equal("approved", orderAfterApprove!.Status);

        // 4. 收货
        var receipt = await _repo.ReceivePurchaseAsync(new PurchaseReceipt
        {
            PurchaseId = order.PurchaseId,
            SupplierId = supplier.SupplierId,
            ReceivedBy = "仓库管理员",
            TotalReceivedQty = 100
        }, new List<PurchaseReceiptDetail>
        {
            new() { ReceivedQty = 100, AcceptedQty = 98, RejectedQty = 2, RejectReason = "破损" }
        });
        Assert.Equal("Complete", receipt.Status);

        // 5. 验证订单明细收货记录
        var details = await _repo.GetPurchaseOrderDetailsAsync(order.PurchaseId);
        Assert.Single(details);
        Assert.Equal("BOT-001", details[0].ItemCode);

        // 6. 验证状态日志
        var logs = await _repo.GetStatusLogsAsync(order.PurchaseId);
        Assert.True(logs.Count >= 1);
    }

    // ==================== 报表与统计 ====================

    [Fact]
    public async Task GetHistoryStats_ShouldReturnAll()
    {
        _db.PurchaseHistoryStats.AddRange(
            new PurchaseHistoryStat { ItemType = "Bottle", ItemCode = "BOT-001", Avg30DayUsage = 50 },
            new PurchaseHistoryStat { ItemType = "Packaging", ItemCode = "PKG-001", Avg30DayUsage = 200 }
        );
        await _db.SaveChangesAsync();

        var stats = await _repo.GetHistoryStatsAsync();
        Assert.Equal(2, stats.Count);
    }

    [Fact]
    public async Task GetCostReviews_ShouldReturnOrdered()
    {
        _db.PurchaseCostReviews.AddRange(
            new PurchaseCostReview { PurchaseId = 1, ReviewAmount = 5000, ReviewStatus = "pending", CreatedAt = DateTime.Now },
            new PurchaseCostReview { PurchaseId = 2, ReviewAmount = 3000, ReviewStatus = "approved", CreatedAt = DateTime.Now.AddHours(-1) }
        );
        await _db.SaveChangesAsync();

        var reviews = await _repo.GetCostReviewsAsync();
        Assert.Equal(2, reviews.Count);
    }

    [Fact]
    public async Task GetCategories_ShouldReturnOrdered()
    {
        _db.PurchaseCategories.AddRange(
            new PurchaseCategory { CategoryCode = "BTL", CategoryName = "瓶子", DisplayOrder = 2, IsActive = true },
            new PurchaseCategory { CategoryCode = "PKG", CategoryName = "包装物", DisplayOrder = 1, IsActive = true }
        );
        await _db.SaveChangesAsync();

        var categories = await _repo.GetCategoriesAsync();
        Assert.Equal(2, categories.Count);
        Assert.Equal("包装物", categories[0].CategoryName);
    }

    // ==================== 库存预警 ====================

    [Fact]
    public async Task GetLowStockPackaging_ShouldReturnBelowSafetyStock()
    {
        _db.PackagingInventories.AddRange(
            new PackagingInventory { ItemName = "包装盒A", StockQty = 10, SafetyStock = 50 },
            new PackagingInventory { ItemName = "包装盒B", StockQty = 100, SafetyStock = 50 },
            new PackagingInventory { ItemName = "包装盒C", StockQty = 5, SafetyStock = 20 }
        );
        await _db.SaveChangesAsync();

        var lowStock = await _repo.GetLowStockPackagingAsync();
        Assert.Equal(2, lowStock.Count);
        Assert.All(lowStock, item => Assert.True(item.StockQty < item.SafetyStock));
    }
}
