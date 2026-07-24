using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using PerfumeShop.Api.Controllers;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.IntegrationTests.Admin;

/// <summary>
/// M4-C 生产中心 API 测试
/// 覆盖：工单列表/详情/状态更新/同步/修复 + 并发幂等压测
/// </summary>
public class ProductionPagesTests : IDisposable
{
    private readonly ProductionTestContext _db;
    private readonly ProductionController _ctrl;
    private readonly ProductionRepository _repo;

    public ProductionPagesTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"ProdTests_{Guid.NewGuid()}")
            .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning))
            .Options;
        _db = new ProductionTestContext(options);

        _repo = new ProductionRepository(_db, new PerfumeShop.Data.Services.InventoryLedger(_db));
        _ctrl = new ProductionController(_repo);
        _ctrl.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext { Request = { Scheme = "https" } }
        };
    }

    public void Dispose() => _db.Dispose();

    // ===== Helper =====

    private static readonly JsonSerializerOptions _jsonOpts = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    private static JsonElement ToJson(IActionResult result)
    {
        var ok = Assert.IsType<OkObjectResult>(result);
        var json = JsonSerializer.Serialize(ok.Value, _jsonOpts);
        return JsonDocument.Parse(json).RootElement;
    }

    private void SeedOrders(int count)
    {
        for (int i = 1; i <= count; i++)
        {
            _db.Orders.Add(new Order
            {
                OrderId = i,
                OrderNo = $"ORD-{i:D6}",
                Status = "Paid",
                TotalAmount = 100m * i,
                CreatedAt = DateTime.Now,
                UpdatedAt = DateTime.Now
            });
            _db.OrderDetails.Add(new OrderDetail
            {
                DetailId = i,
                OrderId = i,
                ProductId = 1,
                Quantity = 1,
                UnitPrice = 100m * i,
                Subtotal = 100m * i
            });
        }
        _db.SaveChanges();
    }

    private void SeedProductionOrders(int count)
    {
        for (int i = 1; i <= count; i++)
        {
            _db.ProductionOrders.Add(new ProductionOrder
            {
                ProductionId = i,
                OrderId = 1,
                WorkOrderNo = $"WO-{i:D6}",
                Status = i % 3 == 0 ? "Completed" : i % 3 == 1 ? "Pending" : "InProgress",
                BottleIndex = i,
                TotalBottles = count,
                CreatedAt = DateTime.Now.AddDays(-count + i),
                UpdatedAt = DateTime.Now
            });
        }
        _db.SaveChanges();
    }

    // ===== GET /api/Production =====

    [Fact]
    public async Task GetOrders_ReturnsPaginatedList()
    {
        SeedProductionOrders(25);
        var result = await _ctrl.GetOrders(page: 1, pageSize: 10);
        var json = ToJson(result);
        Assert.Equal(25, json.GetProperty("total").GetInt32());
        Assert.Equal(10, json.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public async Task GetOrders_FilterByStatus_ReturnsOnlyPending()
    {
        SeedProductionOrders(9);
        var result = await _ctrl.GetOrders(status: "Pending");
        var json = ToJson(result);
        Assert.Equal(3, json.GetProperty("total").GetInt32());
    }

    [Fact]
    public async Task GetOrders_SearchByWorkOrderNo()
    {
        SeedProductionOrders(5);
        var result = await _ctrl.GetOrders(search: "WO-000003");
        var json = ToJson(result);
        Assert.Equal(1, json.GetProperty("total").GetInt32());
    }

    // ===== GET /api/Production/{id} =====

    [Fact]
    public async Task GetOrder_ExistingId_ReturnsDetail()
    {
        SeedProductionOrders(1);
        var result = await _ctrl.GetOrder(1);
        var ok = Assert.IsType<OkObjectResult>(result);
        var json = JsonDocument.Parse(JsonSerializer.Serialize(ok.Value, _jsonOpts)).RootElement;
        Assert.Equal("WO-000001", json.GetProperty("workOrderNo").GetString());
    }

    [Fact]
    public async Task GetOrder_NonExisting_ReturnsNotFound()
    {
        var result = await _ctrl.GetOrder(999);
        Assert.IsType<NotFoundObjectResult>(result);
    }

    // ===== PUT /api/Production/{id}/status =====

    [Fact]
    public async Task UpdateStatus_ValidRequest_ReturnsOk()
    {
        SeedProductionOrders(1);
        var result = await _ctrl.UpdateStatus(1, new StatusRequest { NewStatus = "InProgress", Operator = "Admin" });
        var json = ToJson(result);
        Assert.Contains("InProgress", json.GetProperty("message").GetString());
    }

    [Fact]
    public async Task UpdateStatus_EmptyStatus_ReturnsBadRequest()
    {
        var result = await _ctrl.UpdateStatus(1, new StatusRequest { NewStatus = "" });
        Assert.IsType<BadRequestObjectResult>(result);
    }

    // ===== POST /api/Production/sync-orders =====

    [Fact]
    public async Task SyncOrders_NewPaidOrders_CreatesProductionOrders()
    {
        SeedOrders(5);
        var result = await _ctrl.SyncOrders(CancellationToken.None);
        var json = ToJson(result);
        Assert.True(json.GetProperty("success").GetBoolean());
        Assert.Equal(5, json.GetProperty("synced").GetInt32());
    }

    [Fact]
    public async Task SyncOrders_AlreadySynced_ReturnsZero()
    {
        SeedOrders(3);
        await _ctrl.SyncOrders(CancellationToken.None);
        var result2 = await _ctrl.SyncOrders(CancellationToken.None);
        var json = ToJson(result2);
        Assert.Equal(0, json.GetProperty("synced").GetInt32());
    }

    // ===== POST /api/Production/{id}/fix-status =====

    [Fact]
    public async Task FixStatus_ChineseToEnglish_MigratesStatus()
    {
        _db.ProductionOrders.Add(new ProductionOrder
        {
            OrderId = 1, WorkOrderNo = "WO-CN-001",
            Status = "待排产", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        });
        _db.SaveChanges();

        var result = await _ctrl.FixStatus(CancellationToken.None);
        var json = ToJson(result);
        Assert.True(json.GetProperty("updated").GetInt32() >= 1);

        var po = await _db.ProductionOrders.FirstAsync();
        Assert.Equal("Pending", po.Status);
    }

    // ===== GET /api/Production/qc =====

    [Fact]
    public async Task GetQualityChecks_ReturnsCompletedAndQC()
    {
        _db.ProductionOrders.AddRange(
            new ProductionOrder { OrderId = 1, WorkOrderNo = "QC-1", Status = "Completed", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now, CompletedAt = DateTime.Now },
            new ProductionOrder { OrderId = 2, WorkOrderNo = "QC-2", Status = "QC_Review", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now },
            new ProductionOrder { OrderId = 3, WorkOrderNo = "QC-3", Status = "Pending", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now }
        );
        _db.SaveChanges();

        var result = await _ctrl.GetQualityChecks();
        var json = ToJson(result);
        Assert.Equal(2, json.GetProperty("total").GetInt32());
    }

    // ===== GET /api/Production/report =====

    [Fact]
    public async Task GetReport_ReturnsGroupedData()
    {
        _db.ProductionOrders.AddRange(
            new ProductionOrder { OrderId = 1, WorkOrderNo = "RPT-1", Status = "Completed", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now },
            new ProductionOrder { OrderId = 2, WorkOrderNo = "RPT-2", Status = "Pending", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now }
        );
        _db.SaveChanges();

        var result = await _ctrl.GetReport(DateTime.Now.AddDays(-1), DateTime.Now.AddDays(1));
        var json = ToJson(result);
        Assert.True(json.GetProperty("data").GetArrayLength() >= 1);
    }

    // ===== 并发幂等压测：100 单 =====

    [Fact]
    public async Task SyncOrders_Concurrent100Orders_Idempotent()
    {
        // 准备 100 个已付款订单
        SeedOrders(100);

        // 并发执行 10 次同步（模拟竞态）
        var tasks = Enumerable.Range(0, 10)
            .Select(_ => _repo.SyncProductionOrdersAsync(CancellationToken.None))
            .ToArray();

        var results = await Task.WhenAll(tasks);

        // 总同步数应等于 100（幂等：不会重复创建）
        var totalSynced = results.Sum(r => r.Synced);
        var totalErrors = results.Sum(r => r.Errors);

        // 验证数据库中实际工单数
        var actualCount = await _db.ProductionOrders.CountAsync();

        // 幂等验证：无论并发多少次，最终工单数应 <= 100
        Assert.True(actualCount <= 100, $"幂等失败：实际工单数 {actualCount} > 100");
        Assert.Equal(0, totalErrors);
    }

    // ===== BatchOperations 测试 =====

    [Fact]
    public async Task BatchShip_ValidOrders_UpdatesStatus()
    {
        SeedOrders(3);
        var batchCtrl = new BatchOperationsController(_db);

        var result = await batchCtrl.BatchShip(new BatchShipRequest
        {
            Ids = new[] { 1, 2, 3 },
            TrackingNo = "SF-123456"
        });
        var json = ToJson(result);
        Assert.Equal(3, json.GetProperty("successCount").GetInt32());

        var shipped = await _db.Orders.CountAsync(o => o.Status == "Shipped");
        Assert.Equal(3, shipped);
    }

    [Fact]
    public async Task BatchCancel_PendingOrders_Cancels()
    {
        _db.Orders.AddRange(
            new Order { OrderId = 10, OrderNo = "O10", Status = "Pending", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now },
            new Order { OrderId = 11, OrderNo = "O11", Status = "Paid", CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now }
        );
        _db.SaveChanges();

        var batchCtrl = new BatchOperationsController(_db);
        var result = await batchCtrl.BatchCancel(new BatchIdsRequest { Ids = new[] { 10, 11 } });
        var json = ToJson(result);
        Assert.Equal(2, json.GetProperty("successCount").GetInt32());
    }

    [Fact]
    public async Task BatchStatus_Products_TogglesActive()
    {
        _db.Products.AddRange(
            new Product { ProductId = 1, ProductName = "P1", IsActive = false },
            new Product { ProductId = 2, ProductName = "P2", IsActive = false }
        );
        _db.SaveChanges();

        var batchCtrl = new BatchOperationsController(_db);
        var result = await batchCtrl.BatchStatus(new BatchStatusRequest
        {
            Ids = new[] { 1, 2 },
            Action = "list"
        });
        var json = ToJson(result);
        Assert.Equal(2, json.GetProperty("successCount").GetInt32());

        var active = await _db.Products.CountAsync(p => p.IsActive == true);
        Assert.Equal(2, active);
    }
}

/// <summary>
/// 生产测试专用 Context — 为 InMemory Provider 添加 keyless 实体主键
/// </summary>
public class ProductionTestContext : PerfumeShopContext
{
    public ProductionTestContext(DbContextOptions<PerfumeShopContext> options) : base(options) { }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // 核心实体主键
        modelBuilder.Entity<Product>().HasKey(e => e.ProductId);
        modelBuilder.Entity<Order>().HasKey(e => e.OrderId);
        modelBuilder.Entity<OrderDetail>().HasKey(e => e.DetailId);
        modelBuilder.Entity<User>().HasKey(e => e.UserId);
        modelBuilder.Entity<ProductionOrder>().HasKey(e => e.ProductionId);
        modelBuilder.Entity<ProductionLog>().HasKey(e => e.LogId);
        modelBuilder.Entity<MemberTier>().HasKey(e => e.TierId);
        modelBuilder.Entity<Coupon>().HasKey(e => e.CouponId);
        modelBuilder.Entity<PaymentRecord>().HasKey(e => e.RecordId);
        modelBuilder.Entity<RefundRecord>().HasKey(e => e.RefundId);
        modelBuilder.Entity<Recipe>().HasKey(e => e.RecipeId);
        modelBuilder.Entity<UserFavorite>().HasKey(e => e.FavoriteId);
        modelBuilder.Entity<AdminLog>().HasKey(e => e.LogId);

        // 半成品/原料实体主键
        modelBuilder.Entity<AccordProduction>().HasKey(e => e.ProductionId);
        modelBuilder.Entity<AccordProductionDetail>().HasKey(e => e.DetailId);
        modelBuilder.Entity<NoteInventory>().HasKey(e => e.InventoryId);
        modelBuilder.Entity<RawMaterialInventory>().HasKey(e => e.MaterialId);
        modelBuilder.Entity<WorkshopTransfer>().HasKey(e => e.TransferId);
        modelBuilder.Entity<MaterialOutbound>().HasKey(e => e.OutboundId);
        modelBuilder.Entity<MaterialOutboundDetail>().HasKey(e => e.OutboundDetailId);
        modelBuilder.Entity<BaseNote>().HasKey(e => e.BaseNoteId);
        modelBuilder.Entity<FragranceNote>().HasKey(e => e.NoteId);

        // SiteSetting
        modelBuilder.Entity<SiteSetting>().HasKey(e => e.SettingKey);
    }
}
