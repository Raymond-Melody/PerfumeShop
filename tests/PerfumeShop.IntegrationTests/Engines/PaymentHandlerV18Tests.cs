using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// PaymentHandler V18 完整算法单元测试
/// 使用 EngineTestContext，每个测试创建新的 Handler 实例（幂等性存储已改为实例级）
/// </summary>
public class PaymentHandlerV18Tests : IDisposable
{
    private readonly EngineTestContext _db;
    private readonly PaymentHandler _handler;

    public PaymentHandlerV18Tests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"PayV18_{Guid.NewGuid()}")
            .Options;
        _db = new EngineTestContext(options);
        _handler = new PaymentHandler(_db);
    }

    public void Dispose()
    {
        _db.Database.EnsureDeleted();
        _db.Dispose();
    }

    // ==================== 辅助 ====================

    private async Task<int> SeedOrderAsync(int userId, decimal amount, string status = "Pending")
    {
        var order = new Order
        {
            OrderNo = $"ORD{DateTime.Now:yyyyMMddHHmmss}{userId:D4}",
            UserId = userId, TotalAmount = amount, Status = status,
            Notes = "Test order", PaymentMethod = "1",
            ShippingName = "Test", ShippingPhone = "13800138000",
            ShippingAddress = "Test Address",
            CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        };
        _db.Orders.Add(order);
        await _db.SaveChangesAsync();
        return order.OrderId;
    }

    private async Task SeedOrderWithDetailsAsync(int orderId, int productId, int qty)
    {
        _db.OrderDetails.Add(new OrderDetail
        {
            OrderId = orderId, ProductId = productId, Quantity = qty,
            UnitPrice = 100m, Subtotal = 100m * qty, ProductName = $"Product{productId}"
        });
        await _db.SaveChangesAsync();
    }

    // ==================== 1. CreatePaymentOrderAsync ====================

    /// <summary>正常创建支付订单</summary>
    [Fact]
    public async Task CreatePaymentOrder_ValidParams_ReturnsSuccess()
    {
        _db.Users.Add(new User { UserId = 1, Username = "u1", Email = "u1@t.com", Password = "h" });
        await _db.SaveChangesAsync();

        var result = await _handler.CreatePaymentOrderAsync(1, 299.99m, "香水订单", 1,
            "张三", "13800138000", "上海市");

        Assert.True(result.Success);
        Assert.True(result.OrderId > 0);
        Assert.StartsWith("ORD", result.OrderNo);
    }

    /// <summary>无效金额 — 返回失败</summary>
    [Fact]
    public async Task CreatePaymentOrder_InvalidAmount_ReturnsFail()
    {
        var result = await _handler.CreatePaymentOrderAsync(1, -10m, "test", 1);
        Assert.False(result.Success);
    }

    /// <summary>无效支付方式 — 返回失败</summary>
    [Fact]
    public async Task CreatePaymentOrder_InvalidPaymentMethod_ReturnsFail()
    {
        var result = await _handler.CreatePaymentOrderAsync(1, 100m, "test", 99);
        Assert.False(result.Success);
    }

    // ==================== 2. ProcessCallbackAsync ====================

    /// <summary>正常支付回调 — 状态更新为 Paid</summary>
    [Fact]
    public async Task ProcessCallback_ValidData_UpdatesToPaid()
    {
        var orderId = await SeedOrderAsync(1, 299m);
        var callbackData = new Dictionary<string, string>
        {
            { "order_id", orderId.ToString() },
            { "transaction_id", "WX-20240101-001" }
        };

        var result = await _handler.ProcessCallbackAsync(1, callbackData);

        Assert.True(result.Success);
        Assert.Equal(orderId, result.OrderId);
        Assert.False(result.IsDuplicate);

        var order = await _db.Orders.FindAsync(orderId);
        Assert.Equal("Paid", order!.Status);
    }

    /// <summary>重复回调（幂等性） — 第二次调用返回 IsDuplicate=true</summary>
    [Fact]
    public async Task ProcessCallback_Duplicate_ReturnsIdempotent()
    {
        var orderId = await SeedOrderAsync(1, 200m);
        var callbackData = new Dictionary<string, string>
        {
            { "order_id", orderId.ToString() },
            { "transaction_id", "WX-DUP-001" }
        };

        var result1 = await _handler.ProcessCallbackAsync(1, callbackData);
        Assert.True(result1.Success);
        Assert.False(result1.IsDuplicate);

        var result2 = await _handler.ProcessCallbackAsync(1, callbackData);
        Assert.True(result2.Success);
        Assert.True(result2.IsDuplicate);
    }

    /// <summary>无效 order_id — 返回失败</summary>
    [Fact]
    public async Task ProcessCallback_InvalidOrderId_ReturnsFail()
    {
        var callbackData = new Dictionary<string, string> { { "transaction_id", "WX-001" } };
        var result = await _handler.ProcessCallbackAsync(1, callbackData);
        Assert.False(result.Success);
    }

    // ==================== 3. CheckRiskAsync ====================

    /// <summary>正常用户 — 风控通过</summary>
    [Fact]
    public async Task CheckRisk_NormalUser_Passes()
    {
        _db.Users.Add(new User { UserId = 1, Username = "normal", Email = "n@t.com", Password = "h" });
        _db.Orders.Add(new Order
        {
            OrderNo = "O1", UserId = 1, TotalAmount = 200, Status = "Paid",
            CreatedAt = DateTime.Now.AddDays(-10), UpdatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        var result = await _handler.CheckRiskAsync(1, 300m);

        Assert.True(result.Passed);
        Assert.Equal("low", result.RiskLevel);
    }

    /// <summary>高退货率 + 低消费用户 — 风控不通过</summary>
    [Fact]
    public async Task CheckRisk_HighReturnRate_LowSpend_Fails()
    {
        _db.Users.Add(new User { UserId = 2, Username = "risky", Email = "r@t.com", Password = "h" });
        // 10 个订单, 4 个退货 (40%), 总消费 60 < 100 (触发 V18 D 级用户判断)
        for (int i = 0; i < 10; i++)
        {
            _db.Orders.Add(new Order
            {
                OrderNo = $"R{i}", UserId = 2, TotalAmount = 10,
                Status = i < 4 ? "Returned" : "Paid",
                CreatedAt = DateTime.Now.AddDays(-30 + i), UpdatedAt = DateTime.Now
            });
        }
        await _db.SaveChangesAsync();

        var result = await _handler.CheckRiskAsync(2, 100m);

        Assert.False(result.Passed);
        Assert.Equal("high", result.RiskLevel);
    }

    // ==================== 4. AutoCreateProductionOrderAsync ====================

    /// <summary>支付成功自动创建生产工单</summary>
    [Fact]
    public async Task AutoCreateProductionOrder_CreatesOrdersAndLogs()
    {
        var orderId = await SeedOrderAsync(1, 300m);
        await SeedOrderWithDetailsAsync(orderId, 101, 2);

        var result = await _handler.AutoCreateProductionOrderAsync(orderId);

        Assert.True(result);

        var prodOrders = await _db.ProductionOrders.Where(po => po.OrderId == orderId).ToListAsync();
        Assert.Equal(2, prodOrders.Count);

        var prodIds = prodOrders.Select(po => po.ProductionId).ToList();
        var logs = await _db.ProductionLogs.Where(l => prodIds.Contains(l.ProductionId)).ToListAsync();
        Assert.Equal(2, logs.Count);
        Assert.All(logs, l => Assert.Equal("SYSTEM", l.CreatedBy));
    }

    /// <summary>幂等性 — 重复调用不创建重复工单</summary>
    [Fact]
    public async Task AutoCreateProductionOrder_Idempotent_NoDuplicates()
    {
        var orderId = await SeedOrderAsync(1, 200m);
        await SeedOrderWithDetailsAsync(orderId, 201, 3);

        await _handler.AutoCreateProductionOrderAsync(orderId);
        var count1 = await _db.ProductionOrders.CountAsync(po => po.OrderId == orderId);

        await _handler.AutoCreateProductionOrderAsync(orderId);
        var count2 = await _db.ProductionOrders.CountAsync(po => po.OrderId == orderId);

        Assert.Equal(count1, count2);
        Assert.Equal(3, count1);
    }

    /// <summary>无订单明细时 — 返回 true 不创建工单</summary>
    [Fact]
    public async Task AutoCreateProductionOrder_NoDetails_ReturnsTrue()
    {
        var orderId = await SeedOrderAsync(1, 100m);
        var result = await _handler.AutoCreateProductionOrderAsync(orderId);
        Assert.True(result);
        Assert.Equal(0, await _db.ProductionOrders.CountAsync(po => po.OrderId == orderId));
    }

    // ==================== 5. ConfirmPaymentAsync ====================

    /// <summary>ConfirmPayment — 整合幂等+状态+生产工单</summary>
    [Fact]
    public async Task ConfirmPayment_FullFlow_UpdatesAll()
    {
        var orderId = await SeedOrderAsync(1, 500m);
        await SeedOrderWithDetailsAsync(orderId, 301, 1);
        _db.PaymentRecords.Add(new PaymentRecord
        {
            OrderId = orderId, PaymentMethod = "wechat", Amount = 500m,
            Status = "Pending", CreatedAt = DateTime.Now
        });
        await _db.SaveChangesAsync();

        var result = await _handler.ConfirmPaymentAsync(orderId, "TX-CONFIRM-001");

        Assert.True(result);

        var order = await _db.Orders.FindAsync(orderId);
        Assert.Equal("Paid", order!.Status);

        var payRecord = await _db.PaymentRecords.FirstAsync(pr => pr.OrderId == orderId);
        Assert.Equal("Paid", payRecord.Status);
        Assert.Equal("TX-CONFIRM-001", payRecord.TransactionNo);

        var prodCount = await _db.ProductionOrders.CountAsync(po => po.OrderId == orderId);
        Assert.Equal(1, prodCount);
    }

    // ==================== 6. SyncPaymentStatusAsync ====================

    /// <summary>状态同步 — Pending→Paid 触发生产工单</summary>
    [Fact]
    public async Task SyncPaymentStatus_Paid_TriggersProductionOrder()
    {
        var orderId = await SeedOrderAsync(1, 400m);
        await SeedOrderWithDetailsAsync(orderId, 401, 2);

        var result = await _handler.SyncPaymentStatusAsync(orderId, PaymentStatus.Paid, "TX-SYNC-001");

        Assert.True(result);
        var order = await _db.Orders.FindAsync(orderId);
        Assert.Equal("Paid", order!.Status);
        Assert.Contains("TX-SYNC-001", order.Notes!);

        var prodCount = await _db.ProductionOrders.CountAsync(po => po.OrderId == orderId);
        Assert.Equal(2, prodCount);
    }

    // ==================== 7. 接口覆盖检查 ====================

    /// <summary>IPaymentHandler 覆盖 V18 所有支付函数</summary>
    [Fact]
    public void IPaymentHandler_CoversAllV18Functions()
    {
        var methods = typeof(IPaymentHandler).GetMethods().Select(m => m.Name).ToHashSet();

        Assert.Contains("CreatePaymentOrderAsync", methods);
        Assert.Contains("CreatePaymentAsync", methods);
        Assert.Contains("ProcessCallbackAsync", methods);
        Assert.Contains("CheckRiskAsync", methods);
        Assert.Contains("SyncPaymentStatusAsync", methods);
        Assert.Contains("AutoCreateProductionOrderAsync", methods);
        Assert.Contains("ConfirmPaymentAsync", methods);
        Assert.Contains("CancelOrderAsync", methods);
        Assert.Contains("ConfirmDeliveryAsync", methods);
        Assert.Contains("RequestRefundAsync", methods);
        Assert.Contains("ProcessRefundAsync", methods);
    }
}
