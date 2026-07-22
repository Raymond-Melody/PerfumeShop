using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests.Engines;

/// <summary>
/// PaymentHandler V18 关键路径 SQLite 测试
/// 使用真实 SQLite 内存数据库验证（非 InMemory Provider）
/// 重点验证：ExecuteUpdateAsync 在真实 DB、幂等性、ProductionOrders 创建、事务一致性
/// </summary>
public class PaymentHandlerV18SqliteTests : SqliteTestBase
{
    private PaymentHandler CreateHandler() => new(Db);

    // ==================== 辅助 ====================

    private async Task<int> SeedOrderAsync(int userId, decimal amount, string status = "Pending")
    {
        var order = new Order
        {
            OrderNo = $"ORD{DateTime.Now:yyyyMMddHHmmss}{userId:D4}{Guid.NewGuid():N}"[..40],
            UserId = userId, TotalAmount = amount, Status = status,
            Notes = "Test order", PaymentMethod = "1",
            ShippingName = "Test", ShippingPhone = "13800138000",
            ShippingAddress = "Test Address",
            CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
        };
        Db.Orders.Add(order);
        await Db.SaveChangesAsync();
        return order.OrderId;
    }

    private async Task SeedOrderWithDetailsAsync(int orderId, int productId, int qty)
    {
        Db.OrderDetails.Add(new OrderDetail
        {
            OrderId = orderId, ProductId = productId, Quantity = qty,
            UnitPrice = 100m, Subtotal = 100m * qty, ProductName = $"Product{productId}"
        });
        await Db.SaveChangesAsync();
    }

    // ==================== 1. CreatePaymentOrderAsync ====================

    /// <summary>SQLite 创建支付订单 — 验证真实 DB 插入</summary>
    [Fact]
    public async Task CreatePaymentOrder_Sqlite_InsertsOrder()
    {
        Db.Users.Add(new User { UserId = 1, Username = "u1", Email = "u1@t.com", Password = "h" });
        await Db.SaveChangesAsync();

        var result = await CreateHandler().CreatePaymentOrderAsync(1, 299.99m, "香水订单", 1,
            "张三", "13800138000", "上海市");

        Assert.True(result.Success);
        Assert.True(result.OrderId > 0);
        Assert.StartsWith("ORD", result.OrderNo);

        // 验证真实 DB 中订单存在
        var order = await Db.Orders.FindAsync(result.OrderId);
        Assert.NotNull(order);
        Assert.Equal("Pending", order.Status);
    }

    // ==================== 2. ProcessCallbackAsync 幂等性 ====================

    /// <summary>ProcessCallback SQLite — 幂等性在真实 DB（第二次调用 IsDuplicate=true）</summary>
    [Fact]
    public async Task ProcessCallback_Sqlite_Idempotent()
    {
        var orderId = await SeedOrderAsync(1, 200m);
        var callbackData = new Dictionary<string, string>
        {
            { "order_id", orderId.ToString() },
            { "transaction_id", "WX-DUP-001" }
        };

        var handler = CreateHandler();
        var result1 = await handler.ProcessCallbackAsync(1, callbackData);
        Assert.True(result1.Success);
        Assert.False(result1.IsDuplicate);

        // 验证 DB 状态已更新
        var order = await Db.Orders.FindAsync(orderId);
        Assert.Equal("Paid", order!.Status);

        // 重复回调
        var result2 = await handler.ProcessCallbackAsync(1, callbackData);
        Assert.True(result2.Success);
        Assert.True(result2.IsDuplicate);
    }

    // ==================== 3. CheckRiskAsync ====================

    /// <summary>CheckRisk SQLite — 正常用户风控通过（真实 DB JOIN 查询）</summary>
    [Fact]
    public async Task CheckRisk_Sqlite_NormalUser_Passes()
    {
        Db.Users.Add(new User { UserId = 1, Username = "normal", Email = "n@t.com", Password = "h" });
        Db.Orders.Add(new Order
        {
            OrderNo = "O1", UserId = 1, TotalAmount = 200, Status = "Paid",
            CreatedAt = DateTime.Now.AddDays(-10), UpdatedAt = DateTime.Now
        });
        await Db.SaveChangesAsync();

        var result = await CreateHandler().CheckRiskAsync(1, 300m);

        Assert.True(result.Passed);
        Assert.Equal("low", result.RiskLevel);
    }

    /// <summary>CheckRisk SQLite — 高退货率+低消费用户风控不通过</summary>
    [Fact]
    public async Task CheckRisk_Sqlite_HighReturnRate_Fails()
    {
        Db.Users.Add(new User { UserId = 2, Username = "risky", Email = "r@t.com", Password = "h" });
        for (int i = 0; i < 10; i++)
        {
            Db.Orders.Add(new Order
            {
                OrderNo = $"R{i}", UserId = 2, TotalAmount = 10,
                Status = i < 4 ? "Returned" : "Paid",
                CreatedAt = DateTime.Now.AddDays(-30 + i), UpdatedAt = DateTime.Now
            });
        }
        await Db.SaveChangesAsync();

        var result = await CreateHandler().CheckRiskAsync(2, 100m);

        Assert.False(result.Passed);
        Assert.Equal("high", result.RiskLevel);
    }

    // ==================== 4. AutoCreateProductionOrderAsync ====================

    /// <summary>AutoCreateProductionOrder SQLite — 验证 ProductionOrders + ProductionLogs 写入</summary>
    [Fact]
    public async Task AutoCreateProductionOrder_Sqlite_CreatesOrdersAndLogs()
    {
        var orderId = await SeedOrderAsync(1, 300m);
        await SeedOrderWithDetailsAsync(orderId, 101, 2);

        var result = await CreateHandler().AutoCreateProductionOrderAsync(orderId);
        Assert.True(result);

        var prodOrders = await Db.ProductionOrders.Where(po => po.OrderId == orderId).ToListAsync();
        Assert.Equal(2, prodOrders.Count);

        var prodIds = prodOrders.Select(po => po.ProductionId).ToList();
        var logs = await Db.ProductionLogs.Where(l => prodIds.Contains(l.ProductionId)).ToListAsync();
        Assert.Equal(2, logs.Count);
        Assert.All(logs, l => Assert.Equal("SYSTEM", l.CreatedBy));
    }

    /// <summary>AutoCreateProductionOrder SQLite — 幂等性（重复调用不创建重复工单）</summary>
    [Fact]
    public async Task AutoCreateProductionOrder_Sqlite_Idempotent()
    {
        var orderId = await SeedOrderAsync(1, 200m);
        await SeedOrderWithDetailsAsync(orderId, 201, 3);

        var handler = CreateHandler();
        await handler.AutoCreateProductionOrderAsync(orderId);
        var count1 = await Db.ProductionOrders.CountAsync(po => po.OrderId == orderId);

        await handler.AutoCreateProductionOrderAsync(orderId);
        var count2 = await Db.ProductionOrders.CountAsync(po => po.OrderId == orderId);

        Assert.Equal(count1, count2);
        Assert.Equal(3, count1);
    }

    // ==================== 5. ConfirmPaymentAsync ====================

    /// <summary>ConfirmPayment SQLite — 整合流程（跟踪实体更新 + 生产工单创建在真实 DB）</summary>
    [Fact]
    public async Task ConfirmPayment_Sqlite_FullFlowUpdatesAll()
    {
        var orderId = await SeedOrderAsync(1, 500m);
        await SeedOrderWithDetailsAsync(orderId, 301, 1);
        Db.PaymentRecords.Add(new PaymentRecord
        {
            OrderId = orderId, PaymentMethod = "wechat", Amount = 500m,
            Status = "Pending", CreatedAt = DateTime.Now
        });
        await Db.SaveChangesAsync();

        var result = await CreateHandler().ConfirmPaymentAsync(orderId, "TX-CONFIRM-001");
        Assert.True(result);

        // 验证订单状态
        var order = await Db.Orders.FindAsync(orderId);
        Assert.Equal("Paid", order!.Status);

        // 验证支付记录
        var payRecord = await Db.PaymentRecords.FirstAsync(pr => pr.OrderId == orderId);
        Assert.Equal("Paid", payRecord.Status);
        Assert.Equal("TX-CONFIRM-001", payRecord.TransactionNo);

        // 验证生产工单
        var prodCount = await Db.ProductionOrders.CountAsync(po => po.OrderId == orderId);
        Assert.Equal(1, prodCount);
    }

    // ==================== 6. SyncPaymentStatusAsync ====================

    /// <summary>SyncPaymentStatus SQLite — Paid 触发 AutoCreateProductionOrder</summary>
    [Fact]
    public async Task SyncPaymentStatus_Sqlite_PaidTriggersProduction()
    {
        var orderId = await SeedOrderAsync(1, 400m);
        await SeedOrderWithDetailsAsync(orderId, 401, 2);

        var result = await CreateHandler().SyncPaymentStatusAsync(orderId, PaymentStatus.Paid, "TX-SYNC-001");
        Assert.True(result);

        var order = await Db.Orders.FindAsync(orderId);
        Assert.Equal("Paid", order!.Status);

        var prodCount = await Db.ProductionOrders.CountAsync(po => po.OrderId == orderId);
        Assert.Equal(2, prodCount);
    }

    // ==================== 7. CancelOrderAsync ====================

    /// <summary>CancelOrder SQLite — ExecuteUpdateAsync 在真实 DB（AsNoTracking 验证）</summary>
    [Fact]
    public async Task CancelOrder_Sqlite_ExecuteUpdateWorks()
    {
        var orderId = await SeedOrderAsync(1, 100m);

        var result = await CreateHandler().CancelOrderAsync(orderId, 1);
        Assert.True(result);

        // AsNoTracking 绕过 change tracker 缓存（ExecuteUpdateAsync 直接写 DB）
        var order = await Db.Orders.AsNoTracking().FirstAsync(o => o.OrderId == orderId);
        Assert.Equal("Cancelled", order.Status);
    }
}
