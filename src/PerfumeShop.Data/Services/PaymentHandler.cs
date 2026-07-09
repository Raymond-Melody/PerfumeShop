using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 支付处理器 — 订单状态机
/// 状态流转: Pending → Paid → Shipped → Delivered / Refunded / Cancelled
/// </summary>
public class PaymentHandler : IPaymentHandler
{
    private readonly PerfumeShopContext _db;

    public PaymentHandler(PerfumeShopContext db)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
    }

    public async Task<int> CreatePaymentAsync(int orderId, string paymentMethod, decimal amount, CancellationToken ct = default)
    {
        var record = new PaymentRecord
        {
            OrderId = orderId,
            PaymentMethod = paymentMethod,
            Amount = amount,
            Status = "Pending",
            CreatedAt = DateTime.Now
        };
        _db.PaymentRecords.Add(record);
        await _db.SaveChangesAsync(ct);
        return record.RecordId;
    }

    public async Task<bool> ConfirmPaymentAsync(int orderId, string transactionId, CancellationToken ct = default)
    {
        // 更新支付记录
        await _db.PaymentRecords
            .Where(pr => pr.OrderId == orderId && pr.Status == "Pending")
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(pr => pr.Status, "Paid")
                .SetProperty(pr => pr.TransactionNo, transactionId)
                .SetProperty(pr => pr.UpdatedAt, DateTime.Now), ct);

        // 更新订单状态
        await _db.Orders
            .Where(o => o.OrderId == orderId && o.Status == "Pending")
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(o => o.Status, "Paid")
                .SetProperty(o => o.UpdatedAt, DateTime.Now), ct);

        return true;
    }

    public async Task<bool> CancelOrderAsync(int orderId, int userId, CancellationToken ct = default)
    {
        var rows = await _db.Orders
            .Where(o => o.OrderId == orderId && o.UserId == userId &&
                        (o.Status == "Pending" || o.Status == "Paid"))
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(o => o.Status, "Cancelled")
                .SetProperty(o => o.UpdatedAt, DateTime.Now), ct);

        return rows > 0;
    }

    public async Task<bool> ConfirmDeliveryAsync(int orderId, int userId, CancellationToken ct = default)
    {
        var rows = await _db.Orders
            .Where(o => o.OrderId == orderId && o.UserId == userId && o.Status == "Shipped")
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(o => o.Status, "Delivered")
                .SetProperty(o => o.DeliveredAt, DateTime.Now)
                .SetProperty(o => o.UpdatedAt, DateTime.Now), ct);

        return rows > 0;
    }

    public async Task<bool> RequestRefundAsync(int orderId, int userId, decimal amount, string reason, CancellationToken ct = default)
    {
        var order = await _db.Orders
            .FirstOrDefaultAsync(o => o.OrderId == orderId && o.UserId == userId, ct);

        if (order == null || order.Status == "Cancelled" || order.Status == "Refunded")
            return false;

        var refund = new RefundRecord
        {
            OrderId = orderId,
            RefundAmount = amount,
            RefundReason = reason,
            Status = "Pending",
            CreatedAt = DateTime.Now
        };
        _db.RefundRecords.Add(refund);
        await _db.SaveChangesAsync(ct);
        return true;
    }

    public async Task<bool> ProcessRefundAsync(int orderId, CancellationToken ct = default)
    {
        // 更新退款记录
        await _db.RefundRecords
            .Where(r => r.OrderId == orderId && r.Status == "Pending")
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(r => r.Status, "Refunded")
                .SetProperty(r => r.CompletedAt, DateTime.Now), ct);

        // 更新订单状态
        await _db.Orders
            .Where(o => o.OrderId == orderId)
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(o => o.Status, "Refunded")
                .SetProperty(o => o.UpdatedAt, DateTime.Now), ct);

        return true;
    }
}
