namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 支付/订单状态机接口
/// </summary>
public interface IPaymentHandler
{
    /// <summary>创建支付记录</summary>
    Task<int> CreatePaymentAsync(int orderId, string paymentMethod, decimal amount, CancellationToken ct = default);

    /// <summary>确认支付 — 更新订单状态为 Paid</summary>
    Task<bool> ConfirmPaymentAsync(int orderId, string transactionId, CancellationToken ct = default);

    /// <summary>取消订单</summary>
    Task<bool> CancelOrderAsync(int orderId, int userId, CancellationToken ct = default);

    /// <summary>确认收货</summary>
    Task<bool> ConfirmDeliveryAsync(int orderId, int userId, CancellationToken ct = default);

    /// <summary>申请退款</summary>
    Task<bool> RequestRefundAsync(int orderId, int userId, decimal amount, string reason, CancellationToken ct = default);

    /// <summary>处理退款</summary>
    Task<bool> ProcessRefundAsync(int orderId, CancellationToken ct = default);
}
