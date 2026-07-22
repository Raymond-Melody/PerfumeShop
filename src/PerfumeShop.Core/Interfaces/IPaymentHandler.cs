namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 支付/订单状态机接口
/// 状态流转: Pending → Paid → Shipped → Delivered / Refunded / Cancelled
/// 对应 ASP: payment_handler.asp + payment_callback.asp + checkout_payment_processor.asp
/// </summary>
public interface IPaymentHandler
{
    // ========== 创建支付订单 (V18: CreatePaymentOrder / SafeCreatePaymentOrder) ==========

    /// <summary>创建支付订单 — 完整流程（生成订单号、验证参数、插入订单）</summary>
    Task<CreatePaymentOrderResult> CreatePaymentOrderAsync(int userId, decimal orderAmount, string orderDesc,
        int paymentMethod, string? shippingName = null, string? shippingPhone = null,
        string? shippingAddress = null, CancellationToken ct = default);

    /// <summary>创建支付记录</summary>
    Task<int> CreatePaymentAsync(int orderId, string paymentMethod, decimal amount, CancellationToken ct = default);

    // ========== 支付回调处理 (V18: VerifyPaymentCallback) ==========

    /// <summary>处理支付回调 — 验证签名、幂等检查、更新状态</summary>
    Task<PaymentCallbackResult> ProcessCallbackAsync(int paymentMethod, Dictionary<string, string> callbackData, CancellationToken ct = default);

    // ========== 风控检查 (V18: CheckRisk + RC_SafeNum) ==========

    /// <summary>风控检查 — 用户信用、金额异常、地址重复、IP频率</summary>
    Task<RiskCheckResult> CheckRiskAsync(int userId, decimal orderTotal, string? shippingAddress = null,
        string? shippingPhone = null, string? ipAddress = null, CancellationToken ct = default);

    // ========== 支付状态同步 (V18: UpdateOrderPaymentStatus) ==========

    /// <summary>同步支付状态 — 状态机流转</summary>
    Task<bool> SyncPaymentStatusAsync(int orderId, PaymentStatus newStatus, string transactionId, CancellationToken ct = default);

    // ========== 自动创建生产工单 (V18: AutoCreateProductionOrder) ==========

    /// <summary>支付成功后自动创建生产工单 — 写入 ProductionOrders + ProductionLogs，含幂等检查</summary>
    Task<bool> AutoCreateProductionOrderAsync(int orderId, CancellationToken ct = default);

    // ========== 确认支付 (V18: 整合支付确认流程) ==========

    /// <summary>确认支付 — 整合幂等检查+状态更新+自动创建生产工单</summary>
    Task<bool> ConfirmPaymentAsync(int orderId, string transactionId, CancellationToken ct = default);

    // ========== 订单操作 ==========

    /// <summary>取消订单</summary>
    Task<bool> CancelOrderAsync(int orderId, int userId, CancellationToken ct = default);

    /// <summary>确认收货</summary>
    Task<bool> ConfirmDeliveryAsync(int orderId, int userId, CancellationToken ct = default);

    /// <summary>申请退款</summary>
    Task<bool> RequestRefundAsync(int orderId, int userId, decimal amount, string reason, CancellationToken ct = default);

    /// <summary>处理退款</summary>
    Task<bool> ProcessRefundAsync(int orderId, CancellationToken ct = default);
}

/// <summary>支付状态枚举</summary>
public enum PaymentStatus
{
    Pending,
    Paid,
    Failed,
    Refunded,
    Cancelled
}

/// <summary>创建支付订单结果</summary>
public class CreatePaymentOrderResult
{
    public bool Success { get; set; }
    public int OrderId { get; set; }
    public string OrderNo { get; set; } = "";
    public string Message { get; set; } = "";
}

/// <summary>支付回调结果</summary>
public class PaymentCallbackResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public int? OrderId { get; set; }
    public string? TransactionId { get; set; }
    public bool IsDuplicate { get; set; }
}

/// <summary>风控检查结果</summary>
public class RiskCheckResult
{
    public string RiskLevel { get; set; } = "low";
    public int RiskCount { get; set; }
    public bool Passed { get; set; } = true;
    public List<RiskItem> Risks { get; set; } = new();
    public DateTime Timestamp { get; set; } = DateTime.Now;
}

/// <summary>风险项</summary>
public class RiskItem
{
    public string Type { get; set; } = "";
    public string Level { get; set; } = "";
    public string Message { get; set; } = "";
}
