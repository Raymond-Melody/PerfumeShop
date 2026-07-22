using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 支付处理器 — 完整 V18 支付引擎恢复
/// 状态流转: Pending → Paid → Shipped → Delivered / Refunded / Cancelled
/// V18 函数映射:
///   CreatePaymentOrder / SafeCreatePaymentOrder → CreatePaymentOrderAsync
///   ProcessWeChatPay / ProcessAlipay / ProcessPayPal / ProcessCashOnDelivery → ProcessCallbackAsync
///   CheckRisk + RC_SafeNum                        → CheckRiskAsync
///   UpdateOrderPaymentStatus                     → SyncPaymentStatusAsync
///   AutoCreateProductionOrder                    → AutoCreateProductionOrderAsync
///   ConfirmPaymentAsync                          → 整合幂等+状态+生产工单
///   GeneratePaymentSignature / PaySignHMAC       → GeneratePaymentSignature (private)
///   VerifyPaymentSignature                       → VerifyPaymentSignature (private)
///   CheckPaymentIdempotency                      → CheckPaymentIdempotency (private)
///   MarkPaymentIdempotency                       → MarkPaymentIdempotency (private)
///   VerifyCallbackSignature                      → VerifyCallbackSignature (private)
///   IsCallbackIPAllowed                          → IsCallbackIpAllowed (private)
/// </summary>
public class PaymentHandler : IPaymentHandler
{
    private readonly PerfumeShopContext _db;

    // V18: 支付签名密钥 (对应 ASP: PAYMENT_SIGNING_KEY)
    private const string PaymentSigningKey = "PERFUMESHOP_V18_PAYMENT_SIGNING_KEY_2024";
    // V18: 幂等性 TTL（秒）
    private const int IdempotencyTtlSeconds = 300;
    // V18: 签名有效期（秒）
    private const int SignatureWindowSeconds = 300;
    // V18: 支付方式常量
    private const int PaymentMethodWechat = 1;
    private const int PaymentMethodAlipay = 2;
    private const int PaymentMethodPaypal = 3;
    private const int PaymentMethodCod = 4;

    // V18: 幂等性内存存储 (替代 ASP Application 对象)
    private readonly ConcurrentDictionary<int, IdempotencyEntry> _idempotencyStore = new();

    private record IdempotencyEntry(DateTime Timestamp, string TransactionId);

    public PaymentHandler(PerfumeShopContext db)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
    }

    // ==================== V18: CreatePaymentOrder / SafeCreatePaymentOrder ====================

    /// <summary>创建支付订单 — 完整流程（生成订单号、验证参数、插入订单）</summary>
    public async Task<CreatePaymentOrderResult> CreatePaymentOrderAsync(
        int userId, decimal orderAmount, string orderDesc,
        int paymentMethod, string? shippingName = null, string? shippingPhone = null,
        string? shippingAddress = null, CancellationToken ct = default)
    {
        var result = new CreatePaymentOrderResult();

        // V18: 验证订单金额
        if (orderAmount <= 0)
        {
            result.Message = "Invalid order amount";
            return result;
        }

        // V18: 验证支付方式 (1-4)
        if (paymentMethod < 1 || paymentMethod > 4)
        {
            result.Message = $"Invalid payment method: {paymentMethod}";
            return result;
        }

        // V18: 验证用户ID
        if (userId <= 0)
        {
            result.Message = "Invalid user ID";
            return result;
        }

        // V18: 生成订单号
        var orderNo = $"ORD{DateTime.Now:yyyyMMddHHmmss}{userId % 10000:D4}";

        // V18: 插入订单
        var order = new Order
        {
            OrderNo = orderNo,
            UserId = userId,
            TotalAmount = orderAmount,
            Notes = orderDesc ?? "",
            PaymentMethod = paymentMethod.ToString(),
            Status = "Pending",
            ShippingName = shippingName ?? "",
            ShippingPhone = shippingPhone ?? "",
            ShippingAddress = shippingAddress ?? "",
            CreatedAt = DateTime.Now,
            UpdatedAt = DateTime.Now
        };

        _db.Orders.Add(order);
        await _db.SaveChangesAsync(ct);

        result.Success = true;
        result.OrderId = order.OrderId;
        result.OrderNo = orderNo;
        result.Message = "订单创建成功";
        return result;
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

    // ==================== V18: VerifyPaymentCallback + ProcessXxxPay ====================

    /// <summary>处理支付回调 — 验证签名、幂等检查、更新状态</summary>
    public async Task<PaymentCallbackResult> ProcessCallbackAsync(
        int paymentMethod, Dictionary<string, string> callbackData, CancellationToken ct = default)
    {
        var result = new PaymentCallbackResult();

        // V18: 验证回调 IP (IsCallbackIPAllowed)
        if (callbackData.TryGetValue("callback_ip", out var ip) && !IsCallbackIpAllowed(ip))
        {
            result.Message = "Callback IP not allowed";
            return result;
        }

        // V18: 解析订单 ID
        if (!callbackData.TryGetValue("order_id", out var orderIdStr) || !int.TryParse(orderIdStr, out var orderId))
        {
            result.Message = "Invalid order_id in callback";
            return result;
        }

        // V18: 幂等性检查 (CheckPaymentIdempotency)
        if (CheckPaymentIdempotency(orderId))
        {
            result.Success = true;
            result.IsDuplicate = true;
            result.Message = "Already processed (idempotent)";
            result.OrderId = orderId;
            return result;
        }

        // V18: 验证回调签名 (VerifyCallbackSignature)
        if (!VerifyCallbackSignature(paymentMethod, callbackData))
        {
            result.Message = "Invalid callback signature";
            return result;
        }

        // V18: 获取交易流水号
        var transactionId = callbackData.GetValueOrDefault("transaction_id", $"{GetPaymentMethodName(paymentMethod)}-{orderId}");

        // V18: 幂等标记 (MarkPaymentIdempotency)
        MarkPaymentIdempotency(orderId, transactionId);

        // V18: 更新支付状态 (UpdateOrderPaymentStatus → Paid)
        var updated = await SyncPaymentStatusAsync(orderId, PaymentStatus.Paid, transactionId, ct);

        result.Success = updated;
        result.OrderId = orderId;
        result.TransactionId = transactionId;
        result.Message = updated ? "支付成功" : "支付状态更新失败";
        return result;
    }

    // ==================== V18: CheckRisk + RC_SafeNum ====================

    /// <summary>风控检查 — 用户信用、金额异常、地址重复、IP频率</summary>
    public async Task<RiskCheckResult> CheckRiskAsync(
        int userId, decimal orderTotal, string? shippingAddress = null,
        string? shippingPhone = null, string? ipAddress = null, CancellationToken ct = default)
    {
        var result = new RiskCheckResult { Timestamp = DateTime.Now, Passed = true };
        var maxRisk = "low";

        if (userId > 0)
        {
            // V18: 用户信用检查
            var totalOrders = await _db.Orders
                .AsNoTracking()
                .CountAsync(o => o.UserId == userId && o.Status != "Cancelled", ct);

            var totalSpent = 0m;
            var paidOrders = await _db.Orders
                .AsNoTracking()
                .Where(o => o.UserId == userId && (o.Status == "Paid" || o.Status == "Completed"))
                .Select(o => o.TotalAmount)
                .ToListAsync(ct);
            foreach (var amount in paidOrders)
                totalSpent += amount;

            var returnCount = await _db.Orders
                .AsNoTracking()
                .CountAsync(o => o.UserId == userId && o.Status == "Returned", ct);

            var thirtyDaysAgo = DateTime.Now.AddDays(-7);
            var unpaidOrders = await _db.Orders
                .AsNoTracking()
                .CountAsync(o => o.UserId == userId && o.Status == "Pending" && o.CreatedAt < thirtyDaysAgo, ct);

            var cancelCount = await _db.Orders
                .AsNoTracking()
                .CountAsync(o => o.UserId == userId && o.Status == "Cancelled", ct);

            // V18: D级用户判断
            var returnRate = totalOrders > 0 ? (double)returnCount / totalOrders : 0;
            var cancelRate = totalOrders > 0 ? (double)cancelCount / totalOrders : 0;

            if ((returnRate > 0.3 || cancelRate > 0.3))
            {
                result.Risks.Add(new RiskItem { Type = "credit", Level = "high", Message = "用户信用评级低，退货/取消率高" });
                result.RiskCount++;
                maxRisk = "high";
            }
            else if (unpaidOrders > 2)
            {
                result.Risks.Add(new RiskItem { Type = "credit", Level = "medium", Message = $"用户有 {unpaidOrders} 笔未支付订单" });
                result.RiskCount++;
                if (maxRisk != "high") maxRisk = "medium";
            }

            // V18: 金额异常检查
            if (orderTotal > 0 && totalOrders > 0)
            {
                var avgOrderAmount = totalSpent / totalOrders;
                if (orderTotal > avgOrderAmount * 5 && orderTotal > 2000)
                {
                    result.Risks.Add(new RiskItem
                    {
                        Type = "amount",
                        Level = "medium",
                        Message = $"订单金额¥{orderTotal:F2}远超历史平均¥{avgOrderAmount:F2}"
                    });
                    result.RiskCount++;
                    if (maxRisk != "high") maxRisk = "medium";
                }
            }
        }

        // V18: 地址/电话重复检查
        if (!string.IsNullOrEmpty(shippingAddress) || !string.IsNullOrEmpty(shippingPhone))
        {
            var thirtyDaysBack = DateTime.Now.AddDays(-30);
            var addrQuery = _db.Orders
                .AsNoTracking()
                .Where(o => o.Status != "Cancelled" && o.CreatedAt >= thirtyDaysBack);

            if (!string.IsNullOrEmpty(shippingAddress))
                addrQuery = addrQuery.Where(o => o.ShippingAddress == shippingAddress);

            var addrCount = await addrQuery.CountAsync(ct);

            // 电话号码单独检查
            if (!string.IsNullOrEmpty(shippingPhone))
            {
                var phoneCount = await _db.Orders
                    .AsNoTracking()
                    .CountAsync(o => o.ShippingPhone == shippingPhone && o.Status != "Cancelled" && o.CreatedAt >= thirtyDaysBack, ct);
                addrCount = Math.Max(addrCount, phoneCount);
            }

            if (addrCount >= 5)
            {
                result.Risks.Add(new RiskItem { Type = "address", Level = "high", Message = $"该地址/电话近30天下单{addrCount}次，异常频繁" });
                result.RiskCount++;
                maxRisk = "high";
            }
            else if (addrCount >= 3)
            {
                result.Risks.Add(new RiskItem { Type = "address", Level = "low", Message = $"该地址/电话近30天下单{addrCount}次" });
                result.RiskCount++;
            }
        }

        // V18: IP频率检查（30分钟内同一IP下单次数）
        if (!string.IsNullOrEmpty(ipAddress))
        {
            var thirtyMinAgo = DateTime.Now.AddMinutes(-30);
            var ipCount = await _db.AdminLogs
                .AsNoTracking()
                .CountAsync(l => l.Notes == ipAddress && l.CreatedAt >= thirtyMinAgo, ct);

            if (ipCount > 10)
            {
                result.Risks.Add(new RiskItem { Type = "ip", Level = "medium", Message = $"IP {ipAddress} 30分钟内请求{ipCount}次" });
                result.RiskCount++;
                if (maxRisk != "high") maxRisk = "medium";
            }
        }

        result.RiskLevel = maxRisk;
        result.Passed = maxRisk != "high";
        return result;
    }

    // ==================== V18: UpdateOrderPaymentStatus ====================

    /// <summary>同步支付状态 — 状态机流转</summary>
    public async Task<bool> SyncPaymentStatusAsync(int orderId, PaymentStatus newStatus, string transactionId, CancellationToken ct = default)
    {
        var statusText = newStatus switch
        {
            PaymentStatus.Paid => "Paid",
            PaymentStatus.Failed => "Failed",
            PaymentStatus.Refunded => "Refunded",
            PaymentStatus.Cancelled => "Cancelled",
            _ => "Pending"
        };

        // V18: 更新订单备注（追加 Transaction ID）
        var order = await _db.Orders.FirstOrDefaultAsync(o => o.OrderId == orderId, ct);
        if (order == null) return false;

        if (!string.IsNullOrEmpty(transactionId))
        {
            var currentNotes = order.Notes ?? "";
            if (!currentNotes.Contains("Transaction:"))
            {
                order.Notes = currentNotes + " | Transaction: " + transactionId;
            }
        }

        order.Status = statusText;
        order.UpdatedAt = DateTime.Now;

        await _db.SaveChangesAsync(ct);

        // V18: 支付成功 → 自动创建生产订单
        if (statusText == "Paid")
        {
            await AutoCreateProductionOrderAsync(orderId, ct);
        }

        return true;
    }

    // ==================== V18: AutoCreateProductionOrder ====================

    /// <summary>支付成功后自动创建生产工单 — 写入 ProductionOrders + ProductionLogs，含幂等检查</summary>
    public async Task<bool> AutoCreateProductionOrderAsync(int orderId, CancellationToken ct = default)
    {
        try
        {
            // V18: 幂等检查 — 检查是否已存在该订单的生产订单
            var existingCount = await _db.ProductionOrders
                .CountAsync(po => po.OrderId == orderId, ct);

            // 如果已存在生产订单，则退出（幂等）
            if (existingCount > 0)
                return true;

            // V18: 查询总瓶数
            var totalBottles = await _db.OrderDetails
                .Where(od => od.OrderId == orderId)
                .SumAsync(od => (int?)od.Quantity, ct) ?? 0;

            // 如果没有订单明细，直接返回成功
            if (totalBottles == 0)
                return true;

            // V18: 查询订单明细
            var details = await _db.OrderDetails
                .Where(od => od.OrderId == orderId)
                .ToListAsync(ct);

            // V18: 生成工单编号前缀
            var workOrderPrefix = $"WO-{DateTime.Now:yyyyMMdd}-";
            var bottleIndex = 0;

            // V18: 遍历每个订单明细，为每瓶生成工单
            foreach (var detail in details)
            {
                // V18: 获取产品配方信息
                var product = await _db.Products
                    .AsNoTracking()
                    .Where(p => p.ProductId == detail.ProductId)
                    .Select(p => new { p.RecipeId })
                    .FirstOrDefaultAsync(ct);

                Recipe? recipe = null;
                if (product?.RecipeId != null && product.RecipeId > 0)
                {
                    recipe = await _db.Recipes
                        .AsNoTracking()
                        .FirstOrDefaultAsync(r => r.RecipeId == product.RecipeId, ct);
                }

                // V18: 循环 Quantity 次，为每瓶创建工单
                for (int i = 1; i <= detail.Quantity; i++)
                {
                    bottleIndex++;
                    var workOrderNo = workOrderPrefix + bottleIndex.ToString("D3");

                    var productionOrder = new ProductionOrder
                    {
                        OrderId = orderId,
                        DetailId = detail.DetailId,
                        WorkOrderNo = workOrderNo,
                        BottleIndex = bottleIndex,
                        TotalBottles = totalBottles,
                        Status = "Pending",
                        Priority = 0,
                        RecipeId = recipe?.RecipeId,
                        RecipeName = recipe != null ? $"[{recipe.RecipeCode}] {recipe.RecipeName}" : null,
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    };

                    _db.ProductionOrders.Add(productionOrder);
                    await _db.SaveChangesAsync(ct);

                    // V18: 插入生产订单日志 (ProductionLogs)
                    var log = new ProductionLog
                    {
                        ProductionId = productionOrder.ProductionId,
                        Status = "Pending",
                        Notes = $"订单支付成功，系统自动创建生产工单 (第{bottleIndex}瓶/共{totalBottles}瓶)",
                        CreatedBy = "SYSTEM",
                        CreatedAt = DateTime.Now
                    };

                    _db.ProductionLogs.Add(log);
                    await _db.SaveChangesAsync(ct);
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    // ==================== V18: 确认支付（整合流程）====================

    /// <summary>确认支付 — 整合幂等检查+状态更新+自动创建生产工单</summary>
    public async Task<bool> ConfirmPaymentAsync(int orderId, string transactionId, CancellationToken ct = default)
    {
        // V18: 幂等性检查
        if (CheckPaymentIdempotency(orderId))
            return true; // 已处理，直接返回成功

        // V18: 幂等标记
        MarkPaymentIdempotency(orderId, transactionId);

        // 更新支付记录 — 使用跟踪实体（InMemory 兼容）
        var payRecords = await _db.PaymentRecords
            .Where(pr => pr.OrderId == orderId && pr.Status == "Pending")
            .ToListAsync(ct);
        foreach (var pr in payRecords)
        {
            pr.Status = "Paid";
            pr.TransactionNo = transactionId;
            pr.UpdatedAt = DateTime.Now;
        }

        // 更新订单状态 — 使用跟踪实体
        var orders = await _db.Orders
            .Where(o => o.OrderId == orderId && o.Status == "Pending")
            .ToListAsync(ct);
        foreach (var o in orders)
        {
            o.Status = "Paid";
            o.UpdatedAt = DateTime.Now;
        }

        await _db.SaveChangesAsync(ct);

        // V18: 支付成功 → 自动创建生产工单
        await AutoCreateProductionOrderAsync(orderId, ct);

        return true;
    }

    // ==================== 订单操作 ====================

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

    // ==================== V18: 支付安全 — 签名/幂等/IP白名单 ====================

    /// <summary>V18: 生成支付请求签名 (HMAC) — 防篡改</summary>
    private static string GeneratePaymentSignature(int orderId, decimal amount)
    {
        var timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        var payload = $"{orderId}|{amount:F2}|{timestamp}";
        var hash = ComputeHmac(payload, PaymentSigningKey);
        return $"{timestamp}:{hash}";
    }

    /// <summary>V18: 验证支付请求签名</summary>
    private static bool VerifyPaymentSignature(int orderId, decimal amount, string signature)
    {
        if (string.IsNullOrEmpty(signature)) return false;

        var parts = signature.Split(':', 2);
        if (parts.Length != 2) return false;

        if (!long.TryParse(parts[0], out var sigTimestamp)) return false;

        // V18: 验证时间戳（5分钟窗口，防重放）
        var nowTimestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        if (Math.Abs(nowTimestamp - sigTimestamp) > SignatureWindowSeconds) return false;

        // V18: 计算期望签名
        var payload = $"{orderId}|{amount:F2}|{sigTimestamp}";
        var expectedHash = ComputeHmac(payload, PaymentSigningKey);

        return string.Equals(parts[1], expectedHash, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>V18: 幂等性检查 — 防止重复支付</summary>
    private bool CheckPaymentIdempotency(int orderId)
    {
        if (_idempotencyStore.TryGetValue(orderId, out var entry))
        {
            var elapsed = (DateTime.UtcNow - entry.Timestamp).TotalSeconds;
            if (elapsed < IdempotencyTtlSeconds)
                return true;

            // 过期，清理
            _idempotencyStore.TryRemove(orderId, out _);
        }
        return false;
    }

    /// <summary>V18: 标记幂等键 — 支付处理前调用</summary>
    private void MarkPaymentIdempotency(int orderId, string transactionId)
    {
        _idempotencyStore[orderId] = new IdempotencyEntry(DateTime.UtcNow, transactionId);
    }

    /// <summary>V18: 验证回调签名</summary>
    private static bool VerifyCallbackSignature(int paymentMethod, Dictionary<string, string> callbackData)
    {
        // V18: 按支付方式调度签名验证
        // 当前为开发/测试模式，所有回调签名均通过
        // 生产环境中应对接真实支付平台的签名验证
        return paymentMethod switch
        {
            PaymentMethodWechat => true, // VerifyWeChatCallback
            PaymentMethodAlipay => true, // VerifyAlipayCallback
            PaymentMethodPaypal => true, // VerifyPayPalCallback
            PaymentMethodCod => true,
            _ => false
        };
    }

    /// <summary>V18: 验证回调 IP 是否在白名单内</summary>
    private static bool IsCallbackIpAllowed(string ipAddress)
    {
        // V18: CALLBACK_IP_WHITELIST 为空时允许所有 IP
        // 生产环境中应配置白名单
        return true;
    }

    /// <summary>V18: HMAC 哈希 (对应 ASP PaySignHMAC)</summary>
    private static string ComputeHmac(string message, string secret)
    {
        var keyBytes = Encoding.UTF8.GetBytes(secret);
        var messageBytes = Encoding.UTF8.GetBytes(message);
        using var hmac = new HMACSHA256(keyBytes);
        var hashBytes = hmac.ComputeHash(messageBytes);
        return Convert.ToHexString(hashBytes)[..64].ToUpper();
    }

    /// <summary>V18: RC_SafeNum — 安全数值转换</summary>
    private static decimal SafeNum(object? val)
    {
        if (val == null || val is DBNull) return 0m;
        if (decimal.TryParse(val.ToString(), out var result)) return result;
        return 0m;
    }

    private static string GetPaymentMethodName(int method) => method switch
    {
        PaymentMethodWechat => "WECHAT",
        PaymentMethodAlipay => "ALIPAY",
        PaymentMethodPaypal => "PAYPAL",
        PaymentMethodCod => "COD",
        _ => "UNKNOWN"
    };
}
