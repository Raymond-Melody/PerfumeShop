using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// V20 费用分摊引擎 — 从 V18 admin/finance/expense_allocation.asp 完整移植
/// 三种分摊：运费（按规则自动计算）、平台扣点（支付编码匹配）、推广费（GMV两级归因）
/// 特性：幂等（先删后插）、账期取订单创建月、尾差末行归一、日期区间含结束当天
/// </summary>
public class ExpenseAllocationService
{
    private readonly PerfumeShopContext _db;

    public ExpenseAllocationService(PerfumeShopContext db) => _db = db;

    public record AllocationResult(bool Success, string Message, int OrdersProcessed, int RecordsCreated);

    // ==================== 配置读取 ====================

    private async Task<decimal> GetConfigDecimalAsync(string key, decimal defaultValue)
    {
        var val = await _db.SiteSettings.AsNoTracking()
            .Where(s => s.SettingKey == key).Select(s => s.SettingValue).FirstOrDefaultAsync();
        return decimal.TryParse(val, out var d) ? d : defaultValue;
    }

    /// <summary>读取运费规则配置（与 V18 SiteSettings 同键共享）</summary>
    public async Task<ShippingConfig> GetShippingConfigAsync() => new(
        FirstWeight: await GetConfigDecimalAsync("ShippingFirstWeight", 1m),
        FirstPrice: await GetConfigDecimalAsync("ShippingFirstPrice", 10m),
        ContinueWeight: await GetConfigDecimalAsync("ShippingContinueWeight", 1m),
        ContinuePrice: await GetConfigDecimalAsync("ShippingContinuePrice", 5m),
        VolumeFactor: await GetConfigDecimalAsync("ShippingVolumeFactor", 5000m),
        DefaultUnitWeight: await GetConfigDecimalAsync("ShippingDefaultUnitWeight", 0.5m),
        DefaultUnitVolume: await GetConfigDecimalAsync("ShippingDefaultUnitVolume", 750m));

    /// <summary>读取平台费率配置（百分比值，与 V18 同键共享）</summary>
    public async Task<PlatformFeeConfig> GetPlatformFeeConfigAsync() => new(
        Alipay: await GetConfigDecimalAsync("PlatformFeeAlipay", 0.6m) / 100m,
        Wechat: await GetConfigDecimalAsync("PlatformFeeWechat", 0.6m) / 100m,
        Stripe: await GetConfigDecimalAsync("PlatformFeeStripe", 2.9m) / 100m,
        PayPal: await GetConfigDecimalAsync("PlatformFeePayPal", 4.4m) / 100m,
        UnionPay: await GetConfigDecimalAsync("PlatformFeeUnionPay", 0.6m) / 100m,
        FixedFee: await GetConfigDecimalAsync("PlatformFixedFee", 0m));

    public record ShippingConfig(decimal FirstWeight, decimal FirstPrice, decimal ContinueWeight,
        decimal ContinuePrice, decimal VolumeFactor, decimal DefaultUnitWeight, decimal DefaultUnitVolume);

    public record PlatformFeeConfig(decimal Alipay, decimal Wechat, decimal Stripe,
        decimal PayPal, decimal UnionPay, decimal FixedFee);

    // ==================== 公共辅助 ====================

    private IQueryable<Order> PaidOrdersInRange(int? orderId, DateTime? start, DateTime? end)
    {
        var q = _db.Orders.AsNoTracking().Where(o => o.Status == "Paid");
        if (orderId.HasValue) q = q.Where(o => o.OrderId == orderId.Value);
        if (start.HasValue) q = q.Where(o => o.CreatedAt >= start.Value.Date);
        // 对齐 V18 修复：含结束当天 (CreatedAt < endDate + 1day)
        if (end.HasValue) q = q.Where(o => o.CreatedAt < end.Value.Date.AddDays(1));
        return q;
    }

    private static string PeriodOf(DateTime? createdAt) =>
        (createdAt ?? DateTime.Now).ToString("yyyy-MM");

    // ==================== 1. 运费分摊 ====================

    /// <summary>
    /// 运费分摊：按首重/续重/体积重规则为每笔订单自动计算运费，再按方式分摊到SKU
    /// method: weight | volume | equal（缺数据自动回退按数量平均）
    /// </summary>
    public async Task<AllocationResult> AllocateShippingAsync(
        int? orderId, DateTime? startDate, DateTime? endDate, string method)
    {
        var cfg = await GetShippingConfigAsync();
        var orders = await PaidOrdersInRange(orderId, startDate, endDate)
            .Select(o => new { o.OrderId, o.CreatedAt }).ToListAsync();

        int ordersProcessed = 0, recordsCreated = 0;

        foreach (var order in orders)
        {
            // SKU 明细（INNER JOIN Products 取重量/体积）
            var items = await _db.OrderDetails.AsNoTracking()
                .Where(d => d.OrderId == order.OrderId)
                .Join(_db.Products.AsNoTracking(), d => d.ProductId, p => p.ProductId,
                    (d, p) => new { d.ProductId, d.Quantity, p.Weight, p.Volume })
                .ToListAsync();
            if (items.Count == 0) continue;

            // 第一遍：累加计费重量（每SKU取 max(实重, 体积重)）与汇总
            decimal chargeableWeight = 0, totalWeight = 0, totalVolume = 0;
            int totalQty = 0;
            foreach (var it in items)
            {
                var qty = it.Quantity <= 0 ? 1 : it.Quantity;
                var w = (it.Weight ?? 0) <= 0 ? cfg.DefaultUnitWeight : it.Weight!.Value;
                var v = (it.Volume ?? 0) <= 0 ? cfg.DefaultUnitVolume : it.Volume!.Value;
                var itemWeight = w * qty;
                var itemVolume = v * qty / cfg.VolumeFactor;
                chargeableWeight += Math.Max(itemWeight, itemVolume);
                totalWeight += w * qty;
                totalVolume += v * qty;
                totalQty += qty;
            }
            if (chargeableWeight <= 0) continue;

            // 按规则计算订单运费：首重 + Ceil(超出/续重)*续重价
            decimal orderFreight;
            if (chargeableWeight <= cfg.FirstWeight)
                orderFreight = cfg.FirstPrice;
            else
            {
                var extraUnits = cfg.ContinueWeight > 0
                    ? (int)Math.Ceiling((chargeableWeight - cfg.FirstWeight) / cfg.ContinueWeight)
                    : 0;
                orderFreight = cfg.FirstPrice + extraUnits * cfg.ContinuePrice;
            }
            if (orderFreight <= 0) continue;

            // 幂等：先清除该订单旧运费分摊
            await _db.ExpenseRecords
                .Where(e => e.OrderId == order.OrderId && e.ExpenseType == "Shipping")
                .ExecuteDeleteAsync();

            // 第二遍：分摊到SKU（末行承担尾差）
            var period = PeriodOf(order.CreatedAt);
            decimal remaining = orderFreight;
            for (int i = 0; i < items.Count; i++)
            {
                var it = items[i];
                var qty = it.Quantity <= 0 ? 1 : it.Quantity;
                var w = (it.Weight ?? 0) <= 0 ? cfg.DefaultUnitWeight : it.Weight!.Value;
                var v = (it.Volume ?? 0) <= 0 ? cfg.DefaultUnitVolume : it.Volume!.Value;

                decimal alloc = method switch
                {
                    "weight" when totalWeight > 0 => orderFreight * (w * qty) / totalWeight,
                    "volume" when totalVolume > 0 => orderFreight * (v * qty) / totalVolume,
                    _ => orderFreight / totalQty * qty  // equal/回退：按数量
                };

                if (i == items.Count - 1) alloc = remaining;           // 尾差归一
                else { alloc = Math.Round(alloc, 2); remaining -= alloc; }

                if (alloc <= 0) continue;
                _db.ExpenseRecords.Add(new ExpenseRecord
                {
                    OrderId = order.OrderId, ProductId = it.ProductId,
                    ExpenseType = "Shipping", ExpenseName = "运费分摊",
                    Amount = alloc, AllocationMethod = method,
                    AllocationRatio = (double)Math.Round(alloc / orderFreight, 4),
                    Period = period, CreatedAt = DateTime.Now
                });
                recordsCreated++;
            }
            ordersProcessed++;
        }

        await _db.SaveChangesAsync();
        await SyncOrderExpenseProfitAsync(); // V21: 回写费用并重算利润
        return new(true, $"运费分摊完成：处理 {ordersProcessed} 笔订单，生成 {recordsCreated} 条记录", ordersProcessed, recordsCreated);
    }

    // ==================== 2. 平台扣点分摊 ====================

    /// <summary>
    /// 平台扣点：按数字支付编码匹配费率（1=微信/2=支付宝/3=PayPal/4=货到付款=0），
    /// 兼容英文子串旧数据；按SKU金额比例分摊。
    /// </summary>
    public async Task<AllocationResult> AllocatePlatformFeeAsync(DateTime? startDate, DateTime? endDate)
    {
        var cfg = await GetPlatformFeeConfigAsync();
        var orders = await PaidOrdersInRange(null, startDate, endDate)
            .Select(o => new { o.OrderId, o.TotalAmount, o.PaymentMethod, o.CreatedAt }).ToListAsync();

        int ordersProcessed = 0, recordsCreated = 0;

        foreach (var order in orders)
        {
            var rate = ResolvePlatformRate(order.PaymentMethod, cfg);
            var feeAmount = order.TotalAmount * rate + cfg.FixedFee;
            if (feeAmount <= 0) continue;

            var items = await _db.OrderDetails.AsNoTracking()
                .Where(d => d.OrderId == order.OrderId)
                .Join(_db.Products.AsNoTracking(), d => d.ProductId, p => p.ProductId,
                    (d, p) => new { d.ProductId, d.Subtotal })
                .ToListAsync();
            if (items.Count == 0) continue;

            // 幂等
            await _db.ExpenseRecords
                .Where(e => e.OrderId == order.OrderId && e.ExpenseType == "PlatformFee")
                .ExecuteDeleteAsync();

            var period = PeriodOf(order.CreatedAt);
            decimal remaining = feeAmount;
            for (int i = 0; i < items.Count; i++)
            {
                var it = items[i];
                decimal alloc = order.TotalAmount > 0
                    ? feeAmount * it.Subtotal / order.TotalAmount
                    : feeAmount / items.Count;

                if (i == items.Count - 1) alloc = remaining;
                else { alloc = Math.Round(alloc, 2); remaining -= alloc; }

                if (alloc <= 0) continue;
                _db.ExpenseRecords.Add(new ExpenseRecord
                {
                    OrderId = order.OrderId, ProductId = it.ProductId,
                    ExpenseType = "PlatformFee", ExpenseName = "平台扣点",
                    Amount = alloc, AllocationMethod = "PaymentMethod",
                    AllocationRatio = (double)Math.Round(rate, 4),
                    Period = period, CreatedAt = DateTime.Now
                });
                recordsCreated++;
            }
            ordersProcessed++;
        }

        await _db.SaveChangesAsync();
        await SyncOrderExpenseProfitAsync(); // V21: 回写费用并重算利润
        return new(true, $"平台费用分摊完成：处理 {ordersProcessed} 笔订单，生成 {recordsCreated} 条记录", ordersProcessed, recordsCreated);
    }

    /// <summary>支付方式→费率解析（对齐 V18 与 payment_config.asp 编码）</summary>
    private static decimal ResolvePlatformRate(string? paymentMethod, PlatformFeeConfig cfg)
    {
        var pm = (paymentMethod ?? "").Trim();
        return pm switch
        {
            "1" => cfg.Wechat,     // 微信支付
            "2" => cfg.Alipay,     // 支付宝
            "3" => cfg.PayPal,     // PayPal
            "4" => 0m,             // 货到付款（默认无平台扣点）
            _ => MatchLegacy(pm.ToLowerInvariant(), cfg)  // 旧英文子串兼容
        };
        static decimal MatchLegacy(string pm, PlatformFeeConfig cfg) =>
            pm.Contains("alipay") ? cfg.Alipay :
            pm.Contains("wechat") ? cfg.Wechat :
            pm.Contains("stripe") ? cfg.Stripe :
            pm.Contains("paypal") ? cfg.PayPal :
            pm.Contains("union") ? cfg.UnionPay : 0m;
    }

    // ==================== 3. 推广费分摊 ====================

    /// <summary>
    /// 推广费GMV归因：单笔承担 = 总消耗/GMV × 订单金额；订单内再按SKU销售额占比二级分摊。
    /// </summary>
    public async Task<AllocationResult> AllocatePromotionAsync(
        DateTime? startDate, DateTime? endDate, decimal totalPromoAmount, decimal gmvAmount)
    {
        if (totalPromoAmount <= 0) return new(false, "推广费用必须大于0", 0, 0);
        if (gmvAmount <= 0) return new(false, "有效成交额必须大于0", 0, 0);

        var orders = await PaidOrdersInRange(null, startDate, endDate)
            .Select(o => new { o.OrderId, o.TotalAmount, o.CreatedAt }).ToListAsync();
        if (orders.Count == 0) return new(true, "所选区间无已支付订单", 0, 0);

        int ordersProcessed = 0, recordsCreated = 0;
        decimal remainingTotal = totalPromoAmount;

        for (int oi = 0; oi < orders.Count; oi++)
        {
            var order = orders[oi];
            decimal alloc = totalPromoAmount * order.TotalAmount / gmvAmount;

            if (oi == orders.Count - 1) alloc = remainingTotal;       // 订单级尾差
            else { alloc = Math.Round(alloc, 2); remainingTotal -= alloc; }
            if (alloc <= 0) continue;

            var items = await _db.OrderDetails.AsNoTracking()
                .Where(d => d.OrderId == order.OrderId)
                .Join(_db.Products.AsNoTracking(), d => d.ProductId, p => p.ProductId,
                    (d, p) => new { d.ProductId, d.Subtotal })
                .ToListAsync();
            if (items.Count == 0) continue;

            // 幂等
            await _db.ExpenseRecords
                .Where(e => e.OrderId == order.OrderId && e.ExpenseType == "Promotion")
                .ExecuteDeleteAsync();

            var period = PeriodOf(order.CreatedAt);
            decimal remainingSku = alloc;
            for (int i = 0; i < items.Count; i++)
            {
                var it = items[i];
                decimal skuAlloc = order.TotalAmount > 0
                    ? alloc * it.Subtotal / order.TotalAmount
                    : alloc / items.Count;

                if (i == items.Count - 1) skuAlloc = remainingSku;     // SKU级尾差
                else { skuAlloc = Math.Round(skuAlloc, 2); remainingSku -= skuAlloc; }

                if (skuAlloc <= 0) continue;
                _db.ExpenseRecords.Add(new ExpenseRecord
                {
                    OrderId = order.OrderId, ProductId = it.ProductId,
                    ExpenseType = "Promotion", ExpenseName = "推广费分摊",
                    Amount = skuAlloc, AllocationMethod = "GMVRatio",
                    AllocationRatio = (double)Math.Round(it.Subtotal / gmvAmount, 6),
                    SourceOrderId = 0,
                    Period = period, CreatedAt = DateTime.Now
                });
                recordsCreated++;
            }
            ordersProcessed++;
        }

        await _db.SaveChangesAsync();
        await SyncOrderExpenseProfitAsync(); // V21: 回写费用并重算利润
        return new(true, $"推广费分摊完成：处理 {ordersProcessed} 笔订单，生成 {recordsCreated} 条记录", ordersProcessed, recordsCreated);
    }

    // ==================== V21: 分摊后回写订单费用并重算利润 ====================

    /// <summary>
    /// 将 ExpenseRecords 按订单汇总写入 Orders.ExpenseAmount，并令
    /// ProfitAmount = TotalAmount - CostAmount - ShippingFee - ExpenseAmount（下限 0）。
    /// 集合式、幂等，与成本引擎口径一致 — 对标 V18 expense_allocation.asp SyncAllOrderExpenseProfit。
    /// </summary>
    public async Task SyncOrderExpenseProfitAsync()
    {
        // 回写各订单费用合计
        await _db.Database.ExecuteSqlRawAsync(@"
UPDATE o SET o.ExpenseAmount = COALESCE(e.ExpSum,0)
FROM Orders o
INNER JOIN (SELECT OrderID, SUM(Amount) AS ExpSum FROM ExpenseRecords WHERE OrderID IS NOT NULL GROUP BY OrderID) e
    ON o.OrderID = e.OrderID");

        // 重算利润（已含费用）
        await _db.Database.ExecuteSqlRawAsync(@"
UPDATE Orders SET ProfitAmount =
    CASE WHEN (COALESCE(TotalAmount,0) - COALESCE(CostAmount,0) - COALESCE(ShippingFee,0) - COALESCE(ExpenseAmount,0)) < 0
         THEN 0
         ELSE (COALESCE(TotalAmount,0) - COALESCE(CostAmount,0) - COALESCE(ShippingFee,0) - COALESCE(ExpenseAmount,0)) END
WHERE COALESCE(ExpenseAmount,0) > 0");
    }

    // ==================== 配置保存 ====================

    /// <summary>保存配置项（UPSERT，与 V18 SaveConfig 一致）</summary>
    public async Task SaveConfigAsync(Dictionary<string, string> settings)
    {
        foreach (var (key, value) in settings)
        {
            var existing = await _db.SiteSettings.FirstOrDefaultAsync(s => s.SettingKey == key);
            if (existing != null) existing.SettingValue = value;
            else _db.SiteSettings.Add(new SiteSetting { SettingKey = key, SettingValue = value });
        }
        await _db.SaveChangesAsync();
    }
}
