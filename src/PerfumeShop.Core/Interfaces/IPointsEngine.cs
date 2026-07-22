namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 积分与奖励引擎接口 — 对齐 V18 points_engine.asp (20.5KB)
/// 功能: 积分获取/消费/兑换/过期/签到/规则缓存/余额同步
/// V19 改进: 三表同步写入 (PointsLedger + UserPoints + Users.Points) 使用 EF Core Transaction 包裹
/// </summary>
public interface IPointsEngine
{
    // ==================== 规则缓存 ====================

    /// <summary>获取积分规则值（带缓存） — 对应 PE_GetRule()</summary>
    Task<decimal> GetRuleAsync(string ruleCode, CancellationToken ct = default);

    /// <summary>刷新规则缓存 — 对应 PE_GetRuleCache() 初始化</summary>
    Task RefreshRuleCacheAsync(CancellationToken ct = default);

    // ==================== 核心操作 ====================

    /// <summary>获取积分 — 对应 PE_EarnPoints()</summary>
    Task<bool> EarnAsync(int userId, int points, string source, string? description = null, int? referenceId = null, CancellationToken ct = default);

    /// <summary>消费积分 — 对应 PE_RedeemPoints()</summary>
    Task<bool> RedeemAsync(int userId, int points, string redemptionType, int? referenceId = null, CancellationToken ct = default);

    /// <summary>获取用户可用积分（汇总过期过滤） — 对应 PE_GetAvailablePoints()</summary>
    Task<int> GetBalanceAsync(int userId, CancellationToken ct = default);

    /// <summary>处理过期积分 — 对应 PE_ExpireOutdatedPoints()</summary>
    Task ApplyExpirationAsync(int userId, CancellationToken ct = default);

    /// <summary>同步更新余额（三表一致性写入） — 对应 PE_UpdateBalance()</summary>
    Task UpdateBalanceAsync(int userId, CancellationToken ct = default);

    // ==================== 积分计算 ====================

    /// <summary>计算积分的货币价值（元） — 对应 PE_CalcPointsValue()</summary>
    Task<decimal> CalcPointsValueAsync(int points, CancellationToken ct = default);

    /// <summary>计算消费应得积分 — 对应 PE_CalcOrderPoints()</summary>
    Task<int> CalcOrderPointsAsync(decimal orderAmount, CancellationToken ct = default);

    /// <summary>获取最大抵扣百分比 — 对应 PE_GetMaxRedeemPct()</summary>
    Task<decimal> GetMaxRedeemPctAsync(CancellationToken ct = default);

    /// <summary>计算订单可抵扣的最大积分 — 对应 PE_CalcMaxRedeemablePoints()</summary>
    Task<int> CalcMaxRedeemablePointsAsync(int userId, decimal orderAmount, CancellationToken ct = default);

    // ==================== 签到 ====================

    /// <summary>签到检查 — 对应 PE_CheckSignIn()</summary>
    Task<bool> CheckSignInAsync(int userId, CancellationToken ct = default);

    /// <summary>执行签到 — 对应 PE_DoSignIn()</summary>
    Task<int> DoSignInAsync(int userId, CancellationToken ct = default);

    // ==================== 账本与汇总 ====================

    /// <summary>获取积分账本（分页） — 对应 PE_GetPointsLedger()</summary>
    Task<PointsLedgerResult> GetLedgerAsync(int userId, int page = 1, int pageSize = 20, CancellationToken ct = default);

    /// <summary>获取积分汇总（仪表盘用） — 对应 PE_GetPointsSummary()</summary>
    Task<PointsSummaryDto> GetPointsSummaryAsync(int userId, CancellationToken ct = default);

    /// <summary>获取订单中已获/已用的积分 — 对应 PE_GetOrderPoints()</summary>
    Task<OrderPointsDto> GetOrderPointsAsync(int orderId, CancellationToken ct = default);

    // ==================== 兑换商城 ====================

    /// <summary>积分兑换处理 — 对应 PE_DoRedeem()</summary>
    Task<string> DoRedeemAsync(int userId, int redemptionId, CancellationToken ct = default);
}

// ==================== DTOs ====================

/// <summary>积分账本分页结果</summary>
public class PointsLedgerResult
{
    public List<PointsLedgerItem> Items { get; set; } = new();
    public int Total { get; set; }
}

/// <summary>积分账本条目</summary>
public class PointsLedgerItem
{
    public int LedgerId { get; set; }
    public int UserId { get; set; }
    public int Points { get; set; }
    public string PointType { get; set; } = "";
    public string Source { get; set; } = "";
    public int? ReferenceId { get; set; }
    public string? Description { get; set; }
    public DateTime? ExpiresAt { get; set; }
    public bool IsExpired { get; set; }
    public DateTime CreatedAt { get; set; }
}

/// <summary>积分汇总 DTO — 对应 PE_GetPointsSummary() 返回的 Dictionary</summary>
public class PointsSummaryDto
{
    public int Available { get; set; }
    public int TotalEarned { get; set; }
    public int TotalRedeemed { get; set; }
    public int TodayEarned { get; set; }
    public int ExpiringSoon { get; set; }
}

/// <summary>订单积分 DTO — 对应 PE_GetOrderPoints() 返回的 Dictionary</summary>
public class OrderPointsDto
{
    public int Earned { get; set; }
    public int Redeemed { get; set; }
    public decimal Discount { get; set; }
}
