using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Interfaces;

/// <summary>
/// 秒杀活动服务接口 — 替代 flash_sale.asp 前端 + admin/operation/flash_sale.asp
/// </summary>
public interface IFlashSaleService
{
    /// <summary>获取当前进行中的秒杀列表(含商品详情), 支持分页</summary>
    Task<(List<FlashSaleDto> Items, int Total)> GetActiveFlashSalesAsync(int page = 1, int pageSize = 12);

    /// <summary>获取即将开始的秒杀活动</summary>
    Task<List<FlashSaleDto>> GetUpcomingFlashSalesAsync(int top = 6);

    /// <summary>获取单个秒杀详情</summary>
    Task<FlashSaleDto?> GetFlashSaleByIdAsync(int flashSaleId);

    /// <summary>抢购(扣减库存+记录)</summary>
    Task<FlashSalePurchaseResult> PurchaseAsync(int flashSaleId, int userId, int quantity);

    /// <summary>管理后台: 创建/更新秒杀活动</summary>
    Task<int> SaveFlashSaleAsync(FlashSale entity);

    /// <summary>管理后台: 删除秒杀活动</summary>
    Task<bool> DeleteFlashSaleAsync(int flashSaleId);

    /// <summary>管理后台: 切换启用状态</summary>
    Task<bool> ToggleActiveAsync(int flashSaleId);

    /// <summary>管理后台: 统计数据</summary>
    Task<FlashSaleAdminStats> GetAdminStatsAsync();
}

public class FlashSaleAdminStats
{
    public int Total { get; set; }
    public int Active { get; set; }
    public int Upcoming { get; set; }
    public int Expired { get; set; }
}

/// <summary>
/// 拼团服务接口 — 替代 group_buy.asp 前端 + admin/operation/group_buy.asp
/// </summary>
public interface IGroupBuyService
{
    /// <summary>获取进行中的拼团计划(含商品信息)</summary>
    Task<List<GroupBuyPlanDto>> GetActivePlansAsync();

    /// <summary>获取某计划下可加入的团</summary>
    Task<List<OpenGroupDto>> GetOpenGroupsAsync(int planId);

    /// <summary>发起新团</summary>
    Task<GroupBuyStartResult> StartGroupAsync(int planId, int userId);

    /// <summary>加入已有团</summary>
    Task<GroupBuyJoinResult> JoinGroupAsync(int groupId, int userId);

    /// <summary>获取团的详情(参与者列表)</summary>
    Task<GroupDetailDto?> GetGroupDetailAsync(int groupId);

    /// <summary>获取计划统计</summary>
    Task<GroupBuyStats> GetPlanStatsAsync(int planId);

    /// <summary>管理后台: 创建/更新拼团计划</summary>
    Task<int> SavePlanAsync(GroupBuyPlan entity);

    /// <summary>管理后台: 删除拼团计划</summary>
    Task<bool> DeletePlanAsync(int planId);

    /// <summary>管理后台: 切换启用状态</summary>
    Task<bool> TogglePlanActiveAsync(int planId);
}

/// <summary>
/// 订阅服务接口 — 替代 subscribe.asp + admin/operation/subscription_plans.asp
/// </summary>
public interface ISubscriptionService
{
    /// <summary>获取所有活跃订阅计划</summary>
    Task<List<SubscriptionPlan>> GetActivePlansAsync();

    /// <summary>获取全部订阅计划(含不活跃)</summary>
    Task<List<SubscriptionPlan>> GetAllPlansAsync();

    /// <summary>获取用户当前订阅</summary>
    Task<UserSubscriptionDto?> GetUserSubscriptionAsync(int userId);

    /// <summary>创建订阅</summary>
    Task<SubscribeResult> SubscribeAsync(int userId, int planId, bool autoRenew = true);

    /// <summary>取消订阅</summary>
    Task<bool> CancelSubscriptionAsync(int subscriptionId, int userId);

    /// <summary>暂停/恢复订阅</summary>
    Task<bool> ToggleAutoRenewAsync(int subscriptionId, int userId, bool autoRenew);

    /// <summary>获取用户配送历史</summary>
    Task<List<SubscriptionDelivery>> GetDeliveryHistoryAsync(int subscriptionId);

    /// <summary>管理后台: 保存订阅计划</summary>
    Task<int> SavePlanAsync(SubscriptionPlan entity);

    /// <summary>管理后台: 删除订阅计划</summary>
    Task<bool> DeletePlanAsync(int planId);

    /// <summary>管理后台: 切换启用状态</summary>
    Task<bool> TogglePlanActiveAsync(int planId);
}

/// <summary>
/// 积分服务接口 — 替代 admin/operation/points.asp + 积分引擎
/// </summary>
public interface IPointsService
{
    /// <summary>获取用户积分余额</summary>
    Task<PointsBalanceDto> GetBalanceAsync(int userId);

    /// <summary>获取积分流水(分页)</summary>
    Task<(List<PointsLedger> Items, int Total)> GetLedgerAsync(int userId, int page = 1, int pageSize = 20);

    /// <summary>发放积分(下单/签到/活动等)</summary>
    Task<int> AwardPointsAsync(int userId, int points, string source, string? description = null, int? referenceId = null);

    /// <summary>扣减积分(兑换/抵扣)</summary>
    Task<bool> DeductPointsAsync(int userId, int points, string source, string? description = null, int? referenceId = null);

    /// <summary>获取积分兑换商城商品列表</summary>
    Task<List<PointsRedemption>> GetRedemptionItemsAsync();

    /// <summary>积分兑换</summary>
    Task<RedeemResult> RedeemAsync(int userId, int redemptionId);

    /// <summary>获取积分规则</summary>
    Task<List<PointsRule>> GetRulesAsync();
}

/// <summary>
/// 购物车服务接口 — 替代 api/cart_*.asp 系列
/// </summary>
public interface ICartService
{
    /// <summary>获取用户购物车汇总</summary>
    Task<CartSummary> GetCartAsync(int userId);

    /// <summary>添加商品到购物车</summary>
    Task<int> AddItemAsync(int userId, int productId, int quantity, string? size = null);

    /// <summary>更新购物车商品数量</summary>
    Task<bool> UpdateQuantityAsync(int userId, int productId, int quantity);

    /// <summary>移除购物车商品</summary>
    Task<bool> RemoveItemAsync(int userId, int productId);

    /// <summary>清空购物车</summary>
    Task<bool> ClearCartAsync(int userId);

    /// <summary>获取购物车商品数量</summary>
    Task<int> GetCartCountAsync(int userId);
}
