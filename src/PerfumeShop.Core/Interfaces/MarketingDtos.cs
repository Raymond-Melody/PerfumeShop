namespace PerfumeShop.Core.Interfaces;

// ========== 秒杀 DTOs ==========

public class FlashSaleDto
{
    public int FlashSaleId { get; set; }
    public int ProductId { get; set; }
    public string ProductName { get; set; } = "";
    public string? ImageUrl { get; set; }
    public decimal FlashPrice { get; set; }
    public decimal BasePrice { get; set; }
    public int Stock { get; set; }
    public int SoldCount { get; set; }
    public int Remaining => Stock - SoldCount;
    public int LimitPerUser { get; set; }
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public string? Category { get; set; }
    public string? Description { get; set; }
    public decimal DiscountPercent => BasePrice > 0 ? Math.Round((1 - FlashPrice / BasePrice) * 100, 1) : 0;
}

public class FlashSalePurchaseResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public int? OrderId { get; set; }
}

// ========== 拼团 DTOs ==========

public class GroupBuyPlanDto
{
    public int PlanId { get; set; }
    public int ProductId { get; set; }
    public string ProductName { get; set; } = "";
    public string? ImageUrl { get; set; }
    public decimal GroupPrice { get; set; }
    public decimal BasePrice { get; set; }
    public int TeamSize { get; set; }
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public int DurationHours { get; set; }
    public string? Description { get; set; }
    public int OpenGroupCount { get; set; }
    public int SuccessGroupCount { get; set; }
}

public class OpenGroupDto
{
    public int GroupId { get; set; }
    public string GroupSn { get; set; } = "";
    public int CurrentSize { get; set; }
    public int TargetSize { get; set; }
    public string InitiatorName { get; set; } = "";
    public DateTime CreatedAt { get; set; }
    public int HoursPassed { get; set; }
    public int SlotsRemaining => TargetSize - CurrentSize;
}

public class GroupBuyStartResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public int? GroupId { get; set; }
    public string? GroupSn { get; set; }
}

public class GroupBuyJoinResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public bool IsGroupComplete { get; set; }
}

public class GroupDetailDto
{
    public int GroupId { get; set; }
    public string GroupSn { get; set; } = "";
    public int PlanId { get; set; }
    public string ProductName { get; set; } = "";
    public int CurrentSize { get; set; }
    public int TargetSize { get; set; }
    public int Status { get; set; }
    public DateTime CreatedAt { get; set; }
    public DateTime? CompletedAt { get; set; }
    public List<GroupParticipantDto> Participants { get; set; } = new();
}

public class GroupParticipantDto
{
    public int ParticipantId { get; set; }
    public int UserId { get; set; }
    public string Username { get; set; } = "";
    public bool IsInitiator { get; set; }
    public int Status { get; set; }
    public DateTime JoinedAt { get; set; }
}

public class GroupBuyStats
{
    public int OpenGroups { get; set; }
    public int SuccessGroups { get; set; }
    public int TotalParticipants { get; set; }
}

// ========== 订阅 DTOs ==========

public class UserSubscriptionDto
{
    public int SubscriptionId { get; set; }
    public int PlanId { get; set; }
    public string PlanName { get; set; } = "";
    public string Period { get; set; } = "";
    public decimal Price { get; set; }
    public string Status { get; set; } = "";
    public DateTime StartDate { get; set; }
    public DateTime? EndDate { get; set; }
    public bool AutoRenew { get; set; }
    public DateTime CreatedAt { get; set; }
    public int DeliveryCount { get; set; }
}

public class SubscribeResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public int? SubscriptionId { get; set; }
}

// ========== 积分 DTOs ==========

public class PointsBalanceDto
{
    public int UserId { get; set; }
    public int AvailablePoints { get; set; }
    public int TotalPoints { get; set; }
    public int UsedPoints { get; set; }
    public int ExpiredPoints { get; set; }
    public DateTime? LastUpdatedAt { get; set; }
}

public class RedeemResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public int? PointsRemaining { get; set; }
}

// ========== 购物车 DTOs ==========

public class CartItem
{
    public int ProductId { get; set; }
    public string ProductName { get; set; } = "";
    public string? ImageUrl { get; set; }
    public decimal UnitPrice { get; set; }
    public int Quantity { get; set; }
    public string? ProductType { get; set; }
    public string? Size { get; set; }
    public decimal Subtotal => UnitPrice * Quantity;
}

public class CartSummary
{
    public int UserId { get; set; }
    public List<CartItem> Items { get; set; } = new();
    public int ItemCount => Items.Sum(i => i.Quantity);
    public decimal Subtotal => Items.Sum(i => i.Subtotal);
    public decimal Discount { get; set; }
    public decimal ShippingFee { get; set; }
    public decimal Total => Subtotal - Discount + ShippingFee;
    public string? AppliedCoupon { get; set; }
}

// ========== 搜索 DTOs ==========

public class SearchSuggestion
{
    public string Text { get; set; } = "";
    public string? Type { get; set; }  // "product", "category", "brand"
    public int? ProductId { get; set; }
    public string? ImageUrl { get; set; }
}
