using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 订单操作 API — V19 M3-A 扩展
/// 对齐 V18 user/orders.asp, user/order_detail.asp
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class OrdersController : ControllerBase
{
    private readonly IOrderRepository _orderRepo;
    private readonly IPaymentHandler _payment;
    private readonly IPromotionEngine _promo;
    private readonly ICostEngine _cost;
    private readonly PerfumeShopContext _db;

    public OrdersController(
        IOrderRepository orderRepo,
        IPaymentHandler payment,
        IPromotionEngine promo,
        ICostEngine cost,
        PerfumeShopContext db)
    {
        _orderRepo = orderRepo;
        _payment = payment;
        _promo = promo;
        _cost = cost;
        _db = db;
    }

    /// <summary>获取用户订单列表（分页+状态筛选）— 对齐 V18 user/orders.asp</summary>
    [HttpGet("user/{userId:int}")]
    public async Task<IActionResult> GetUserOrders(
        int userId,
        [FromQuery] int page = 1,
        [FromQuery] int pageSize = 10,
        [FromQuery] string? status = null)
    {
        IQueryable<Order> query = _db.Orders.Where(o => o.UserId == userId && o.Status != "Deleted");

        if (!string.IsNullOrEmpty(status))
            query = query.Where(o => o.Status == status);

        var total = await query.CountAsync();
        var items = await query
            .OrderByDescending(o => o.CreatedAt)
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .ToListAsync();

        return Ok(new
        {
            items = items.Select(o => new
            {
                o.OrderId,
                o.OrderNo,
                o.Status,
                o.TotalAmount,
                o.ShippingStatus,
                o.PaymentMethod,
                o.ShippingName,
                o.Notes,
                o.CreatedAt
            }),
            total,
            page,
            pageSize
        });
    }

    /// <summary>获取用户订单详情 — 对齐 V18 user/order_detail.asp</summary>
    [HttpGet("user/{userId:int}/detail/{orderId:int}")]
    public async Task<IActionResult> GetUserOrderDetail(int userId, int orderId)
    {
        var order = await _db.Orders
            .FirstOrDefaultAsync(o => o.OrderId == orderId && o.UserId == userId);

        if (order == null)
            return NotFound(new { message = "订单不存在" });

        // 获取订单商品明细
        var details = await _db.OrderDetails
            .Where(d => d.OrderId == orderId)
            .ToListAsync();

        // 获取香调配比 (仅定制产品)
        var noteSelections = await _db.OrderDetailNoteSelections
            .Where(s => details.Select(d => d.DetailId).Contains(s.DetailId))
            .Join(_db.FragranceNotes,
                s => s.NoteId,
                n => n.NoteId,
                (s, n) => new { s.DetailId, s.NoteType, NoteName = n.NoteName, s.Percentage })
            .ToListAsync();

        // 获取成分信息
        var ingredients = await _db.OrderIngredients
            .Where(i => i.OrderId == orderId)
            .ToListAsync();

        return Ok(new
        {
            order.OrderId,
            order.OrderNo,
            order.Status,
            order.TotalAmount,
            order.CostAmount,
            order.ProfitAmount,
            order.ShippingFee,
            order.ShippingStatus,
            order.ShippingCompany,
            order.TrackingNumber,
            order.PaymentMethod,
            order.ShippingName,
            order.ShippingAddress,
            order.ShippingCity,
            order.ShippingPhone,
            order.Notes,
            order.CouponCode,
            order.CouponDiscount,
            order.PointsRedeemed,
            order.PointsDiscount,
            order.CreatedAt,
            order.UpdatedAt,
            order.ShippedAt,
            order.DeliveredAt,
            details = details.Select(d => new
            {
                d.DetailId,
                d.ProductId,
                d.ProductName,
                d.Quantity,
                d.UnitPrice,
                d.Subtotal,
                d.VolumeName,
                d.VolumeMl,
                d.BottleName,
                d.CustomLabel,
                d.TopNoteName,
                d.MiddleNoteName,
                d.BaseNoteName,
                noteSelections = noteSelections
                    .Where(ns => ns.DetailId == d.DetailId)
                    .Select(ns => new { ns.NoteType, ns.NoteName, ns.Percentage }),
                ingredients = ingredients
                    .Where(i => i.DetailId == d.DetailId)
                    .Select(i => new { i.IngredientName })
            })
        });
    }

    /// <summary>删除订单（仅已完成/已取消可删）— 对齐 V18 delete_order.asp</summary>
    [HttpDelete("{id:int}")]
    public async Task<IActionResult> DeleteOrder(int id, [FromQuery] int userId)
    {
        var order = await _orderRepo.GetByIdAsync(id);
        if (order == null)
            return NotFound(new { message = "订单不存在" });

        if (order.UserId != userId)
            return Forbid();

        if (order.Status != "Completed" && order.Status != "Cancelled" && order.Status != "Refunded")
            return BadRequest(new { message = "只有已完成、已取消或已退款的订单可以删除" });

        // 软删除
        order.Status = "Deleted";
        _orderRepo.Update(order);
        await _orderRepo.SaveChangesAsync();

        return Ok(new { message = "订单已删除" });
    }

    /// <summary>获取订单详情（管理员用）</summary>
    [HttpGet("{id:int}")]
    public async Task<IActionResult> GetOrder(int id)
    {
        var order = await _orderRepo.GetByIdAsync(id);
        if (order == null)
            return NotFound(new { message = "订单不存在" });

        return Ok(new
        {
            order.OrderId,
            order.OrderNo,
            order.Status,
            order.TotalAmount,
            order.CostAmount,
            order.ProfitAmount,
            order.ShippingFee,
            order.ShippingStatus,
            order.ShippingCompany,
            order.TrackingNumber,
            order.PaymentMethod,
            order.ShippingName,
            order.ShippingAddress,
            order.ShippingCity,
            order.ShippingPhone,
            order.CreatedAt,
            order.UpdatedAt
        });
    }

    /// <summary>创建订单 (简化版)</summary>
    [HttpPost]
    public async Task<IActionResult> CreateOrder([FromBody] CreateOrderRequest request)
    {
        if (request.UserId <= 0 || request.Items == null || request.Items.Count == 0)
            return BadRequest(new { message = "无效的订单请求" });

        var orderNo = $"ORD{DateTime.Now:yyyyMMddHHmmss}{new Random().Next(1000, 9999)}";
        var order = new Order
        {
            OrderNo = orderNo,
            UserId = request.UserId,
            Status = "Pending",
            TotalAmount = request.Items.Sum(i => i.Price * i.Quantity),
            ShippingAddress = request.ShippingAddress,
            ShippingCity = request.ShippingCity,
            ShippingName = request.ShippingName,
            ShippingPhone = request.ShippingPhone,
            PaymentMethod = request.PaymentMethod ?? "online",
            CreatedAt = DateTime.Now,
            UpdatedAt = DateTime.Now
        };

        await _orderRepo.AddAsync(order);
        await _orderRepo.SaveChangesAsync();

        return Ok(new { orderId = order.OrderId, orderNo = order.OrderNo, message = "订单创建成功" });
    }

    /// <summary>取消订单</summary>
    [HttpPost("{id:int}/cancel")]
    public async Task<IActionResult> CancelOrder(int id, [FromBody] CancelRequest request)
    {
        var result = await _payment.CancelOrderAsync(id, request.UserId);
        return result
            ? Ok(new { message = "订单已取消" })
            : BadRequest(new { message = "取消失败，订单状态不允许取消" });
    }

    /// <summary>确认支付</summary>
    [HttpPost("{id:int}/pay")]
    public async Task<IActionResult> ConfirmPayment(int id, [FromBody] PayRequest request)
    {
        await _payment.ConfirmPaymentAsync(id, request.TransactionId);
        return Ok(new { message = "支付确认成功" });
    }

    /// <summary>确认收货</summary>
    [HttpPost("{id:int}/deliver")]
    public async Task<IActionResult> ConfirmDelivery(int id, [FromBody] CancelRequest request)
    {
        var result = await _payment.ConfirmDeliveryAsync(id, request.UserId);
        return result
            ? Ok(new { message = "确认收货成功" })
            : BadRequest(new { message = "确认收货失败" });
    }

    /// <summary>申请退款</summary>
    [HttpPost("{id:int}/refund")]
    public async Task<IActionResult> RequestRefund(int id, [FromBody] RefundRequest request)
    {
        var result = await _payment.RequestRefundAsync(id, request.UserId, request.Amount, request.Reason);
        return result
            ? Ok(new { message = "退款申请已提交" })
            : BadRequest(new { message = "退款申请失败" });
    }
}

public class CreateOrderRequest
{
    public int UserId { get; set; }
    public List<OrderItemRequest>? Items { get; set; }
    public string? ShippingAddress { get; set; }
    public string? ShippingCity { get; set; }
    public string? ShippingName { get; set; }
    public string? ShippingPhone { get; set; }
    public string? PaymentMethod { get; set; }
}

public class OrderItemRequest
{
    public int ProductId { get; set; }
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class CancelRequest
{
    public int UserId { get; set; }
}

public class PayRequest
{
    public string TransactionId { get; set; } = "";
}

public class RefundRequest
{
    public int UserId { get; set; }
    public decimal Amount { get; set; }
    public string Reason { get; set; } = "";
}
