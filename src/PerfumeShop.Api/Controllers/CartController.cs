using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/cart")]
public class CartController : ControllerBase
{
    private readonly ICartService _svc;
    public CartController(ICartService svc) => _svc = svc;

    /// <summary>GET /api/cart/{userId} — 获取购物车</summary>
    [HttpGet("{userId}")]
    public async Task<IActionResult> GetCart(int userId)
    {
        var cart = await _svc.GetCartAsync(userId);
        return Ok(new { success = true, data = cart });
    }

    /// <summary>GET /api/cart/{userId}/count — 购物车商品数量</summary>
    [HttpGet("{userId}/count")]
    public async Task<IActionResult> GetCount(int userId)
    {
        var count = await _svc.GetCartCountAsync(userId);
        return Ok(new { success = true, count });
    }

    /// <summary>POST /api/cart/add — 添加商品到购物车</summary>
    [HttpPost("add")]
    public async Task<IActionResult> Add([FromBody] CartAddRequest req)
    {
        if (req.UserId <= 0) return BadRequest(new { success = false, message = "请先登录" });
        var qty = await _svc.AddItemAsync(req.UserId, req.ProductId, req.Quantity > 0 ? req.Quantity : 1, req.Size);
        return Ok(new { success = true, message = "已加入购物车", quantity = qty });
    }

    /// <summary>POST /api/cart/update — 更新购物车商品数量</summary>
    [HttpPost("update")]
    public async Task<IActionResult> Update([FromBody] CartUpdateRequest req)
    {
        var ok = await _svc.UpdateQuantityAsync(req.UserId, req.ProductId, req.Quantity);
        return Ok(new { success = ok, message = ok ? "已更新" : "商品不在购物车中" });
    }

    /// <summary>POST /api/cart/remove — 移除购物车商品</summary>
    [HttpPost("remove")]
    public async Task<IActionResult> Remove([FromBody] CartRemoveRequest req)
    {
        var ok = await _svc.RemoveItemAsync(req.UserId, req.ProductId);
        return Ok(new { success = ok, message = ok ? "已移除" : "商品不在购物车中" });
    }

    /// <summary>POST /api/cart/clear — 清空购物车</summary>
    [HttpPost("clear")]
    public async Task<IActionResult> Clear([FromBody] CartClearRequest req)
    {
        var ok = await _svc.ClearCartAsync(req.UserId);
        return Ok(new { success = ok, message = "购物车已清空" });
    }
}

public class CartAddRequest
{
    public int UserId { get; set; }
    public int ProductId { get; set; }
    public int Quantity { get; set; } = 1;
    public string? Size { get; set; }
}

public class CartUpdateRequest
{
    public int UserId { get; set; }
    public int ProductId { get; set; }
    public int Quantity { get; set; }
}

public class CartRemoveRequest
{
    public int UserId { get; set; }
    public int ProductId { get; set; }
}

public class CartClearRequest
{
    public int UserId { get; set; }
}
