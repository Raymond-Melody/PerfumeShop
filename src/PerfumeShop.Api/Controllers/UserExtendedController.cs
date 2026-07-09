using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/user")]
public class UserExtendedController : ControllerBase
{
    private readonly PerfumeShopContext _db;
    public UserExtendedController(PerfumeShopContext db) => _db = db;

    // ========== 地址管理 ==========

    /// <summary>GET /api/user/{userId}/addresses — 获取用户地址列表</summary>
    [HttpGet("{userId}/addresses")]
    public async Task<IActionResult> GetAddresses(int userId)
    {
        // UserAddress is keyless
        var addresses = await _db.UserAddresses
            .Where(a => a.UserId == userId)
            .OrderByDescending(a => a.IsDefault)
            .ThenByDescending(a => a.CreatedAt)
            .ToListAsync();
        return Ok(new { success = true, data = addresses });
    }

    /// <summary>POST /api/user/{userId}/addresses — 新增地址</summary>
    [HttpPost("{userId}/addresses")]
    public async Task<IActionResult> AddAddress(int userId, [FromBody] AddressRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Address) || string.IsNullOrWhiteSpace(req.Consignee) || string.IsNullOrWhiteSpace(req.Phone))
            return BadRequest(new { success = false, message = "请填写完整地址信息" });

        var addr = new UserAddress
        {
            UserId = userId,
            Consignee = req.Consignee,
            Phone = req.Phone,
            Address = req.Address,
            Province = req.Province,
            City = req.City,
            District = req.District,
            IsDefault = req.IsDefault,
            CreatedAt = DateTime.Now
        };
        _db.UserAddresses.Add(addr);
        await _db.SaveChangesAsync();

        // 如果设为默认，取消其他默认
        if (req.IsDefault == true)
        {
            var others = await _db.UserAddresses
                .Where(a => a.UserId == userId && a.AddressId != addr.AddressId && a.IsDefault == true)
                .ToListAsync();
            foreach (var o in others) o.IsDefault = false;
            await _db.SaveChangesAsync();
        }

        return Ok(new { success = true, message = "地址已添加" });
    }

    /// <summary>DELETE /api/user/{userId}/addresses/{addressId} — 删除地址</summary>
    [HttpDelete("{userId}/addresses/{addressId}")]
    public async Task<IActionResult> DeleteAddress(int userId, int addressId)
    {
        var addr = await _db.UserAddresses
            .FirstOrDefaultAsync(a => a.UserId == userId && a.AddressId == addressId);
        if (addr == null) return NotFound(new { success = false, message = "地址不存在" });

        _db.UserAddresses.Remove(addr);
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已删除" });
    }

    // ========== 收藏管理 ==========

    /// <summary>GET /api/user/{userId}/favorites — 获取收藏列表</summary>
    [HttpGet("{userId}/favorites")]
    public async Task<IActionResult> GetFavorites(int userId)
    {
        // UserFavorite is keyless
        var favorites = await (from f in _db.UserFavorites
                               join p in _db.Products on f.ProductId equals p.ProductId
                               where f.UserId == userId
                               orderby f.CreatedTime descending
                               select new
                               {
                                   favoriteId = f.FavoriteId,
                                   productId = f.ProductId,
                                   name = p.ProductName,
                                   price = p.BasePrice,
                                   image = p.ImageUrl,
                                   category = p.Category,
                                   createdTime = f.CreatedTime
                               }).ToListAsync();
        return Ok(new { success = true, data = favorites });
    }

    /// <summary>POST /api/user/{userId}/favorites — 添加收藏</summary>
    [HttpPost("{userId}/favorites")]
    public async Task<IActionResult> AddFavorite(int userId, [FromBody] FavoriteRequest req)
    {
        // 检查重复
        var exists = await _db.UserFavorites
            .AnyAsync(f => f.UserId == userId && f.ProductId == req.ProductId);
        if (exists) return Ok(new { success = false, message = "已在收藏中" });

        _db.UserFavorites.Add(new UserFavorite
        {
            UserId = userId,
            ProductId = req.ProductId,
            CreatedTime = DateTime.Now
        });
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已收藏" });
    }

    /// <summary>DELETE /api/user/{userId}/favorites/{productId} — 取消收藏</summary>
    [HttpDelete("{userId}/favorites/{productId}")]
    public async Task<IActionResult> RemoveFavorite(int userId, int productId)
    {
        var fav = await _db.UserFavorites
            .FirstOrDefaultAsync(f => f.UserId == userId && f.ProductId == productId);
        if (fav == null) return NotFound(new { success = false, message = "收藏不存在" });

        _db.UserFavorites.Remove(fav);
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已取消收藏" });
    }

    /// <summary>GET /api/user/{userId}/favorites/check/{productId} — 检查是否已收藏</summary>
    [HttpGet("{userId}/favorites/check/{productId}")]
    public async Task<IActionResult> CheckFavorite(int userId, int productId)
    {
        var exists = await _db.UserFavorites.AnyAsync(f => f.UserId == userId && f.ProductId == productId);
        return Ok(new { success = true, isFavorite = exists });
    }
}

public class AddressRequest
{
    public string Consignee { get; set; } = "";
    public string Phone { get; set; } = "";
    public string Address { get; set; } = "";
    public string? Province { get; set; }
    public string? City { get; set; }
    public string? District { get; set; }
    public bool? IsDefault { get; set; }
}

public class FavoriteRequest
{
    public int ProductId { get; set; }
}
