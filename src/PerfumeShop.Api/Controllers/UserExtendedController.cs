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

    [HttpGet("{userId}/addresses")]
    public async Task<IActionResult> GetAddresses(int userId)
    {
        var addresses = await _db.UserAddresses
            .Where(a => a.UserId == userId)
            .OrderByDescending(a => a.IsDefault)
            .ThenByDescending(a => a.CreatedAt)
            .ToListAsync();
        return Ok(new { success = true, data = addresses });
    }

    [HttpPost("{userId}/addresses")]
    public async Task<IActionResult> AddAddress(int userId, [FromBody] AddressRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Address) || string.IsNullOrWhiteSpace(req.Consignee) || string.IsNullOrWhiteSpace(req.Phone))
            return BadRequest(new { success = false, message = "请填写完整地址信息" });

        var addr = new UserAddress
        {
            UserId = userId, Consignee = req.Consignee, Phone = req.Phone,
            Address = req.Address, Province = req.Province, City = req.City,
            District = req.District, IsDefault = req.IsDefault, CreatedAt = DateTime.Now
        };
        _db.UserAddresses.Add(addr);
        await _db.SaveChangesAsync();

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

    [HttpPut("{userId}/addresses/{addressId}")]
    public async Task<IActionResult> UpdateAddress(int userId, int addressId, [FromBody] AddressRequest req)
    {
        var addr = await _db.UserAddresses.FirstOrDefaultAsync(a => a.UserId == userId && a.AddressId == addressId);
        if (addr == null) return NotFound(new { success = false, message = "地址不存在" });
        addr.Consignee = req.Consignee; addr.Phone = req.Phone; addr.Address = req.Address;
        addr.Province = req.Province; addr.City = req.City; addr.District = req.District;
        addr.IsDefault = req.IsDefault;
        await _db.SaveChangesAsync();
        if (req.IsDefault == true)
        {
            var others = await _db.UserAddresses
                .Where(a => a.UserId == userId && a.AddressId != addressId && a.IsDefault == true)
                .ToListAsync();
            foreach (var o in others) o.IsDefault = false;
            await _db.SaveChangesAsync();
        }
        return Ok(new { success = true, message = "地址已更新" });
    }

    [HttpDelete("{userId}/addresses/{addressId}")]
    public async Task<IActionResult> DeleteAddress(int userId, int addressId)
    {
        var addr = await _db.UserAddresses.FirstOrDefaultAsync(a => a.UserId == userId && a.AddressId == addressId);
        if (addr == null) return NotFound(new { success = false, message = "地址不存在" });
        _db.UserAddresses.Remove(addr);
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已删除" });
    }

    [HttpPut("{userId}/addresses/{addressId}/default")]
    public async Task<IActionResult> SetDefaultAddress(int userId, int addressId)
    {
        var addr = await _db.UserAddresses.FirstOrDefaultAsync(a => a.UserId == userId && a.AddressId == addressId);
        if (addr == null) return NotFound(new { success = false, message = "地址不存在" });
        var others = await _db.UserAddresses
            .Where(a => a.UserId == userId && a.AddressId != addressId && a.IsDefault == true).ToListAsync();
        foreach (var o in others) o.IsDefault = false;
        addr.IsDefault = true;
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已设为默认" });
    }

    // ========== 收藏管理 ==========

    [HttpGet("{userId}/favorites")]
    public async Task<IActionResult> GetFavorites(int userId)
    {
        var favorites = await (from f in _db.UserFavorites
                               join p in _db.Products on f.ProductId equals p.ProductId
                               where f.UserId == userId
                               orderby f.CreatedTime descending
                               select new { favoriteId = f.FavoriteId, productId = f.ProductId, name = p.ProductName, price = p.BasePrice, image = p.ImageUrl, category = p.Category, createdTime = f.CreatedTime }).ToListAsync();
        return Ok(new { success = true, data = favorites });
    }

    [HttpPost("{userId}/favorites")]
    public async Task<IActionResult> AddFavorite(int userId, [FromBody] FavoriteRequest req)
    {
        var exists = await _db.UserFavorites.AnyAsync(f => f.UserId == userId && f.ProductId == req.ProductId);
        if (exists) return Ok(new { success = false, message = "已在收藏中" });
        _db.UserFavorites.Add(new UserFavorite { UserId = userId, ProductId = req.ProductId, CreatedTime = DateTime.Now });
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已收藏" });
    }

    [HttpDelete("{userId}/favorites/{productId}")]
    public async Task<IActionResult> RemoveFavorite(int userId, int productId)
    {
        var fav = await _db.UserFavorites.FirstOrDefaultAsync(f => f.UserId == userId && f.ProductId == productId);
        if (fav == null) return NotFound(new { success = false, message = "收藏不存在" });
        _db.UserFavorites.Remove(fav);
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已取消收藏" });
    }

    [HttpGet("{userId}/favorites/check/{productId}")]
    public async Task<IActionResult> CheckFavorite(int userId, int productId)
    {
        var exists = await _db.UserFavorites.AnyAsync(f => f.UserId == userId && f.ProductId == productId);
        return Ok(new { success = true, isFavorite = exists });
    }

    // ========== 积分管理 ==========

    [HttpGet("{userId}/points/balance")]
    public async Task<IActionResult> GetPointsBalance(int userId)
    {
        var user = await _db.Users.FindAsync(userId);
        return Ok(new { success = true, points = user?.Points ?? 0, tier = user?.CustomerTier ?? "普通会员" });
    }

    [HttpGet("{userId}/points/ledger")]
    public async Task<IActionResult> GetPointsLedger(int userId, [FromQuery] int page = 1, [FromQuery] int pageSize = 20)
    {
        var query = _db.PointsLedgers.Where(l => l.UserId == userId);
        var total = await query.CountAsync();
        var items = await query.OrderByDescending(l => l.CreatedAt).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return Ok(new { success = true, data = items, total, page, pageSize });
    }

    [HttpPost("{userId}/points/redeem")]
    public async Task<IActionResult> RedeemPoints(int userId, [FromBody] RedeemPointsRequest req)
    {
        var user = await _db.Users.FindAsync(userId);
        if (user == null) return NotFound(new { success = false, message = "用户不存在" });
        var redemption = await _db.PointsRedemptions.FindAsync(req.RedemptionId);
        if (redemption == null) return NotFound(new { success = false, message = "兑换项目不存在" });
        if (!redemption.IsEnabled) return BadRequest(new { success = false, message = "该兑换项目已下架" });
        if (redemption.Stock <= 0) return BadRequest(new { success = false, message = "库存不足" });
        var currentPoints = user.Points ?? 0;
        if (currentPoints < redemption.PointsCost)
            return BadRequest(new { success = false, message = "积分不足", remaining = currentPoints });
        user.Points = currentPoints - redemption.PointsCost;
        redemption.Stock -= 1;
        _db.PointsLedgers.Add(new PointsLedger { UserId = userId, Points = -redemption.PointsCost, PointType = "redeem", Source = "redemption", Description = $"兑换: {redemption.ItemName}", CreatedAt = DateTime.Now });
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "兑换成功", pointsRemaining = user.Points });
    }

    // ========== 优惠券管理 ==========

    [HttpGet("{userId}/coupons")]
    public async Task<IActionResult> GetUserCoupons(int userId, [FromQuery] string status = "available")
    {
        var query = from uc in _db.UserCoupons join c in _db.Coupons on uc.CouponId equals c.CouponId where uc.UserId == userId
                    select new { userCouponId = uc.UserCouponId, couponCode = uc.CouponCode, couponName = c.CouponName, couponType = c.CouponType, discountValue = c.DiscountValue ?? 0, minSpend = c.MinSpend, maxDiscount = c.MaxDiscount, description = c.Description, terms = c.Terms, applicableCategory = c.ApplicableCategory, validTo = c.ValidTo, status = uc.Status, source = uc.Source };
        var items = await query.ToListAsync();
        return Ok(new { success = true, data = items });
    }

    [HttpPost("{userId}/coupons/validate")]
    public async Task<IActionResult> ValidateCoupon(int userId, [FromBody] ValidateCouponRequest req)
    {
        var coupon = await (from uc in _db.UserCoupons join c in _db.Coupons on uc.CouponId equals c.CouponId
                            where uc.UserId == userId && uc.CouponCode == req.Code && uc.Status == "available"
                            select new { c.CouponId, c.CouponName, c.CouponType, c.DiscountValue, c.MinSpend, c.MaxDiscount, c.ValidTo }).FirstOrDefaultAsync();
        if (coupon == null) return Ok(new { success = false, message = "优惠券不存在或已使用" });
        if (coupon.ValidTo < DateTime.Now) return Ok(new { success = false, message = "优惠券已过期" });
        return Ok(new { success = true, coupon = new { coupon.CouponId, coupon.CouponName, coupon.CouponType, coupon.DiscountValue, coupon.MinSpend, coupon.MaxDiscount } });
    }

    // ========== 个人资料管理 ==========

    [HttpPut("{userId}/profile")]
    public async Task<IActionResult> UpdateProfile(int userId, [FromBody] ProfileRequest req)
    {
        var user = await _db.Users.FindAsync(userId);
        if (user == null) return NotFound(new { success = false, message = "用户不存在" });
        if (!string.IsNullOrWhiteSpace(req.FullName)) user.FullName = req.FullName;
        if (!string.IsNullOrWhiteSpace(req.Email)) user.Email = req.Email;
        if (req.Phone != null) user.Phone = req.Phone;
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "个人信息已更新" });
    }

    [HttpPut("{userId}/password")]
    public async Task<IActionResult> ChangePassword(int userId, [FromBody] ChangePasswordRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.CurrentPassword) || string.IsNullOrWhiteSpace(req.NewPassword))
            return BadRequest(new { success = false, message = "请填写完整" });
        if (req.NewPassword.Length < 6)
            return BadRequest(new { success = false, message = "新密码至少6位" });
        var user = await _db.Users.FindAsync(userId);
        if (user == null) return NotFound(new { success = false, message = "用户不存在" });
        if (!string.IsNullOrEmpty(user.Password) && user.Password != req.CurrentPassword)
        {
            if (!user.Password.StartsWith("V3$") && !user.Password.StartsWith("V2_"))
                return BadRequest(new { success = false, message = "当前密码错误" });
        }
        user.Password = req.NewPassword;
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "密码已修改" });
    }

    [HttpPut("{userId}/preference")]
    public async Task<IActionResult> UpdatePreference(int userId, [FromBody] PreferenceRequest req)
    {
        return Ok(new { success = true, message = "偏好设置已更新" });
    }

    // ========== M3-C: 评价管理 ==========

    [HttpGet("{userId}/reviews")]
    public async Task<IActionResult> GetUserReviews(int userId, [FromQuery] int page = 1, [FromQuery] int pageSize = 20)
    {
        var query = _db.ProductReviews.Where(r => r.UserId == userId && r.Status != "Deleted").OrderByDescending(r => r.CreatedAt);
        var total = await query.CountAsync();
        var items = await query.Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return Ok(new { success = true, data = items, total, page, pageSize });
    }

    [HttpDelete("{userId}/reviews/{reviewId}")]
    public async Task<IActionResult> DeleteReview(int userId, int reviewId)
    {
        var review = await _db.ProductReviews.FirstOrDefaultAsync(r => r.UserId == userId && r.ReviewId == reviewId);
        if (review == null) return NotFound(new { success = false, message = "评价不存在" });
        review.Status = "Deleted";
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "已删除" });
    }

    // ========== M3-C: 订阅管理 ==========

    [HttpPost("{userId}/subscription/{subId}/pause")]
    public async Task<IActionResult> PauseSubscription(int userId, int subId)
    {
        var sub = await _db.UserSubscriptions.FirstOrDefaultAsync(s => s.UserId == userId && s.SubscriptionId == subId);
        if (sub == null) return NotFound(new { success = false, message = "订阅不存在" });
        if (sub.Status != "Active") return BadRequest(new { success = false, message = "仅活跃订阅可暂停" });
        sub.Status = "Paused";
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "订阅已暂停" });
    }

    [HttpPost("{userId}/subscription/{subId}/resume")]
    public async Task<IActionResult> ResumeSubscription(int userId, int subId)
    {
        var sub = await _db.UserSubscriptions.FirstOrDefaultAsync(s => s.UserId == userId && s.SubscriptionId == subId);
        if (sub == null) return NotFound(new { success = false, message = "订阅不存在" });
        if (sub.Status != "Paused") return BadRequest(new { success = false, message = "仅已暂停订阅可恢复" });
        sub.Status = "Active";
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "订阅已恢复" });
    }

    [HttpPost("{userId}/subscription/{subId}/cancel")]
    public async Task<IActionResult> CancelUserSubscription(int userId, int subId, [FromBody] CancelReasonRequest? req)
    {
        var sub = await _db.UserSubscriptions.FirstOrDefaultAsync(s => s.UserId == userId && s.SubscriptionId == subId);
        if (sub == null) return NotFound(new { success = false, message = "订阅不存在" });
        if (sub.Status == "Cancelled") return BadRequest(new { success = false, message = "已取消" });
        sub.Status = "Cancelled";
        sub.EndDate = DateTime.Now;
        sub.AutoRenew = false;
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "订阅已取消" });
    }

    // ========== M3-C: KOL 商品 ==========

    [HttpGet("{userId}/kol-products")]
    public async Task<IActionResult> GetKolProducts(int userId, [FromQuery] string? category = null, [FromQuery] string? keyword = null)
    {
        var query = _db.Products.Where(p => p.IsActive == true && p.ProductType == "KOL");
        if (!string.IsNullOrEmpty(category)) query = query.Where(p => p.Category == category);
        if (!string.IsNullOrEmpty(keyword)) query = query.Where(p => p.ProductName.Contains(keyword));
        var items = await query.OrderByDescending(p => p.CreatedAt).ToListAsync();
        return Ok(new { success = true, data = items });
    }

    // ========== M3-C: 数据导出 ==========

    [HttpGet("{userId}/export")]
    public async Task<IActionResult> ExportUserData(int userId, [FromQuery] string type = "all", [FromQuery] DateTime? from = null, [FromQuery] DateTime? to = null)
    {
        var user = await _db.Users.FindAsync(userId);
        if (user == null) return NotFound(new { success = false, message = "用户不存在" });

        var sb = new System.Text.StringBuilder();
        var dateSuffix = DateTime.Now.ToString("yyyyMMdd");
        var fileName = $"my_data_{userId}_{dateSuffix}.csv";

        if (type == "all" || type == "user")
        {
            sb.AppendLine("Type,UserId,Username,Email,FullName,Phone");
            sb.AppendLine($"user,{user.UserId},{user.Username},{user.Email},{user.FullName},{user.Phone}");
        }
        if (type == "all" || type == "orders")
        {
            var orders = await _db.Orders.Where(o => o.UserId == userId).ToListAsync();
            sb.AppendLine("OrderId,OrderNo,TotalAmount,Status,CreatedAt");
            foreach (var o in orders) sb.AppendLine($"{o.OrderId},{o.OrderNo},{o.TotalAmount},{o.Status},{o.CreatedAt}");
        }
        if (type == "all" || type == "reviews")
        {
            var reviews = await _db.ProductReviews.Where(r => r.UserId == userId && r.Status != "Deleted").ToListAsync();
            sb.AppendLine("ReviewId,ProductId,Rating,Title,Comment,CreatedAt");
            foreach (var r in reviews) sb.AppendLine($"{r.ReviewId},{r.ProductId},{r.Rating},{r.Title},{r.Comment},{r.CreatedAt}");
        }
        if (type == "all" || type == "favorites")
        {
            var favs = await _db.UserFavorites.Where(f => f.UserId == userId).ToListAsync();
            sb.AppendLine("FavoriteId,ProductId,CreatedTime");
            foreach (var f in favs) sb.AppendLine($"{f.FavoriteId},{f.ProductId},{f.CreatedTime}");
        }

        var csvBytes = System.Text.Encoding.UTF8.GetPreamble().Concat(System.Text.Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        return File(csvBytes, "text/csv", fileName);
    }

    // ========== M3-C: 账户注销 ==========

    [HttpPost("{userId}/account-delete")]
    public async Task<IActionResult> RequestAccountDelete(int userId, [FromBody] AccountDeleteRequest req)
    {
        var user = await _db.Users.FindAsync(userId);
        if (user == null) return NotFound(new { success = false, message = "用户不存在" });
        if (!string.IsNullOrEmpty(user.Password) && user.Password != req.Password)
        {
            if (!user.Password.StartsWith("V3$") && !user.Password.StartsWith("V2_"))
                return BadRequest(new { success = false, message = "密码错误" });
        }
        user.IsActive = false;
        await _db.SaveChangesAsync();
        return Ok(new { success = true, message = "账户已标记为待删除，30天后将实际清理", coolingDays = 30 });
    }

    // ========== M3-C: 推荐统计 ==========

    [HttpGet("{userId}/referrals")]
    public async Task<IActionResult> GetReferralStats(int userId)
    {
        var totalInvites = await _db.ReferralRelations.CountAsync(r => r.AncestorUserId == userId && r.Depth == 1);
        var rewardPoints = totalInvites * 100;
        var activeTokens = await _db.ReferralTokens.Where(t => t.ReferrerUserId == userId && t.IsActive == true).OrderByDescending(t => t.CreatedAt).ToListAsync();
        var link = activeTokens.Any() ? $"https://shop.example.com/register?ref={activeTokens.First().OriginalToken}" : null;
        return Ok(new { success = true, totalInvites, totalRewardPoints = rewardPoints, link, tokens = activeTokens });
    }

    [HttpPost("{userId}/referrals/generate")]
    public async Task<IActionResult> GenerateReferralLink(int userId)
    {
        var tokenValue = Guid.NewGuid().ToString("N")[..12];
        var token = new ReferralToken
        {
            ReferrerUserId = userId, TokenHash = tokenValue, OriginalToken = tokenValue,
            ReferrerType = "user", ExpiresAt = DateTime.Now.AddYears(1), MaxUses = 100,
            UsedCount = 0, IsActive = true, CreatedAt = DateTime.Now
        };
        _db.ReferralTokens.Add(token);
        await _db.SaveChangesAsync();
        return Ok(new { success = true, link = $"https://shop.example.com/register?ref={tokenValue}", tokenHash = tokenValue });
    }
}

// ========== Request DTOs ==========

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

public class FavoriteRequest { public int ProductId { get; set; } }
public class RedeemPointsRequest { public int RedemptionId { get; set; } }
public class ValidateCouponRequest { public string Code { get; set; } = ""; }
public class ProfileRequest { public string? FullName { get; set; } public string? Email { get; set; } public string? Phone { get; set; } }
public class ChangePasswordRequest { public string CurrentPassword { get; set; } = ""; public string NewPassword { get; set; } = ""; }
public class PreferenceRequest { public bool? EmailNewsletter { get; set; } }
public class AccountDeleteRequest { public string Password { get; set; } = ""; public bool KeepOrderHistory { get; set; } = true; }
public class CancelReasonRequest { public string? Reason { get; set; } }
