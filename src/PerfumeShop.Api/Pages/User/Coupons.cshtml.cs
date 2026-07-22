using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class CouponsModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public CouponsModel(PerfumeShopContext db) => _db = db;

    public string CurrentTab { get; set; } = "available";
    public int AvailableCount { get; set; }
    public int UsedCount { get; set; }
    public int ExpiredCount { get; set; }
    public List<CouponDisplayItem> DisplayItems { get; set; } = new();

    public class CouponDisplayItem
    {
        public int UserCouponId { get; set; }
        public string CouponCode { get; set; } = "";
        public string CouponName { get; set; } = "";
        public string CouponType { get; set; } = "";
        public decimal DiscountValue { get; set; }
        public decimal MinSpend { get; set; }
        public decimal MaxDiscount { get; set; }
        public string? Description { get; set; }
        public string? Terms { get; set; }
        public string? ApplicableCategory { get; set; }
        public DateTime? ValidTo { get; set; }
        public string Status { get; set; } = "";
        public string Source { get; set; } = "";
    }

    public async Task OnGetAsync([FromQuery] string tab = "available")
    {
        var userId = 1;
        CurrentTab = tab;
        var now = DateTime.Now;

        // 获取所有用户优惠券（含 Coupon 导航属性）
        var allCoupons = await (from uc in _db.UserCoupons
                                join c in _db.Coupons on uc.CouponId equals c.CouponId
                                where uc.UserId == userId
                                select new CouponDisplayItem
                                {
                                    UserCouponId = uc.UserCouponId,
                                    CouponCode = uc.CouponCode,
                                    CouponName = c.CouponName,
                                    CouponType = c.CouponType,
                                    DiscountValue = c.DiscountValue ?? 0,
                                    MinSpend = c.MinSpend,
                                    MaxDiscount = c.MaxDiscount,
                                    Description = c.Description,
                                    Terms = c.Terms,
                                    ApplicableCategory = c.ApplicableCategory,
                                    ValidTo = c.ValidTo,
                                    Status = uc.Status,
                                    Source = uc.Source
                                }).ToListAsync();

        AvailableCount = allCoupons.Count(c => c.Status == "available" && c.ValidTo >= now);
        UsedCount = allCoupons.Count(c => c.Status == "used");
        ExpiredCount = allCoupons.Count(c => c.Status == "expired" || (c.Status == "available" && c.ValidTo < now));

        DisplayItems = tab switch
        {
            "used" => allCoupons.Where(c => c.Status == "used").ToList(),
            "expired" => allCoupons.Where(c => c.Status == "expired" || (c.Status == "available" && c.ValidTo < now)).ToList(),
            _ => allCoupons.Where(c => c.Status == "available" && c.ValidTo >= now).ToList()
        };
    }

    public static string GetTypeLabel(string type) => type?.ToLower() switch
    {
        "fixed" => "满减券",
        "percentage" => "折扣券",
        "free_shipping" => "免邮券",
        "gift" => "礼品券",
        _ => "优惠券"
    };

    public static string GetValueText(string type, decimal value) => type?.ToLower() switch
    {
        "fixed" => $"¥{value:F0}",
        "percentage" => $"{value:F0}折",
        "free_shipping" => "免邮",
        "gift" => "礼品",
        _ => "优惠"
    };

    public static string GetTypeColor(string type) => type?.ToLower() switch
    {
        "fixed" => "#FF5722",
        "percentage" => "#9C27B0",
        "free_shipping" => "#2196F3",
        "gift" => "#FF9800",
        _ => "#607D8B"
    };

    public static string GetTypeIcon(string type) => type?.ToLower() switch
    {
        "fixed" => "yen-sign",
        "percentage" => "percent",
        "free_shipping" => "truck",
        "gift" => "gift",
        _ => "tag"
    };
}
