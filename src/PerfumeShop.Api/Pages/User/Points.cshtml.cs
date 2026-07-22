using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class PointsModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public PointsModel(PerfumeShopContext db) => _db = db;

    public int AvailablePoints { get; set; }
    public int TotalEarned { get; set; }
    public int TotalRedeemed { get; set; }
    public int ExpiringSoon { get; set; }
    public string? MemberTier { get; set; }
    public List<PointsLedger> LedgerItems { get; set; } = new();
    public List<PointsRedemption> RedemptionItems { get; set; } = new();
    public List<PointsRule> Rules { get; set; } = new();
    public int CurrentPage { get; set; } = 1;
    public int TotalPages { get; set; } = 1;

    public async Task OnGetAsync([FromQuery] int page = 1)
    {
        // TODO: 从 AuthService 获取真实 userId
        var userId = 1;
        const int pageSize = 15;
        CurrentPage = Math.Max(1, page);

        // 积分余额
        var user = await _db.Users.FindAsync(userId);
        AvailablePoints = user?.Points ?? 0;
        MemberTier = user?.CustomerTier ?? "普通会员";

        // 累计获得
        TotalEarned = await _db.PointsLedgers
            .Where(l => l.UserId == userId && l.PointType == "earn")
            .SumAsync(l => (int?)l.Points) ?? 0;

        // 累计已使用
        TotalRedeemed = await _db.PointsLedgers
            .Where(l => l.UserId == userId && (l.PointType == "redeem" || l.PointType == "expire"))
            .SumAsync(l => (int?)Math.Abs(l.Points)) ?? 0;

        // 30天内到期
        var now = DateTime.Now;
        var thirtyDays = now.AddDays(30);
        ExpiringSoon = await _db.PointsLedgers
            .Where(l => l.UserId == userId && l.PointType == "earn" && !l.IsExpired
                        && l.ExpiresAt != null && l.ExpiresAt >= now && l.ExpiresAt <= thirtyDays)
            .SumAsync(l => (int?)l.Points) ?? 0;

        // 积分账本（分页）
        var totalCount = await _db.PointsLedgers.CountAsync(l => l.UserId == userId);
        TotalPages = Math.Max(1, (int)Math.Ceiling((double)totalCount / pageSize));

        LedgerItems = await _db.PointsLedgers
            .Where(l => l.UserId == userId)
            .OrderByDescending(l => l.CreatedAt)
            .Skip((CurrentPage - 1) * pageSize)
            .Take(pageSize)
            .ToListAsync();

        // 兑换商品
        RedemptionItems = await _db.PointsRedemptions
            .Where(r => r.IsEnabled)
            .OrderBy(r => r.SortOrder)
            .ToListAsync();

        // 积分规则
        Rules = await _db.PointsRules.ToListAsync();
    }

    public static string GetSourceName(string source) => source switch
    {
        "purchase" => "消费购物",
        "signin" => "每日签到",
        "review" => "发表评价",
        "share" => "分享推广",
        "redeem" => "积分兑换",
        "expire" => "积分过期",
        "adjust" => "人工调整",
        _ => source
    };
}
