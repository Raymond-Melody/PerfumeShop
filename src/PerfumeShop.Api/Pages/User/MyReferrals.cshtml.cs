using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class MyReferralsModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public MyReferralsModel(PerfumeShopContext db) => _db = db;

    public int TotalInvites { get; set; }
    public int TotalRewardPoints { get; set; }
    public int ActiveTokenCount { get; set; }
    public List<InvitedUserItem> InvitedUsers { get; set; } = new();
    public List<TokenItem> Tokens { get; set; } = new();
    public string? CurrentLink { get; set; }

    public class InvitedUserItem
    {
        public int UserId { get; set; }
        public string? Username { get; set; }
        public DateTime? CreatedAt { get; set; }
    }

    public class TokenItem
    {
        public int TokenId { get; set; }
        public string TokenHash { get; set; } = "";
        public int? UsedCount { get; set; }
        public int? MaxUses { get; set; }
        public bool? IsActive { get; set; }
        public DateTime ExpiresAt { get; set; }
        public DateTime? CreatedAt { get; set; }
        public string Link => $"/register?ref={TokenHash}";
    }

    public async Task OnGetAsync()
    {
        // For demo: use userId=1; in production extract from auth
        var userId = 1;
        var tokens = await _db.ReferralTokens.Where(t => t.ReferrerUserId == userId).ToListAsync();
        TotalInvites = await _db.ReferralRelations.CountAsync(r => r.AncestorUserId == userId && r.Depth == 1);
        TotalRewardPoints = tokens.Sum(t => t.UsedCount ?? 0) * 100;
        ActiveTokenCount = tokens.Count(t => t.IsActive == true);

        InvitedUsers = await (from r in _db.ReferralRelations
                              join u in _db.Users on r.DescendantUserId equals u.UserId
                              where r.AncestorUserId == userId && r.Depth == 1
                              select new InvitedUserItem { UserId = u.UserId, Username = u.Username, CreatedAt = u.CreatedAt }).ToListAsync();

        Tokens = tokens.Select(t => new TokenItem
        {
            TokenId = t.TokenId,
            TokenHash = t.TokenHash,
            UsedCount = t.UsedCount,
            MaxUses = t.MaxUses,
            IsActive = t.IsActive,
            ExpiresAt = t.ExpiresAt,
            CreatedAt = t.CreatedAt,
        }).OrderByDescending(t => t.CreatedAt).ToList();

        var active = tokens.FirstOrDefault(t => t.IsActive == true);
        if (active != null) CurrentLink = $"/register?ref={active.TokenHash}";
    }
}
