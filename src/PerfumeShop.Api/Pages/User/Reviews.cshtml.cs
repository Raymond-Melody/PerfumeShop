using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class ReviewsModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public ReviewsModel(PerfumeShopContext db) => _db = db;

    public List<ReviewItem> Reviews { get; set; } = new();
    public int TotalReviews { get; set; }
    public int CurrentPage { get; set; } = 1;
    public int TotalPages { get; set; }
    public string? Message { get; set; }
    public string? Error { get; set; }

    public class ReviewItem
    {
        public int ReviewId { get; set; }
        public int? Rating { get; set; }
        public string? Title { get; set; }
        public string? Comment { get; set; }
        public DateTime? CreatedAt { get; set; }
        public bool IsVerifiedPurchase { get; set; }
        public int LikeCount { get; set; }
        public string? AIFeelingSummary { get; set; }
        public int? ProductId { get; set; }
        public string? ProductName { get; set; }
        public string? ProductImage { get; set; }
    }

    public async Task OnGetAsync(int page = 1)
    {
        CurrentPage = page < 1 ? 1 : page;
        const int pageSize = 10;

        var query = from r in _db.ProductReviews
                    join p in _db.Products on r.ProductId equals p.ProductId into pp
                    from p in pp.DefaultIfEmpty()
                    where r.Status != "Deleted"
                    orderby r.CreatedAt descending
                    select new ReviewItem
                    {
                        ReviewId = r.ReviewId,
                        Rating = r.Rating,
                        Title = r.Title,
                        Comment = r.Comment,
                        CreatedAt = r.CreatedAt,
                        IsVerifiedPurchase = r.IsVerifiedPurchase,
                        LikeCount = r.LikeCount,
                        AIFeelingSummary = r.AIFeelingSummary,
                        ProductId = r.ProductId,
                        ProductName = p != null ? p.ProductName : null,
                        ProductImage = p != null ? p.ImageUrl : null,
                    };

        TotalReviews = await query.CountAsync();
        TotalPages = (int)Math.Ceiling(TotalReviews / (double)pageSize);
        Reviews = await query.Skip((CurrentPage - 1) * pageSize).Take(pageSize).ToListAsync();
    }
}
