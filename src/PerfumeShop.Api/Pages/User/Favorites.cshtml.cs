using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class FavoritesModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public FavoritesModel(PerfumeShopContext db) => _db = db;

    public List<FavoriteItem> Favorites { get; set; } = new();
    public bool IsEmpty => Favorites.Count == 0;

    public class FavoriteItem
    {
        public int FavoriteId { get; set; }
        public int ProductId { get; set; }
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public string? Image { get; set; }
        public string? Category { get; set; }
        public string? Description { get; set; }
        public DateTime CreatedTime { get; set; }
    }

    public async Task OnGetAsync()
    {
        // TODO: 从 AuthService 获取真实 userId
        var userId = 1;

        Favorites = await (from f in _db.UserFavorites
                           join p in _db.Products on f.ProductId equals p.ProductId
                           where f.UserId == userId
                           orderby f.CreatedTime descending
                           select new FavoriteItem
                           {
                               FavoriteId = f.FavoriteId,
                               ProductId = f.ProductId,
                               Name = p.ProductName,
                               Price = p.BasePrice,
                               Image = p.ImageUrl,
                               Category = p.Category,
                               Description = p.Description,
                               CreatedTime = f.CreatedTime ?? DateTime.MinValue
                           }).ToListAsync();
    }
}
