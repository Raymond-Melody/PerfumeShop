using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "ProductList")]
public class ProductsModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public ProductsModel(PerfumeShopContext db) => _db = db;

    public List<Product> Products { get; set; } = new();
    public string Keyword { get; set; } = "";
    public string? TypeFilter { get; set; }
    public string? SortBy { get; set; }
    public int CurrentPage { get; set; } = 1;
    public int TotalPages { get; set; }
    public int TotalCount { get; set; }

    private const int PageSize = 12;

    public async Task OnGetAsync(string? keyword, string? type, string? sort, int page = 1)
    {
        Keyword = keyword?.Trim() ?? "";
        TypeFilter = type;
        SortBy = sort;
        CurrentPage = Math.Max(1, page);

        var query = _db.Products.AsNoTracking().Where(p => p.IsActive == true);

        if (!string.IsNullOrEmpty(Keyword))
            query = query.Where(p => p.ProductName.Contains(Keyword));

        if (!string.IsNullOrEmpty(TypeFilter))
            query = query.Where(p => p.ProductType == TypeFilter);

        TotalCount = await query.CountAsync();
        TotalPages = (int)Math.Ceiling((double)TotalCount / PageSize);
        CurrentPage = Math.Min(CurrentPage, Math.Max(1, TotalPages));

        query = SortBy switch
        {
            "price_asc" => query.OrderBy(p => p.BasePrice),
            "price_desc" => query.OrderByDescending(p => p.BasePrice),
            "newest" => query.OrderByDescending(p => p.CreatedAt),
            _ => query.OrderByDescending(p => p.ProductId)
        };

        Products = await query
            .Skip((CurrentPage - 1) * PageSize)
            .Take(PageSize)
            .ToListAsync();
    }
}
