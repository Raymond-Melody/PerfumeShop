using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 智能搜索建议 API — 对应 V18 api/search_suggestions.asp
/// </summary>
[ApiController]
[Route("api/v2/search")]
public class SearchSuggestionsController : ControllerBase
{
    private readonly PerfumeShopContext _db;

    // Synonym map for search expansion
    private static readonly Dictionary<string, string[]> SynonymMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["清新"] = new[] { "海洋", "柑橘", "绿茶", "薰衣草" },
        ["清爽"] = new[] { "海洋", "柑橘", "绿茶", "薰衣草" },
        ["浓郁"] = new[] { "东方", "木质", "琥珀", "檀香" },
        ["持久"] = new[] { "东方", "木质", "琥珀", "檀香" },
        ["温柔"] = new[] { "花香", "玫瑰", "茉莉" },
        ["阳光"] = new[] { "柑橘", "柠檬", "海洋" },
        ["性感"] = new[] { "东方", "麝香", "琥珀" },
        ["成熟"] = new[] { "木质", "皮革", "雪松" },
        ["甜美"] = new[] { "果香", "花香", "香草" },
        ["中性"] = new[] { "柑橘", "绿叶", "绿茶" },
    };

    public SearchSuggestionsController(PerfumeShopContext db)
    {
        _db = db;
    }

    /// <summary>GET /api/v2/search/suggestions?q=keyword — 智能搜索建议</summary>
    [HttpGet("suggestions")]
    public async Task<IActionResult> GetSuggestions([FromQuery] string q = "", [FromQuery] int max = 8)
    {
        if (string.IsNullOrWhiteSpace(q))
            return Ok(new { success = true, data = new { items = Array.Empty<object>(), intent = "" } });

        max = Math.Clamp(max, 1, 20);

        try
        {
            // Detect search intent
            var intent = DetectIntent(q);

            // Expand synonyms
            var searchTerms = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { q };
            if (SynonymMap.TryGetValue(q, out var synonyms))
                foreach (var s in synonyms) searchTerms.Add(s);

            // Query products matching any search term
            var seen = new HashSet<int>();
            var items = new List<object>();

            foreach (var term in searchTerms)
            {
                if (items.Count >= max) break;

                var products = await _db.Products
                    .Where(p => p.IsActive == true &&
                        (p.ProductName.Contains(term) ||
                         (p.Description != null && p.Description.Contains(term)) ||
                         (p.Category != null && p.Category.Contains(term))))
                    .Take(max - items.Count + 2)
                    .Select(p => new
                    {
                        p.ProductId,
                        p.ProductName,
                        p.BasePrice,
                        p.ImageUrl,
                        p.ProductType
                    })
                    .ToListAsync();

                foreach (var p in products)
                {
                    if (items.Count >= max) break;
                    if (seen.Add(p.ProductId))
                    {
                        items.Add(new
                        {
                            id = p.ProductId,
                            name = p.ProductName,
                            price = (double)p.BasePrice,
                            image = p.ImageUrl,
                            type = p.ProductType,
                            matchType = "keyword"
                        });
                    }
                }
            }

            return Ok(new
            {
                success = true,
                data = new
                {
                    items,
                    intent
                }
            });
        }
        catch (Exception ex)
        {
            return Ok(new
            {
                success = true,
                data = new
                {
                    items = Array.Empty<object>(),
                    intent = "",
                    error = ex.Message
                }
            });
        }
    }

    private static string DetectIntent(string keyword)
    {
        var kw = keyword.ToLower();

        // Gift intent
        if (kw.Contains("送女友") || kw.Contains("送女生") || kw.Contains("女朋友"))
            return "gift_female";
        if (kw.Contains("送男友") || kw.Contains("送男生") || kw.Contains("男朋友"))
            return "gift_male";
        if (kw.Contains("送礼") || kw.Contains("礼物"))
            return "gift_general";

        // Season intent
        if (kw.Contains("夏天") || kw.Contains("夏季")) return "season_summer";
        if (kw.Contains("冬天") || kw.Contains("冬季")) return "season_winter";
        if (kw.Contains("春天") || kw.Contains("春季")) return "season_spring";
        if (kw.Contains("秋天") || kw.Contains("秋季")) return "season_autumn";

        // Scene intent
        if (kw.Contains("上班") || kw.Contains("职场")) return "scene_work";
        if (kw.Contains("约会") || kw.Contains("派对")) return "scene_social";
        if (kw.Contains("运动") || kw.Contains("户外")) return "scene_sport";

        // Style intent
        if (kw.Contains("清新") || kw.Contains("清爽")) return "style_fresh";
        if (kw.Contains("浓郁") || kw.Contains("持久")) return "style_strong";

        return "";
    }
}
