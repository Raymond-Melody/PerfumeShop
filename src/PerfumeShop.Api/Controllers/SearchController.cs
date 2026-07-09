using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/search")]
public class SearchController : ControllerBase
{
    private readonly PerfumeShopContext _db;
    public SearchController(PerfumeShopContext db) => _db = db;

    /// <summary>GET /api/search/suggest?q=关键词&max=8 — 智能搜索建议</summary>
    [HttpGet("suggest")]
    public async Task<IActionResult> Suggest([FromQuery] string q = "", [FromQuery] int max = 8)
    {
        if (string.IsNullOrWhiteSpace(q) || q.Length < 1)
            return Ok(new { success = true, data = new List<object>() });

        if (max < 1 || max > 20) max = 8;

        var keyword = q.Trim();

        // 商品名称匹配
        var products = await _db.Products
            .Where(p => p.ProductName.Contains(keyword) || (p.Category != null && p.Category.Contains(keyword)))
            .OrderBy(p => p.ProductName)
            .Take(max)
            .Select(p => new SearchSuggestion
            {
                Text = p.ProductName,
                Type = "product",
                ProductId = p.ProductId,
                ImageUrl = p.ImageUrl
            }).ToListAsync();

        // 分类匹配
        var categories = await _db.Products
            .Where(p => p.Category != null && p.Category.Contains(keyword))
            .Select(p => p.Category!)
            .Distinct()
            .Take(3)
            .ToListAsync();

        var suggestions = products.Select(p => new { p.Text, p.Type, p.ProductId, p.ImageUrl })
            .Concat(categories.Select(c => new { Text = c, Type = (string?)"category", ProductId = (int?)null, ImageUrl = (string?)null }))
            .Take(max)
            .ToList();

        return Ok(new { success = true, data = suggestions });
    }

    /// <summary>GET /api/search?q=关键词&page=1&pageSize=20 — 全量搜索</summary>
    [HttpGet]
    public async Task<IActionResult> Search([FromQuery] string q = "", [FromQuery] int page = 1, [FromQuery] int pageSize = 20)
    {
        if (string.IsNullOrWhiteSpace(q))
            return Ok(new { success = true, data = new List<object>(), total = 0 });

        var keyword = q.Trim();
        var query = _db.Products
            .Where(p => p.ProductName.Contains(keyword)
                     || (p.Description != null && p.Description.Contains(keyword))
                     || (p.Category != null && p.Category.Contains(keyword)));

        var total = await query.CountAsync();
        var items = await query
            .OrderBy(p => p.ProductName)
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .Select(p => new
            {
                productId = p.ProductId,
                name = p.ProductName,
                price = p.BasePrice,
                image = p.ImageUrl,
                category = p.Category,
                type = p.ProductType
            }).ToListAsync();

        return Ok(new { success = true, data = items, total, page, pageSize });
    }
}
