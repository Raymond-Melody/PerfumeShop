using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 商品查询 API
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class ProductsController : ControllerBase
{
    private readonly IProductRepository _productRepo;
    private readonly IProductTypeService _productTypeService;
    private readonly IRecommendationEngine _recommendation;

    public ProductsController(
        IProductRepository productRepo,
        IProductTypeService productTypeService,
        IRecommendationEngine recommendation)
    {
        _productRepo = productRepo;
        _productTypeService = productTypeService;
        _recommendation = recommendation;
    }

    /// <summary>获取商品列表 (支持分页和类型筛选)</summary>
    [HttpGet]
    public async Task<IActionResult> GetProducts(
        [FromQuery] int page = 1,
        [FromQuery] int pageSize = 20,
        [FromQuery] string? type = null)
    {
        if (page < 1) page = 1;
        if (pageSize < 1) pageSize = 20;
        if (pageSize > 100) pageSize = 100;

        var products = string.IsNullOrEmpty(type)
            ? await _productRepo.GetPagedAsync(page, pageSize, p => p.IsActive == true, q => q.OrderByDescending(p => p.CreatedAt))
            : await _productRepo.GetPagedAsync(page, pageSize,
                p => p.IsActive == true && p.ProductType == type,
                q => q.OrderByDescending(p => p.CreatedAt));

        return Ok(new
        {
            items = products.Items.Select(p => new
            {
                p.ProductId,
                p.ProductName,
                p.ProductType,
                p.Category,
                p.BasePrice,
                p.ImageUrl,
                p.Description,
                p.IsActive
            }),
            total = products.TotalCount,
            page,
            pageSize
        });
    }

    /// <summary>获取商品详情</summary>
    [HttpGet("{id:int}")]
    public async Task<IActionResult> GetProduct(int id)
    {
        var product = await _productRepo.GetByIdAsync(id);
        if (product == null)
            return NotFound(new { message = "商品不存在" });

        return Ok(new
        {
            product.ProductId,
            product.ProductName,
            product.ProductType,
            product.Category,
            product.BasePrice,
            product.UnitCost,
            product.ImageUrl,
            product.Description,
            product.ReviewStatus,
            product.Engravable,
            product.EngravingPrice,
            product.IsActive,
            product.CreatedAt
        });
    }

    /// <summary>搜索商品</summary>
    [HttpGet("search")]
    public async Task<IActionResult> Search([FromQuery] string q)
    {
        if (string.IsNullOrWhiteSpace(q))
            return BadRequest(new { message = "搜索关键词不能为空" });

        var results = await _productRepo.SearchAsync(q.Trim());
        return Ok(results.Select(p => new
        {
            p.ProductId,
            p.ProductName,
            p.ProductType,
            p.Category,
            p.BasePrice,
            p.ImageUrl
        }));
    }

    /// <summary>获取商品类型列表</summary>
    [HttpGet("types")]
    public async Task<IActionResult> GetProductTypes()
    {
        var types = await _productTypeService.GetActiveProductTypesAsync();
        return Ok(types);
    }

    /// <summary>热门商品推荐</summary>
    [HttpGet("popular")]
    public async Task<IActionResult> GetPopular([FromQuery] int count = 10)
    {
        var ids = await _recommendation.GetPopularProductsAsync(count);
        return Ok(new { productIds = ids });
    }

    /// <summary>新品推荐</summary>
    [HttpGet("new-arrivals")]
    public async Task<IActionResult> GetNewArrivals([FromQuery] int count = 6)
    {
        var ids = await _recommendation.GetNewArrivalsAsync(count);
        return Ok(new { productIds = ids });
    }

    /// <summary>关联商品推荐</summary>
    [HttpGet("{id:int}/related")]
    public async Task<IActionResult> GetRelated(int id, [FromQuery] int count = 4)
    {
        var ids = await _recommendation.GetRelatedProductsAsync(id, count);
        return Ok(new { productIds = ids });
    }
}
