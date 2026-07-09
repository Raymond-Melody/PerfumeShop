using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Interfaces;

/// <summary>
/// 商品仓储接口 — 商品相关的业务查询 (对应 ASP 中的 product_type_utils.asp)
/// </summary>
public interface IProductRepository : IRepository<Product>
{
    /// <summary>获取所有已启用的商品</summary>
    Task<IEnumerable<Product>> GetActiveProductsAsync(CancellationToken ct = default);

    /// <summary>按商品类型筛选已启用商品</summary>
    Task<IEnumerable<Product>> GetByProductTypeAsync(string productType, CancellationToken ct = default);

    /// <summary>按分类获取商品</summary>
    Task<IEnumerable<Product>> GetByCategoryAsync(string category, CancellationToken ct = default);

    /// <summary>搜索商品 (按名称/描述模糊匹配)</summary>
    Task<IEnumerable<Product>> SearchAsync(string keyword, CancellationToken ct = default);

    /// <summary>获取所有启用的商品类型配置</summary>
    Task<IEnumerable<ProductTypeConfig>> GetActiveProductTypesAsync(CancellationToken ct = default);

    /// <summary>判断商品类型是否启用</summary>
    Task<bool> IsProductTypeActiveAsync(string typeCode, CancellationToken ct = default);

    /// <summary>获取商品类型显示名称</summary>
    Task<string?> GetProductTypeDisplayNameAsync(string typeCode, CancellationToken ct = default);

    /// <summary>获取所有已启用类型的类型代码</summary>
    Task<IEnumerable<string>> GetActiveTypeCodesAsync(CancellationToken ct = default);
}
