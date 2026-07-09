namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 商品类型服务接口 — 三种香型业务逻辑
/// 对应 ASP 中的 product_type_utils.asp
/// </summary>
public interface IProductTypeService
{
    /// <summary>获取所有启用的商品类型</summary>
    Task<IEnumerable<ProductTypeDto>> GetActiveProductTypesAsync(CancellationToken ct = default);

    /// <summary>判断商品类型是否启用</summary>
    Task<bool> IsProductTypeActiveAsync(string typeCode, CancellationToken ct = default);

    /// <summary>获取商品类型显示名称 (含i18n)</summary>
    Task<string?> GetDisplayNameAsync(string typeCode, string? dbName = null, CancellationToken ct = default);

    /// <summary>获取所有已启用类型的代码列表</summary>
    Task<IEnumerable<string>> GetActiveTypeCodesAsync(CancellationToken ct = default);
}

/// <summary>商品类型 DTO</summary>
public class ProductTypeDto
{
    public string TypeCode { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string? NavName { get; set; }
    public string? Description { get; set; }
    public string? Icon { get; set; }
    public bool RequiresReview { get; set; }
    public bool RequiresRatio { get; set; }
    public int DisplayOrder { get; set; }
    public bool IsActive { get; set; }
}
