using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 商品类型服务 — 三种香型业务逻辑
/// 对应 ASP 中的 product_type_utils.asp
/// </summary>
public class ProductTypeService : IProductTypeService
{
    private readonly PerfumeShopContext _db;

    public ProductTypeService(PerfumeShopContext db)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
    }

    public async Task<IEnumerable<ProductTypeDto>> GetActiveProductTypesAsync(CancellationToken ct = default)
    {
        return await _db.ProductTypeConfigs
            .AsNoTracking()
            .Where(pt => pt.IsActive == true)
            .OrderBy(pt => pt.DisplayOrder)
            .Select(pt => new ProductTypeDto
            {
                TypeCode = pt.TypeCode,
                DisplayName = pt.DisplayName ?? pt.TypeCode,
                NavName = pt.NavName,
                Description = pt.Description,
                Icon = pt.Icon,
                RequiresReview = pt.RequiresReview ?? false,
                RequiresRatio = pt.RequiresRatio ?? false,
                DisplayOrder = pt.DisplayOrder ?? 0,
                IsActive = pt.IsActive ?? false
            })
            .ToListAsync(ct);
    }

    public async Task<bool> IsProductTypeActiveAsync(string typeCode, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(typeCode)) return false;
        return await _db.ProductTypeConfigs
            .AnyAsync(pt => pt.TypeCode == typeCode && pt.IsActive == true, ct);
    }

    public async Task<string?> GetDisplayNameAsync(string typeCode, string? dbName = null, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(typeCode)) return dbName;

        var config = await _db.ProductTypeConfigs
            .AsNoTracking()
            .FirstOrDefaultAsync(pt => pt.TypeCode == typeCode, ct);

        return config?.DisplayName ?? dbName;
    }

    public async Task<IEnumerable<string>> GetActiveTypeCodesAsync(CancellationToken ct = default)
    {
        return await _db.ProductTypeConfigs
            .AsNoTracking()
            .Where(pt => pt.IsActive == true)
            .OrderBy(pt => pt.DisplayOrder)
            .Select(pt => pt.TypeCode)
            .ToListAsync(ct);
    }
}
