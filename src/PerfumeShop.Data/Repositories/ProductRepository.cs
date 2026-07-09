using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 商品仓储实现 — 对应 ASP 中的 product_type_utils.asp
/// </summary>
public class ProductRepository : Repository<Product>, IProductRepository
{
    public ProductRepository(PerfumeShopContext context) : base(context) { }

    // ========== 重写 GetByIdAsync — Product 为 keyless 实体 ==========

    public override async Task<Product?> GetByIdAsync(int id, CancellationToken ct = default)
    {
        return await _dbSet.FirstOrDefaultAsync(p => p.ProductId == id, ct);
    }

    // ========== IProductRepository 实现 ==========

    public async Task<IEnumerable<Product>> GetActiveProductsAsync(CancellationToken ct = default)
    {
        return await _dbSet.AsNoTracking()
            .Where(p => p.IsActive == true)
            .OrderByDescending(p => p.CreatedAt)
            .ToListAsync(ct);
    }

    public async Task<IEnumerable<Product>> GetByProductTypeAsync(string productType, CancellationToken ct = default)
    {
        return await _dbSet.AsNoTracking()
            .Where(p => p.ProductType == productType && p.IsActive == true)
            .OrderByDescending(p => p.CreatedAt)
            .ToListAsync(ct);
    }

    public async Task<IEnumerable<Product>> GetByCategoryAsync(string category, CancellationToken ct = default)
    {
        return await _dbSet.AsNoTracking()
            .Where(p => p.Category == category && p.IsActive == true)
            .OrderByDescending(p => p.CreatedAt)
            .ToListAsync(ct);
    }

    public async Task<IEnumerable<Product>> SearchAsync(string keyword, CancellationToken ct = default)
    {
        var lowerKeyword = keyword.ToLower();
        return await _dbSet.AsNoTracking()
            .Where(p => p.IsActive == true &&
                        (p.ProductName.Contains(keyword) ||
                         (p.Description != null && p.Description.Contains(keyword))))
            .OrderByDescending(p => p.CreatedAt)
            .Take(50)
            .ToListAsync(ct);
    }

    public async Task<IEnumerable<ProductTypeConfig>> GetActiveProductTypesAsync(CancellationToken ct = default)
    {
        return await _context.ProductTypeConfigs
            .AsNoTracking()
            .Where(pt => pt.IsActive == true)
            .OrderBy(pt => pt.DisplayOrder)
            .ToListAsync(ct);
    }

    public async Task<bool> IsProductTypeActiveAsync(string typeCode, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(typeCode)) return false;
        return await _context.ProductTypeConfigs
            .AnyAsync(pt => pt.TypeCode == typeCode && pt.IsActive == true, ct);
    }

    public async Task<string?> GetProductTypeDisplayNameAsync(string typeCode, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(typeCode)) return null;
        var config = await _context.ProductTypeConfigs
            .AsNoTracking()
            .FirstOrDefaultAsync(pt => pt.TypeCode == typeCode, ct);
        return config?.DisplayName;
    }

    public async Task<IEnumerable<string>> GetActiveTypeCodesAsync(CancellationToken ct = default)
    {
        return await _context.ProductTypeConfigs
            .AsNoTracking()
            .Where(pt => pt.IsActive == true)
            .OrderBy(pt => pt.DisplayOrder)
            .Select(pt => pt.TypeCode)
            .ToListAsync(ct);
    }
}
