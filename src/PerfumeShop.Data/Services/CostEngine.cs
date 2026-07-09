using System.Collections.Concurrent;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 成本自动传导引擎 — 三级成本计算 (原材料 → 香调 → 产品 → 订单)
/// 对应 ASP 中的 cost_engine.asp
/// </summary>
public class CostEngine : ICostEngine
{
    private readonly PerfumeShopContext _db;

    // 请求级别缓存（避免同一请求内重复计算）
    private readonly ConcurrentDictionary<int, decimal> _materialCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _noteCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _bomCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _unitCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _fixedBrandCostCache = new();

    public CostEngine(PerfumeShopContext db)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
    }

    // ==================== Level 1: 原料成本 ====================

    public async Task<decimal> CalculateMaterialCostAsync(int materialId, CancellationToken ct = default)
    {
        if (_materialCostCache.TryGetValue(materialId, out var cached))
            return cached;

        decimal cost = 0;

        // V10: 优先使用加权平均成本
        var material = await _db.RawMaterialInventories
            .AsNoTracking()
            .FirstOrDefaultAsync(m => m.MaterialId == materialId, ct);

        if (material != null)
        {
            // 加权平均成本优先
            if ((material.WeightedUnitCost ?? 0) > 0)
            {
                cost = material.WeightedUnitCost!.Value;
            }
            // 回退到供应商最新报价
            else if (!string.IsNullOrEmpty(material.ItemCode))
            {
                var supplierPrice = await _db.SupplierPrices
                    .AsNoTracking()
                    .Where(sp => sp.ItemCode == material.ItemCode && sp.IsActive == true)
                    .OrderByDescending(sp => sp.CreatedAt)
                    .Select(sp => sp.UnitPrice)
                    .FirstOrDefaultAsync(ct);

                cost = supplierPrice ?? 0;
            }

            // 最终回退到库存中的 UnitPrice
            if (cost <= 0)
                cost = material.UnitPrice ?? 0;
        }

        _materialCostCache.TryAdd(materialId, cost);
        return cost;
    }

    // ==================== Level 2: 香调成本 ====================

    public async Task<decimal> CalculateNoteCostAsync(int noteId, CancellationToken ct = default)
    {
        if (_noteCostCache.TryGetValue(noteId, out var cached))
            return cached;

        decimal totalCost = 0;
        bool hasAccord = false;

        // 路径A: Accord 生产配方
        var accord = await _db.RecipeAccords
            .AsNoTracking()
            .Where(ra => ra.NoteId == noteId && ra.Status == "Published")
            .OrderByDescending(ra => ra.PublishedAt)
            .FirstOrDefaultAsync(ct);

        if (accord != null)
        {
            hasAccord = true;
            var batchSize = accord.BatchSize > 0 ? (decimal)accord.BatchSize : 100m;

            var materials = await _db.RecipeAccordMaterials
                .AsNoTracking()
                .Where(ram => ram.AccordRecipeId == accord.AccordRecipeId)
                .ToListAsync(ct);

            foreach (var mat in materials)
            {
                var matCost = await CalculateMaterialCostAsync(mat.MaterialId ?? 0, ct);
                var pct = (decimal)(mat.Percentage ?? 0);
                totalCost += (pct / batchSize) * matCost;
            }
        }

        // 路径B: BaseNote 成分聚合
        if (!hasAccord)
        {
            var ingredients = await _db.NoteIngredients
                .AsNoTracking()
                .Where(ni => ni.NoteId == noteId)
                .ToListAsync(ct);

            if (ingredients.Count > 0)
            {
                foreach (var ing in ingredients)
                {
                    var baseNoteCost = await CalculateNoteCostAsync(ing.BaseNoteId, ct);

                    // V9: 回退到 BaseNotes 单价
                    if (baseNoteCost <= 0)
                    {
                        baseNoteCost = await _db.BaseNotes
                            .AsNoTracking()
                            .Where(bn => bn.BaseNoteId == ing.BaseNoteId && bn.IsActive == true)
                            .Select(bn => bn.UnitPrice ?? 0)
                            .FirstOrDefaultAsync(ct);
                    }

                    // 兜底: PriceAddition
                    if (baseNoteCost <= 0)
                    {
                        baseNoteCost = await _db.FragranceNotes
                            .AsNoTracking()
                            .Where(fn => fn.NoteId == ing.BaseNoteId)
                            .Select(fn => fn.PriceAddition ?? 0)
                            .FirstOrDefaultAsync(ct);
                    }

                    var pct = (decimal)(ing.Percentage ?? 0);
                    totalCost += baseNoteCost * pct / 100m;
                }
            }
        }

        // 兜底: PriceAddition
        if (totalCost <= 0)
        {
            totalCost = await _db.FragranceNotes
                .AsNoTracking()
                .Where(fn => fn.NoteId == noteId)
                .Select(fn => fn.PriceAddition ?? 0)
                .FirstOrDefaultAsync(ct);
        }

        _noteCostCache.TryAdd(noteId, totalCost);
        return totalCost;
    }

    // ==================== Level 3: 产品 BOM 成本 ====================

    public async Task<decimal> CalculateProductBomCostAsync(int productId, CancellationToken ct = default)
    {
        if (_bomCostCache.TryGetValue(productId, out var cached))
            return cached;

        decimal totalCost = 0;

        // 1. 香调配比成本
        var noteRatios = await _db.ProductNoteRatios
            .AsNoTracking()
            .Where(pnr => pnr.ProductId == productId)
            .ToListAsync(ct);

        foreach (var nr in noteRatios)
        {
            var noteCost = await CalculateNoteCostAsync(nr.NoteId, ct);
            totalCost += noteCost * nr.Percentage / 100m;
        }

        // 2. 瓶身成本
        var bottleCost = await (from pbs in _db.ProductBottleStyles
                                join bs in _db.BottleStyles on pbs.BottleId equals bs.BottleId
                                where pbs.ProductId == productId && bs.IsActive == true
                                select bs.PriceAddition ?? 0)
                                .FirstOrDefaultAsync(ct);

        totalCost += bottleCost;

        _bomCostCache.TryAdd(productId, totalCost);
        return totalCost;
    }

    // ==================== Level 4: 产品完整单位成本 ====================

    public async Task<decimal> CalculateProductUnitCostAsync(int productId, CancellationToken ct = default)
    {
        if (_unitCostCache.TryGetValue(productId, out var cached))
            return cached;

        // 品牌定香(standard/Fixed): 优先使用采购加权平均成本
        var productType = await _db.Products
            .AsNoTracking()
            .Where(p => p.ProductId == productId)
            .Select(p => p.ProductType)
            .FirstOrDefaultAsync(ct);

        if (productType == "standard")
        {
            var fbCost = await GetFixedBrandCostAsync(productId, ct);
            if (fbCost > 0)
            {
                _unitCostCache.TryAdd(productId, fbCost);
                return fbCost;
            }

            // 回退: 数据库中的现有 UnitCost
            var existing = await _db.Products
                .AsNoTracking()
                .Where(p => p.ProductId == productId)
                .Select(p => p.UnitCost ?? 0)
                .FirstOrDefaultAsync(ct);

            if (existing > 0)
            {
                _unitCostCache.TryAdd(productId, existing);
                return existing;
            }
        }

        // 非 Fixed 或回退: BOM + 包装 + 人工
        decimal totalCost = await CalculateProductBomCostAsync(productId, ct);

        // 包装成本
        var packagingCost = await _db.ProductCosts
            .AsNoTracking()
            .Where(pc => pc.ProductId == productId && pc.CostType == "Packaging")
            .SumAsync(pc => pc.TotalCost ?? 0, ct);

        // 人工/管理费用
        var otherCost = await _db.ProductCosts
            .AsNoTracking()
            .Where(pc => pc.ProductId == productId && pc.CostType == "Other")
            .SumAsync(pc => pc.TotalCost ?? 0, ct);

        totalCost += packagingCost + otherCost;

        _unitCostCache.TryAdd(productId, totalCost);
        return totalCost;
    }

    // ==================== 辅助: 品牌定香采购成本 ====================

    private async Task<decimal> GetFixedBrandCostAsync(int productId, CancellationToken ct)
    {
        if (_fixedBrandCostCache.TryGetValue(productId, out var cached))
            return cached;

        decimal? cost = null;
        try
        {
            cost = await (from fbp in _db.FixedBrandProducts
                          join fbi in _db.FixedBrandInventories on fbp.FixedProductId equals fbi.FixedProductId into fbiJoin
                          from fbi in fbiJoin.DefaultIfEmpty()
                          where fbp.ProductId == productId && fbp.Status == "Active"
                          select fbi != null && fbi.AvgUnitCost > 0 ? fbi.AvgUnitCost : fbp.UnitPrice)
                          .FirstOrDefaultAsync(ct);
        }
        catch
        {
            // 表可能不存在，忽略错误
        }

        var result = cost ?? 0m;
        _fixedBrandCostCache.TryAdd(productId, result);
        return result;
    }

    // ==================== Level 5: 更新产品成本到数据库 ====================

    public async Task UpdateProductCostAsync(int productId, CancellationToken ct = default)
    {
        var bomCost = await CalculateProductBomCostAsync(productId, ct);
        var unitCost = await CalculateProductUnitCostAsync(productId, ct);

        await _db.Products
            .Where(p => p.ProductId == productId)
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(p => p.Bomcost, bomCost)
                .SetProperty(p => p.UnitCost, unitCost), ct);
    }

    public async Task<int> UpdateAllProductCostsAsync(CancellationToken ct = default)
    {
        var productIds = await _db.Products
            .AsNoTracking()
            .Where(p => p.IsActive == true)
            .Select(p => p.ProductId)
            .ToListAsync(ct);

        int updated = 0;
        foreach (var pid in productIds)
        {
            try
            {
                await UpdateProductCostAsync(pid, ct);
                updated++;
            }
            catch
            {
                // 跳过计算失败的产品
            }
        }

        return updated;
    }

    // ==================== Level 6 & 7: 订单成本与利润 ====================

    public async Task UpdateOrderCostsAsync(int orderId, CancellationToken ct = default)
    {
        // 计算订单中各商品的成本总和
        var details = await _db.OrderDetails
            .AsNoTracking()
            .Where(od => od.OrderId == orderId)
            .Select(od => new { od.ProductId, od.Quantity })
            .ToListAsync(ct);

        decimal orderCost = 0;
        foreach (var d in details)
        {
            var unitCost = await CalculateProductUnitCostAsync(d.ProductId, ct);
            orderCost += unitCost * d.Quantity;
        }

        // 获取订单总额和运费
        var orderData = await _db.Orders
            .AsNoTracking()
            .Where(o => o.OrderId == orderId)
            .Select(o => new { o.TotalAmount, o.ShippingFee })
            .FirstOrDefaultAsync(ct);

        if (orderData == null) return;

        var profitAmount = orderData.TotalAmount - orderCost - (orderData.ShippingFee ?? 0);
        if (profitAmount < 0) profitAmount = 0;

        await _db.Orders
            .Where(o => o.OrderId == orderId)
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(o => o.CostAmount, orderCost)
                .SetProperty(o => o.ProfitAmount, profitAmount), ct);

        // V10: 记录成本分摊明细
        await RecordOrderCostAllocationAsync(orderId, ct);
    }

    public async Task UpdateAllOrderCostsAsync(CancellationToken ct = default)
    {
        var orderIds = await _db.Orders
            .AsNoTracking()
            .Where(o => o.Status != "Pending" && o.Status != "Cancelled")
            .Select(o => o.OrderId)
            .ToListAsync(ct);

        foreach (var oid in orderIds)
        {
            await UpdateOrderCostsAsync(oid, ct);
        }
    }

    // ==================== 成本分摊记录 ====================

    private async Task RecordOrderCostAllocationAsync(int orderId, CancellationToken ct)
    {
        try
        {
            var orderNo = await _db.Orders
                .AsNoTracking()
                .Where(o => o.OrderId == orderId)
                .Select(o => o.OrderNo)
                .FirstOrDefaultAsync(ct);

            if (string.IsNullOrEmpty(orderNo)) return;

            var details = await (from od in _db.OrderDetails
                                 join p in _db.Products on od.ProductId equals p.ProductId into pJoin
                                 from p in pJoin.DefaultIfEmpty()
                                 where od.OrderId == orderId
                                 select new
                                 {
                                     od.DetailId,
                                     od.ProductId,
                                     ProductName = p != null ? p.ProductName : od.ProductName,
                                     od.Quantity,
                                     UnitCost = p != null ? (p.UnitCost ?? 0) : 0
                                 }).ToListAsync(ct);

            foreach (var d in details)
            {
                var unitCost = d.UnitCost > 0 ? d.UnitCost : await CalculateProductUnitCostAsync(d.ProductId, ct);
                var totalCost = unitCost * d.Quantity;

                if (totalCost > 0)
                {
                    var alloc = new OrderCostAllocation
                    {
                        OrderId = orderId,
                        OrderNo = orderNo,
                        CostType = "Product",
                        ItemCode = d.ProductId.ToString(),
                        ItemName = d.ProductName ?? "",
                        UnitCost = unitCost,
                        Quantity = d.Quantity,
                        TotalCost = totalCost,
                        AllocatedAt = DateTime.Now,
                        CreatedAt = DateTime.Now
                    };
                    _db.OrderCostAllocations.Add(alloc);
                }
            }

            await _db.SaveChangesAsync(ct);
        }
        catch
        {
            // 非关键路径，忽略错误
        }
    }

    // ==================== 成本摘要 ====================

    public async Task<CostSummaryDto> GetCostSummaryAsync(CancellationToken ct = default)
    {
        var totalProducts = await _db.Products.AsNoTracking().Where(p => p.IsActive == true).CountAsync(ct);
        var updatedProducts = await _db.Products.AsNoTracking().Where(p => p.IsActive == true && p.UnitCost > 0).CountAsync(ct);
        var updatedOrders = await _db.Orders.AsNoTracking().Where(o => o.CostAmount > 0).CountAsync(ct);
        var totalOrders = await _db.Orders.AsNoTracking().Where(o => o.Status != "Pending" && o.Status != "Cancelled").CountAsync(ct);

        return new CostSummaryDto
        {
            TotalProducts = totalProducts,
            UpdatedProducts = updatedProducts,
            TotalOrders = totalOrders,
            UpdatedOrders = updatedOrders,
            LastUpdate = DateTime.Now
        };
    }
}
