using System.Collections.Concurrent;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 成本自动传导引擎 — 三级成本计算 (原材料 → 香调 → 产品 → 订单)
/// 对应 ASP 中的 cost_engine.asp (45KB)
/// 使用 IMemoryCache 实现 CE_* 缓存键，对齐 V18 全局字典策略
/// </summary>
public class CostEngine : ICostEngine
{
    private readonly PerfumeShopContext _db;
    private readonly IMemoryCache _cache;

    // 请求级别缓存（同一请求内避免重复计算，对应 V18 的 CE_NoteCostCache / CE_ProductBOMCache / CE_ProductUnitCache）
    private readonly ConcurrentDictionary<int, decimal> _materialCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _noteCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _bomCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _unitCostCache = new();
    private readonly ConcurrentDictionary<int, decimal> _fixedBrandCostCache = new();
    private readonly ConcurrentDictionary<string, decimal> _batchCostCache = new();

    // IMemoryCache 缓存键前缀，对齐 V18 的 CE_* 命名
    private const string CE_MaterialPrices = "CE_MaterialPrices";       // Dictionary<int, decimal>
    private const string CE_BaseNotePrices = "CE_BaseNotePrices";       // Dictionary<int, decimal>
    private const string CE_AccordRecipes = "CE_AccordRecipes";         // Dictionary<int, (int AccordRecipeId, decimal BatchSize)>
    private const string CE_AccordMaterials = "CE_AccordMaterials";     // Dictionary<int, List<(int MatId, decimal Pct, decimal Qty)>>
    private const string CE_NoteIngredients = "CE_NoteIngredients";     // Dictionary<int, List<(int BaseNoteId, decimal Pct)>>
    private const string CE_NotePriceAdditions = "CE_NotePriceAdditions"; // Dictionary<int, decimal>
    private const string CE_ProductNoteRatios = "CE_ProductNoteRatios"; // Dictionary<int, List<(int NoteId, decimal Pct)>>
    private const string CE_BottleAdditions = "CE_BottleAdditions";     // Dictionary<int, decimal>
    private const string CE_ProductExtraCosts = "CE_ProductExtraCosts"; // Dictionary<int, (decimal Packaging, decimal Other)>
    private const string CE_WeightedBatchCosts = "CE_WeightedBatchCosts"; // Dictionary<string, decimal>
    private const string CE_FixedBrandCosts = "CE_FixedBrandCosts";     // Dictionary<int, decimal>
    private const string CE_Stats = "CE_Stats";                         // Dictionary<string, int>
    private const string CE_IsPreloaded = "CE_IsPreloaded";

    // 缓存过期时间：30 分钟（对齐 V18 页面级缓存策略）
    private static readonly TimeSpan CacheDuration = TimeSpan.FromMinutes(30);

    public CostEngine(PerfumeShopContext db, IMemoryCache cache)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
        _cache = cache ?? throw new ArgumentNullException(nameof(cache));
    }

    /// <summary>清除所有引擎缓存 — 预加载后重新计算时使用</summary>
    public void ClearCache()
    {
        _materialCostCache.Clear();
        _noteCostCache.Clear();
        _bomCostCache.Clear();
        _unitCostCache.Clear();
        _fixedBrandCostCache.Clear();
        _batchCostCache.Clear();
        _cache.Remove(CE_MaterialPrices);
        _cache.Remove(CE_BaseNotePrices);
        _cache.Remove(CE_AccordRecipes);
        _cache.Remove(CE_AccordMaterials);
        _cache.Remove(CE_NoteIngredients);
        _cache.Remove(CE_NotePriceAdditions);
        _cache.Remove(CE_ProductNoteRatios);
        _cache.Remove(CE_BottleAdditions);
        _cache.Remove(CE_ProductExtraCosts);
        _cache.Remove(CE_WeightedBatchCosts);
        _cache.Remove(CE_FixedBrandCosts);
        _cache.Remove(CE_Stats);
        _cache.Remove(CE_IsPreloaded);
    }

    // ==================== 缓存预加载 ====================

    /// <summary>
    /// 全局缓存预加载（原材料→香调→产品三层预热）— 对应 CE_PreloadAllCostData()
    /// 加载：原料价格、Accord配方、原料配比、成分聚合、香调配比、瓶身成本、包装/人工分摊
    /// 缓存键设计对齐 V18 的 CE_* 字典命名
    /// </summary>
    public async Task PreloadAllCostDataAsync(CancellationToken ct = default)
    {
        // === 1. 加载最新采购价 (ItemCode → UnitPrice) — 对应 CE_SPPrices ===
        var supplierPrices = await _db.SupplierPrices
            .AsNoTracking()
            .Where(sp => sp.IsActive == true)
            .GroupBy(sp => sp.ItemCode!)
            .Select(g => new
            {
                ItemCode = g.Key,
                UnitPrice = g.OrderByDescending(sp => sp.CreatedAt).First().UnitPrice
            })
            .ToDictionaryAsync(x => x.ItemCode, x => x.UnitPrice ?? 0m, ct);

        // === 2. 加载原料信息并构建 MaterialPrices (优先使用加权成本) ===
        var rawMaterials = await _db.RawMaterialInventories
            .AsNoTracking()
            .Where(m => m.StockQty > 0)
            .ToListAsync(ct);

        var materialPrices = new Dictionary<int, decimal>();
        foreach (var mat in rawMaterials)
        {
            if (mat.MaterialId <= 0) continue;
            decimal matPrice;
            var wuc = mat.WeightedUnitCost ?? 0;
            if (wuc > 0)
            {
                matPrice = wuc;
            }
            else if (!string.IsNullOrEmpty(mat.ItemCode) && supplierPrices.TryGetValue(mat.ItemCode, out var sp))
            {
                matPrice = sp;
            }
            else
            {
                matPrice = mat.UnitPrice ?? 0;
            }
            materialPrices.TryAdd(mat.MaterialId, matPrice);
        }
        _cache.Set(CE_MaterialPrices, materialPrices, CacheDuration);

        // === 2.5. 加载跨品类加权批次成本 ===
        await PreloadBatchCostsAsync(ct);

        // === 3. 加载 Accord 配方 (NoteID → AccordRecipeID|BatchSize) ===
        var accordRecipes = await _db.RecipeAccords
            .AsNoTracking()
            .Where(ra => ra.Status == "Published" && ra.NoteId != null)
            .GroupBy(ra => ra.NoteId!.Value)
            .Select(g => g.OrderByDescending(ra => ra.PublishedAt).First())
            .ToDictionaryAsync(
                ra => ra.NoteId!.Value,
                ra => (ra.AccordRecipeId, (decimal)(ra.BatchSize > 0 ? ra.BatchSize : 100)),
                ct);
        _cache.Set(CE_AccordRecipes, accordRecipes, CacheDuration);

        // === 4. 加载 Accord 原料配比 (AccordRecipeID → List) ===
        var accordMaterialsRaw = await _db.RecipeAccordMaterials
            .AsNoTracking()
            .OrderBy(ram => ram.AccordRecipeId)
            .ToListAsync(ct);

        var accordMaterials = new Dictionary<int, List<(int MatId, decimal Pct, decimal Qty)>>();
        foreach (var ram in accordMaterialsRaw)
        {
            if (ram.AccordRecipeId == null) continue;
            var key = ram.AccordRecipeId.Value;
            if (!accordMaterials.ContainsKey(key))
                accordMaterials[key] = new List<(int, decimal, decimal)>();
            accordMaterials[key].Add((ram.MaterialId ?? 0, (decimal)(ram.Percentage ?? 0), (decimal)(ram.PlannedQty ?? 0)));
        }
        _cache.Set(CE_AccordMaterials, accordMaterials, CacheDuration);

        // === 2.5 (V9): 加载 BaseNotes 单价 ===
        var baseNotePrices = await _db.BaseNotes
            .AsNoTracking()
            .Where(bn => bn.IsActive == true && bn.UnitPrice > 0)
            .ToDictionaryAsync(bn => bn.BaseNoteId, bn => bn.UnitPrice ?? 0m, ct);
        _cache.Set(CE_BaseNotePrices, baseNotePrices, CacheDuration);

        // === 5. 加载 NoteIngredients (NoteID → List) ===
        var noteIngredientsRaw = await _db.NoteIngredients
            .AsNoTracking()
            .OrderBy(ni => ni.NoteId)
            .ToListAsync(ct);

        var noteIngredients = new Dictionary<int, List<(int BaseNoteId, decimal Pct)>>();
        foreach (var ni in noteIngredientsRaw)
        {
            if (!noteIngredients.ContainsKey(ni.NoteId))
                noteIngredients[ni.NoteId] = new List<(int, decimal)>();
            noteIngredients[ni.NoteId].Add((ni.BaseNoteId, (decimal)(ni.Percentage ?? 0)));
        }
        _cache.Set(CE_NoteIngredients, noteIngredients, CacheDuration);

        // === 6. 加载 FragranceNotes PriceAddition ===
        var notePriceAdditions = await _db.FragranceNotes
            .AsNoTracking()
            .Where(fn => fn.IsActive == true)
            .ToDictionaryAsync(fn => fn.NoteId, fn => fn.PriceAddition ?? 0m, ct);
        _cache.Set(CE_NotePriceAdditions, notePriceAdditions, CacheDuration);

        // === 7. 加载 ProductNoteRatios (ProductID → List) ===
        var productNoteRatiosRaw = await _db.ProductNoteRatios
            .AsNoTracking()
            .OrderBy(pnr => pnr.ProductId)
            .ToListAsync(ct);

        var productNoteRatios = new Dictionary<int, List<(int NoteId, decimal Pct)>>();
        foreach (var pnr in productNoteRatiosRaw)
        {
            if (!productNoteRatios.ContainsKey(pnr.ProductId))
                productNoteRatios[pnr.ProductId] = new List<(int, decimal)>();
            productNoteRatios[pnr.ProductId].Add((pnr.NoteId, pnr.Percentage));
        }
        _cache.Set(CE_ProductNoteRatios, productNoteRatios, CacheDuration);

        // === 8. 加载 BottleStyles PriceAddition (ProductID → PriceAddition) ===
        var bottleAdditions = await (from pbs in _db.ProductBottleStyles
                                   join bs in _db.BottleStyles on pbs.BottleId equals bs.BottleId
                                   where bs.IsActive == true
                                   select new { pbs.ProductId, PriceAddition = bs.PriceAddition ?? 0m })
                                   .ToDictionaryAsync(x => x.ProductId, x => x.PriceAddition, ct);
        _cache.Set(CE_BottleAdditions, bottleAdditions, CacheDuration);

        // === 9. 加载 ProductCosts (ProductID → "PackagingCost|OtherCost") ===
        var productCostsRaw = await _db.ProductCosts
            .AsNoTracking()
            .GroupBy(pc => new { pc.ProductId, pc.CostType })
            .Select(g => new { g.Key.ProductId, g.Key.CostType, Total = g.Sum(x => x.TotalCost ?? 0) })
            .ToListAsync(ct);

        var productExtraCosts = new Dictionary<int, (decimal Packaging, decimal Other)>();
        foreach (var pc in productCostsRaw)
        {
            if (!productExtraCosts.ContainsKey(pc.ProductId))
                productExtraCosts[pc.ProductId] = (0, 0);
            var existing = productExtraCosts[pc.ProductId];
            if (pc.CostType == "Packaging")
                productExtraCosts[pc.ProductId] = (pc.Total, existing.Other);
            else if (pc.CostType == "Other")
                productExtraCosts[pc.ProductId] = (existing.Packaging, pc.Total);
        }
        _cache.Set(CE_ProductExtraCosts, productExtraCosts, CacheDuration);

        // === 9.5. 预加载品牌定香(Fixed)采购成本 ===
        await PreloadFixedBrandCostsAsync(ct);

        // === 10. 预加载统计数据 ===
        var stats = new Dictionary<string, int>();
        stats["updatedProducts"] = await _db.Products.AsNoTracking().CountAsync(p => p.IsActive == true && p.UnitCost > 0, ct);
        stats["totalProducts"] = await _db.Products.AsNoTracking().CountAsync(p => p.IsActive == true, ct);
        stats["updatedOrders"] = await _db.Orders.AsNoTracking().CountAsync(o => o.CostAmount > 0, ct);
        stats["totalValidOrders"] = await _db.Orders.AsNoTracking().CountAsync(o => o.Status != "Pending" && o.Status != "Cancelled", ct);
        stats["allOrders"] = await _db.Orders.AsNoTracking().CountAsync(ct);
        stats["rawMaterials"] = await _db.RawMaterialInventories.AsNoTracking().CountAsync(m => m.StockQty > 0, ct);
        _cache.Set(CE_Stats, stats, CacheDuration);

        _cache.Set(CE_IsPreloaded, true, CacheDuration);
    }

    /// <summary>
    /// 预加载跨品类加权批次成本 — 对应 CE_PreloadBatchCosts()
    /// 来源：RawMaterialInventory.WeightedUnitCost 及各类库存的 WeightedUnitCost
    /// 覆盖 5 个品类：RawMaterial, Packaging, Bottle, Printing, SprayHead
    /// </summary>
    public async Task PreloadBatchCostsAsync(CancellationToken ct = default)
    {
        var batchCosts = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        try
        {
            // 使用原始 SQL 的 UNION ALL 查询 5 个品类，避免 EF 模型缺少 WeightedUnitCost 字段
            // 对已有 WeightedUnitCost 映射的 RawMaterialInventory 使用 LINQ
            var rawMats = await _db.RawMaterialInventories
                .AsNoTracking()
                .Where(m => m.StockQty > 0 && m.ItemCode != null)
                .Select(m => new { m.ItemCode, WUC = m.WeightedUnitCost ?? m.UnitPrice ?? 0 })
                .ToListAsync(ct);

            foreach (var item in rawMats)
            {
                if (!string.IsNullOrEmpty(item.ItemCode) && item.WUC > 0)
                    batchCosts.TryAdd(item.ItemCode, item.WUC);
            }

            // 其他品类使用原生 SQL（模型可能缺少 WeightedUnitCost/IsActive 字段）
            var otherCategories = new[]
            {
                ("PackagingInventory", "StockQty > 0"),
                ("BottleStyles", "StockQty > 0"),
                ("PrintingInventory", "StockQty > 0"),
                ("SprayHeadInventory", "StockQty > 0")
            };

            foreach (var (table, where) in otherCategories)
            {
                try
                {
                    var sql = $"SELECT ISNULL(ItemCode,'') AS ItemCode, ISNULL(WeightedUnitCost, ISNULL(UnitPrice,0)) AS WUC FROM [{table}] WHERE {where}";
                    var results = await _db.Database.SqlQueryRaw<BatchCostRow>(sql).ToListAsync(ct);
                    foreach (var row in results)
                    {
                        if (!string.IsNullOrEmpty(row.ItemCode) && row.WUC > 0)
                            batchCosts.TryAdd(row.ItemCode, row.WUC);
                    }
                }
                catch
                {
                    // 表可能不存在（如 PrintingInventory / SprayHeadInventory），忽略
                }
            }
        }
        catch
        {
            // 批量加载失败时保持空字典
        }

        _cache.Set(CE_WeightedBatchCosts, batchCosts, CacheDuration);
    }

    /// <summary>预加载品牌定香(Fixed)采购成本 — ProductID → AvgUnitCost</summary>
    private async Task PreloadFixedBrandCostsAsync(CancellationToken ct = default)
    {
        var fixedCosts = new Dictionary<int, decimal>();
        try
        {
            var results = await (from fbp in _db.FixedBrandProducts
                                 join fbi in _db.FixedBrandInventories on fbp.FixedProductId equals fbi.FixedProductId into fbiJoin
                                 from fbi in fbiJoin.DefaultIfEmpty()
                                 where fbp.ProductId != null && fbp.ProductId > 0 && fbp.Status == "Active"
                                 select new
                                 {
                                     ProductId = fbp.ProductId!.Value,
                                     Cost = fbi != null && fbi.AvgUnitCost > 0 ? fbi.AvgUnitCost : fbp.UnitPrice
                                 }).ToListAsync(ct);

            foreach (var r in results)
            {
                if (r.Cost > 0)
                    fixedCosts.TryAdd(r.ProductId, r.Cost!.Value);
            }
        }
        catch
        {
            // FixedBrandProducts 表可能不存在
        }

        _cache.Set(CE_FixedBrandCosts, fixedCosts, CacheDuration);
    }

    // ==================== 缓存辅助：获取预加载字典 ====================

    private Dictionary<int, decimal>? GetMaterialPrices() =>
        _cache.TryGetValue<Dictionary<int, decimal>>(CE_MaterialPrices, out var d) ? d : null;

    private Dictionary<int, decimal>? GetBaseNotePrices() =>
        _cache.TryGetValue<Dictionary<int, decimal>>(CE_BaseNotePrices, out var d) ? d : null;

    private Dictionary<int, (int AccordRecipeId, decimal BatchSize)>? GetAccordRecipes() =>
        _cache.TryGetValue<Dictionary<int, (int, decimal)>>(CE_AccordRecipes, out var d) ? d : null;

    private Dictionary<int, List<(int MatId, decimal Pct, decimal Qty)>>? GetAccordMaterials() =>
        _cache.TryGetValue<Dictionary<int, List<(int, decimal, decimal)>>>(CE_AccordMaterials, out var d) ? d : null;

    private Dictionary<int, List<(int BaseNoteId, decimal Pct)>>? GetNoteIngredients() =>
        _cache.TryGetValue<Dictionary<int, List<(int, decimal)>>>(CE_NoteIngredients, out var d) ? d : null;

    private Dictionary<int, decimal>? GetNotePriceAdditions() =>
        _cache.TryGetValue<Dictionary<int, decimal>>(CE_NotePriceAdditions, out var d) ? d : null;

    private Dictionary<int, List<(int NoteId, decimal Pct)>>? GetProductNoteRatios() =>
        _cache.TryGetValue<Dictionary<int, List<(int, decimal)>>>(CE_ProductNoteRatios, out var d) ? d : null;

    private Dictionary<int, decimal>? GetBottleAdditions() =>
        _cache.TryGetValue<Dictionary<int, decimal>>(CE_BottleAdditions, out var d) ? d : null;

    private Dictionary<int, (decimal Packaging, decimal Other)>? GetProductExtraCosts() =>
        _cache.TryGetValue<Dictionary<int, (decimal, decimal)>>(CE_ProductExtraCosts, out var d) ? d : null;

    private Dictionary<string, decimal>? GetWeightedBatchCosts() =>
        _cache.TryGetValue<Dictionary<string, decimal>>(CE_WeightedBatchCosts, out var d) ? d : null;

    private Dictionary<int, decimal>? GetFixedBrandCosts() =>
        _cache.TryGetValue<Dictionary<int, decimal>>(CE_FixedBrandCosts, out var d) ? d : null;

    // ==================== Level 1: 原料成本 ====================

    /// <summary>
    /// 计算单个原料成本 (加权平均/供应商报价回退) — 对应 CE_CalculateMaterialCost()
    /// V10: 优先使用 WeightedUnitCost，回退到 SupplierPrices 最新报价，最后回退到库存 UnitPrice
    /// </summary>
    public async Task<decimal> CalculateMaterialCostAsync(int materialId, CancellationToken ct = default)
    {
        if (_materialCostCache.TryGetValue(materialId, out var cached))
            return cached;

        // 优先从预加载缓存获取
        var prices = GetMaterialPrices();
        if (prices != null && prices.TryGetValue(materialId, out var preloaded))
        {
            _materialCostCache.TryAdd(materialId, preloaded);
            return preloaded;
        }

        // 未预加载时回退到 DB 查询（对齐 V18 逻辑）
        decimal cost = 0;

        var material = await _db.RawMaterialInventories
            .AsNoTracking()
            .FirstOrDefaultAsync(m => m.MaterialId == materialId, ct);

        if (material != null)
        {
            if ((material.WeightedUnitCost ?? 0) > 0)
            {
                cost = material.WeightedUnitCost!.Value;
            }
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

            if (cost <= 0)
                cost = material.UnitPrice ?? 0;
        }

        _materialCostCache.TryAdd(materialId, cost);
        return cost;
    }

    // ==================== Level 2: 香调成本 ====================

    /// <summary>
    /// 计算单个香调成本 (Accord配方或成分聚合) — 对应 CE_CalculateNoteCost()
    /// 路径A：通过 RecipeAccords → RecipeAccordMaterials 计算
    /// 路径B：通过 NoteIngredients → BaseNotes → 成分聚合计算
    /// 兜底：FragranceNotes.PriceAddition
    /// </summary>
    public async Task<decimal> CalculateNoteCostAsync(int noteId, CancellationToken ct = default)
    {
        if (_noteCostCache.TryGetValue(noteId, out var cached))
            return cached;

        decimal totalCost = 0;
        bool hasAccord = false;

        // 尝试从预加载缓存获取 Accord 配方
        var accordRecipes = GetAccordRecipes();
        var accordMaterials = GetAccordMaterials();
        var materialPrices = GetMaterialPrices();

        if (accordRecipes != null && accordRecipes.TryGetValue(noteId, out var recipe))
        {
            hasAccord = true;
            var batchSize = recipe.BatchSize > 0 ? recipe.BatchSize : 100m;

            if (accordMaterials != null && accordMaterials.TryGetValue(recipe.AccordRecipeId, out var mats))
            {
                foreach (var (matId, pct, _) in mats)
                {
                    var matUnitCost = materialPrices != null && materialPrices.TryGetValue(matId, out var mp)
                        ? mp
                        : await CalculateMaterialCostAsync(matId, ct);
                    totalCost += (pct / batchSize) * matUnitCost;
                }
            }
        }

        // 路径B: BaseNote 成分聚合
        if (!hasAccord)
        {
            var noteIngredients = GetNoteIngredients();
            var baseNotePrices = GetBaseNotePrices();
            var notePriceAdditions = GetNotePriceAdditions();

            if (noteIngredients != null && noteIngredients.TryGetValue(noteId, out var ingredients))
            {
                foreach (var (baseNoteId, pct) in ingredients)
                {
                    decimal baseNoteCost = 0;

                    // 递归计算基香成本（优先缓存）
                    baseNoteCost = _noteCostCache.TryGetValue(baseNoteId, out var nc) ? nc : 0;
                    if (baseNoteCost <= 0 && noteIngredients != null && noteIngredients.ContainsKey(baseNoteId))
                    {
                        baseNoteCost = await CalculateNoteCostAsync(baseNoteId, ct);
                    }

                    // V9: 优先从 BaseNotes 单价获取
                    if (baseNoteCost <= 0 && baseNotePrices != null && baseNotePrices.TryGetValue(baseNoteId, out var bnp))
                        baseNoteCost = bnp;

                    // 回退到 PriceAddition
                    if (baseNoteCost <= 0 && notePriceAdditions != null && notePriceAdditions.TryGetValue(baseNoteId, out var npa))
                        baseNoteCost = npa;

                    // 最终回退到 DB 查询
                    if (baseNoteCost <= 0)
                    {
                        baseNoteCost = await _db.BaseNotes
                            .AsNoTracking()
                            .Where(bn => bn.BaseNoteId == baseNoteId && bn.IsActive == true)
                            .Select(bn => bn.UnitPrice ?? 0)
                            .FirstOrDefaultAsync(ct);
                    }
                    if (baseNoteCost <= 0)
                    {
                        baseNoteCost = await _db.FragranceNotes
                            .AsNoTracking()
                            .Where(fn => fn.NoteId == baseNoteId)
                            .Select(fn => fn.PriceAddition ?? 0)
                            .FirstOrDefaultAsync(ct);
                    }

                    totalCost += baseNoteCost * pct / 100m;
                }
            }
            else
            {
                // 未预加载时回退到 DB
                var dbIngredients = await _db.NoteIngredients
                    .AsNoTracking()
                    .Where(ni => ni.NoteId == noteId)
                    .ToListAsync(ct);

                if (dbIngredients.Count > 0)
                {
                    foreach (var ing in dbIngredients)
                    {
                        var baseNoteCost = await CalculateNoteCostAsync(ing.BaseNoteId, ct);
                        if (baseNoteCost <= 0)
                        {
                            baseNoteCost = await _db.BaseNotes
                                .AsNoTracking()
                                .Where(bn => bn.BaseNoteId == ing.BaseNoteId && bn.IsActive == true)
                                .Select(bn => bn.UnitPrice ?? 0)
                                .FirstOrDefaultAsync(ct);
                        }
                        if (baseNoteCost <= 0)
                        {
                            baseNoteCost = await _db.FragranceNotes
                                .AsNoTracking()
                                .Where(fn => fn.NoteId == ing.BaseNoteId)
                                .Select(fn => fn.PriceAddition ?? 0)
                                .FirstOrDefaultAsync(ct);
                        }
                        totalCost += baseNoteCost * (decimal)(ing.Percentage ?? 0) / 100m;
                    }
                }
            }
        }

        // 兜底: FragranceNotes.PriceAddition
        if (totalCost <= 0)
        {
            var notePriceAdditions = GetNotePriceAdditions();
            if (notePriceAdditions != null && notePriceAdditions.TryGetValue(noteId, out var fallback))
            {
                totalCost = fallback;
            }
            else
            {
                totalCost = await _db.FragranceNotes
                    .AsNoTracking()
                    .Where(fn => fn.NoteId == noteId)
                    .Select(fn => fn.PriceAddition ?? 0)
                    .FirstOrDefaultAsync(ct);
            }
        }

        _noteCostCache.TryAdd(noteId, totalCost);
        return totalCost;
    }

    // ==================== Level 3: 产品 BOM 成本 ====================

    /// <summary>
    /// 计算产品 BOM 成本 (香调配比 + 瓶身) — 对应 CE_CalculateProductBOMCost()
    /// 来源：ProductNoteRatios × Note成本 + BottleStyles.PriceAddition
    /// </summary>
    public async Task<decimal> CalculateProductBomCostAsync(int productId, CancellationToken ct = default)
    {
        if (_bomCostCache.TryGetValue(productId, out var cached))
            return cached;

        decimal totalCost = 0;

        // 1. 香调配比成本
        var productNoteRatios = GetProductNoteRatios();
        if (productNoteRatios != null && productNoteRatios.TryGetValue(productId, out var ratios))
        {
            foreach (var (noteId, pct) in ratios)
            {
                var noteCost = await CalculateNoteCostAsync(noteId, ct);
                totalCost += noteCost * pct / 100m;
            }
        }
        else
        {
            // 未预加载时回退到 DB
            var dbRatios = await _db.ProductNoteRatios
                .AsNoTracking()
                .Where(pnr => pnr.ProductId == productId)
                .ToListAsync(ct);

            foreach (var nr in dbRatios)
            {
                var noteCost = await CalculateNoteCostAsync(nr.NoteId, ct);
                totalCost += noteCost * nr.Percentage / 100m;
            }
        }

        // 2. 瓶身成本
        var bottleAdditions = GetBottleAdditions();
        if (bottleAdditions != null && bottleAdditions.TryGetValue(productId, out var bottleCost))
        {
            totalCost += bottleCost;
        }
        else
        {
            var dbBottleCost = await (from pbs in _db.ProductBottleStyles
                                      join bs in _db.BottleStyles on pbs.BottleId equals bs.BottleId
                                      where pbs.ProductId == productId && bs.IsActive == true
                                      select bs.PriceAddition ?? 0)
                                      .FirstOrDefaultAsync(ct);
            totalCost += dbBottleCost;
        }

        _bomCostCache.TryAdd(productId, totalCost);
        return totalCost;
    }

    // ==================== Level 4: 产品完整单位成本 ====================

    /// <summary>
    /// 计算产品完整单位成本 (BOM + 包装 + 人工) — 对应 CE_CalculateProductUnitCost()
    /// 品牌定香(standard/Fixed)产品：优先使用采购加权平均成本，回退到 Products.UnitCost
    /// </summary>
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

        // 从预加载缓存或 DB 获取包装/其他成本
        var extraCosts = GetProductExtraCosts();
        if (extraCosts != null && extraCosts.TryGetValue(productId, out var extras))
        {
            totalCost += extras.Packaging + extras.Other;
        }
        else
        {
            var packagingCost = await _db.ProductCosts
                .AsNoTracking()
                .Where(pc => pc.ProductId == productId && pc.CostType == "Packaging")
                .SumAsync(pc => pc.TotalCost ?? 0, ct);

            var otherCost = await _db.ProductCosts
                .AsNoTracking()
                .Where(pc => pc.ProductId == productId && pc.CostType == "Other")
                .SumAsync(pc => pc.TotalCost ?? 0, ct);

            totalCost += packagingCost + otherCost;
        }

        _unitCostCache.TryAdd(productId, totalCost);
        return totalCost;
    }

    // ==================== 缓存查询方法（O(1) 内存查找） ====================

    /// <summary>从缓存获取原料成本 — 对应 CE_GetCachedMaterialCost()</summary>
    public Task<decimal> GetCachedMaterialCostAsync(int materialId, CancellationToken ct = default)
    {
        var prices = GetMaterialPrices();
        if (prices != null && prices.TryGetValue(materialId, out var cost))
            return Task.FromResult(cost);
        if (_materialCostCache.TryGetValue(materialId, out var req))
            return Task.FromResult(req);
        return Task.FromResult(0m);
    }

    /// <summary>从缓存获取香调成本 — 对应 CE_GetCachedNoteCost()，首次计算后缓存</summary>
    public async Task<decimal> GetCachedNoteCostAsync(int noteId, CancellationToken ct = default)
    {
        if (_noteCostCache.TryGetValue(noteId, out var cached))
            return cached;
        return await CalculateNoteCostAsync(noteId, ct);
    }

    /// <summary>从缓存获取产品BOM成本 — 对应 CE_GetCachedProductBOMCost()</summary>
    public async Task<decimal> GetCachedProductBomCostAsync(int productId, CancellationToken ct = default)
    {
        if (_bomCostCache.TryGetValue(productId, out var cached))
            return cached;
        return await CalculateProductBomCostAsync(productId, ct);
    }

    /// <summary>从缓存获取产品单位总成本 — 对应 CE_GetCachedProductUnitCost()</summary>
    public async Task<decimal> GetCachedProductUnitCostAsync(int productId, CancellationToken ct = default)
    {
        if (_unitCostCache.TryGetValue(productId, out var cached))
            return cached;
        return await CalculateProductUnitCostAsync(productId, ct);
    }

    /// <summary>
    /// 从缓存获取品牌定香采购成本 — 对应 CE_GetCachedFixedBrandCost()
    /// 优先使用 FixedBrandInventory.AvgUnitCost，回退到 FixedBrandProducts.UnitPrice
    /// </summary>
    public async Task<decimal> GetFixedBrandCostAsync(int productId, CancellationToken ct = default)
    {
        if (_fixedBrandCostCache.TryGetValue(productId, out var cached))
            return cached;

        // 从预加载缓存获取
        var fixedCosts = GetFixedBrandCosts();
        if (fixedCosts != null && fixedCosts.TryGetValue(productId, out var preloaded))
        {
            _fixedBrandCostCache.TryAdd(productId, preloaded);
            return preloaded;
        }

        // 缓存未命中，实时查询
        decimal fbCost = 0;
        try
        {
            var result = await (from fbp in _db.FixedBrandProducts
                                join fbi in _db.FixedBrandInventories on fbp.FixedProductId equals fbi.FixedProductId into fbiJoin
                                from fbi in fbiJoin.DefaultIfEmpty()
                                where fbp.ProductId == productId && fbp.Status == "Active"
                                select fbi != null && fbi.AvgUnitCost > 0 ? fbi.AvgUnitCost : fbp.UnitPrice)
                                .FirstOrDefaultAsync(ct);
            fbCost = result ?? 0;
        }
        catch
        {
            // 表可能不存在
        }

        _fixedBrandCostCache.TryAdd(productId, fbCost);
        return fbCost;
    }

    // ==================== 批次加权成本 ====================

    /// <summary>获取指定物料的加权批次成本（从缓存） — 对应 CE_GetCachedBatchCost()</summary>
    public Task<decimal> GetBatchWeightedCostAsync(string itemCode, CancellationToken ct = default)
    {
        if (_batchCostCache.TryGetValue(itemCode, out var cached))
            return Task.FromResult(cached);

        var batchCosts = GetWeightedBatchCosts();
        if (batchCosts != null && batchCosts.TryGetValue(itemCode, out var cost))
        {
            _batchCostCache.TryAdd(itemCode, cost);
            return Task.FromResult(cost);
        }

        return Task.FromResult(0m);
    }

    /// <summary>
    /// 获取物料加权平均成本(查询5个品类库存表) — 对应 CE_GetBatchWeightedCost()
    /// 按优先级依次查询: RawMaterial → Packaging → Bottle → Printing → SprayHead
    /// </summary>
    public async Task<decimal> CalculateBatchWeightedCostAsync(string itemCode, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(itemCode)) return 0;

        // 先从缓存获取
        var cachedCost = await GetBatchWeightedCostAsync(itemCode, ct);
        if (cachedCost > 0) return cachedCost;

        // 未命中缓存，按优先级查询 5 个品类
        decimal cost = 0;

        // 1. RawMaterialInventory
        var rawResult = await _db.RawMaterialInventories
            .AsNoTracking()
            .Where(m => m.ItemCode == itemCode)
            .Select(m => m.WeightedUnitCost ?? m.UnitPrice ?? 0)
            .FirstOrDefaultAsync(ct);
        if (rawResult > 0) return rawResult;

        // 2-5. 使用原生 SQL 查询其他品类
        var tables = new[]
        {
            ("PackagingInventory", "StockQty > 0"),
            ("BottleStyles", "StockQty > 0"),
            ("PrintingInventory", "StockQty > 0"),
            ("SprayHeadInventory", "StockQty > 0")
        };

        foreach (var (table, _) in tables)
        {
            try
            {
                var sql = $"SELECT TOP 1 ISNULL(WeightedUnitCost, UnitPrice) AS WUC FROM [{table}] WHERE ItemCode = {{0}}";
                var result = await _db.Database.SqlQueryRaw<BatchCostRow>(sql, itemCode).FirstOrDefaultAsync(ct);
                if (result != null && result.WUC > 0)
                {
                    cost = result.WUC;
                    _batchCostCache.TryAdd(itemCode, cost);
                    return cost;
                }
            }
            catch
            {
                // 表可能不存在
            }
        }

        return cost;
    }

    // ==================== Level 5-6: 更新产品成本 ====================

    /// <summary>更新指定产品的成本到数据库 — 对应 CE_UpdateProductCost()</summary>
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

    /// <summary>批量更新所有启用产品的成本 — 对应 CE_UpdateAllProductCosts()</summary>
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

    // ==================== Level 7-8: 更新订单成本与利润 ====================

    /// <summary>
    /// 更新订单的成本和利润 — 对应 CE_UpdateOrderCosts()
    /// CostAmount = 各商品数量 × 单位成本
    /// ProfitAmount = TotalAmount - CostAmount - ShippingFee
    /// </summary>
    public async Task UpdateOrderCostsAsync(int orderId, CancellationToken ct = default)
    {
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

        var orderData = await _db.Orders
            .AsNoTracking()
            .Where(o => o.OrderId == orderId)
            .Select(o => new { o.TotalAmount, o.ShippingFee, o.ExpenseAmount })
            .FirstOrDefaultAsync(ct);

        if (orderData == null) return;

        // V21: 利润纳入费用分摊：Profit = Total - Cost - ShippingFee - ExpenseAmount（下限 0）
        var profitAmount = orderData.TotalAmount - orderCost - (orderData.ShippingFee ?? 0) - (orderData.ExpenseAmount ?? 0);
        if (profitAmount < 0) profitAmount = 0;

        await _db.Orders
            .Where(o => o.OrderId == orderId)
            .ExecuteUpdateAsync(setters => setters
                .SetProperty(o => o.CostAmount, orderCost)
                .SetProperty(o => o.ProfitAmount, profitAmount), ct);

        // V10: 记录成本分摊明细
        await RecordOrderCostAllocationAsync(orderId, ct);
    }

    /// <summary>批量更新所有有效订单的成本和利润 — 对应 CE_UpdateAllOrderCosts()</summary>
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

    // ==================== Level 9: 成本摘要 ====================

    /// <summary>获取成本传导状态摘要 — 对应 CE_GetCostSummary()</summary>
    public async Task<CostSummaryDto> GetCostSummaryAsync(CancellationToken ct = default)
    {
        // 优先从预加载统计缓存获取
        var stats = _cache.TryGetValue<Dictionary<string, int>>(CE_Stats, out var s) ? s : null;

        int totalProducts, updatedProducts, updatedOrders, totalOrders;
        if (stats != null)
        {
            totalProducts = stats.GetValueOrDefault("totalProducts");
            updatedProducts = stats.GetValueOrDefault("updatedProducts");
            updatedOrders = stats.GetValueOrDefault("updatedOrders");
            totalOrders = stats.GetValueOrDefault("totalValidOrders");
        }
        else
        {
            totalProducts = await _db.Products.AsNoTracking().CountAsync(p => p.IsActive == true, ct);
            updatedProducts = await _db.Products.AsNoTracking().CountAsync(p => p.IsActive == true && p.UnitCost > 0, ct);
            updatedOrders = await _db.Orders.AsNoTracking().CountAsync(o => o.CostAmount > 0, ct);
            totalOrders = await _db.Orders.AsNoTracking().CountAsync(o => o.Status != "Pending" && o.Status != "Cancelled", ct);
        }

        var matPrices = GetMaterialPrices();
        var batchCosts = GetWeightedBatchCosts();

        return new CostSummaryDto
        {
            TotalProducts = totalProducts,
            UpdatedProducts = updatedProducts,
            TotalOrders = totalOrders,
            UpdatedOrders = updatedOrders,
            LastUpdate = DateTime.Now,
            CachedMaterials = matPrices?.Count ?? _materialCostCache.Count,
            CachedBatchCosts = batchCosts?.Count ?? _batchCostCache.Count
        };
    }

    // ==================== 成本分摊记录 ====================

    /// <summary>记录订单成本分摊到 OrderCostAllocation 表 — 对应 CE_RecordOrderCostAllocation()</summary>
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

            // V19.6: 先清除该订单旧分摊，避免重复执行"自动更新订单利润"时累积重复行
            await _db.OrderCostAllocations.Where(a => a.OrderId == orderId).ExecuteDeleteAsync(ct);

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

    // ==================== 辅助类型 ====================

    /// <summary>用于原生 SQL 查询的批次成本行映射</summary>
    internal class BatchCostRow
    {
        public string? ItemCode { get; set; }
        public decimal WUC { get; set; }
    }
}
