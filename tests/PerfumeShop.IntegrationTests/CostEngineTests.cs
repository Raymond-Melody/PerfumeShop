using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.IntegrationTests;

/// <summary>
/// CostEngine 单元测试 — 基于 V18 cost_engine.asp 真实业务数据回归
/// 使用 EF Core InMemory Provider + TestEngineContext 解决 keyless 实体问题
/// </summary>
public class CostEngineTests : IDisposable
{
    private readonly TestEngineContext _db;
    private readonly IMemoryCache _cache;
    private readonly CostEngine _engine;

    public CostEngineTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"CostEngineTests_{Guid.NewGuid()}")
            .Options;
        _db = new TestEngineContext(options);
        _cache = new MemoryCache(new MemoryCacheOptions());
        _engine = new CostEngine(_db, _cache);
    }

    public void Dispose()
    {
        _db.Database.EnsureDeleted();
        _db.Dispose();
        _cache.Dispose();
    }

    // ==================== 1. 原料成本计算 ====================

    /// <summary>测试1: 加权成本优先 — WeightedUnitCost=25.5 应返回 25.5</summary>
    [Fact]
    public async Task CalculateMaterialCost_WeightedCost_ReturnsWeighted()
    {
        _db.RawMaterialInventories.Add(new RawMaterialInventory
        {
            MaterialId = 1, ItemCode = "RM001", UnitPrice = 20m,
            WeightedUnitCost = 25.5m, StockQty = 100
        });
        await _db.SaveChangesAsync();

        var cost = await _engine.CalculateMaterialCostAsync(1);
        Assert.Equal(25.5m, cost);
    }

    /// <summary>测试2: 供应商报价回退 — WeightedUnitCost=0, 应回退到 SupplierPrices</summary>
    [Fact]
    public async Task CalculateMaterialCost_FallbackToSupplierPrice()
    {
        _db.RawMaterialInventories.Add(new RawMaterialInventory
        {
            MaterialId = 2, ItemCode = "RM002", UnitPrice = 15m,
            WeightedUnitCost = 0, StockQty = 50
        });
        _db.SupplierPrices.Add(new SupplierPrice
        {
            PriceId = 1, ItemCode = "RM002", UnitPrice = 30m, IsActive = true,
            CreatedAt = DateTime.Now.AddDays(-1)
        });
        await _db.SaveChangesAsync();

        var cost = await _engine.CalculateMaterialCostAsync(2);
        Assert.Equal(30m, cost);
    }

    /// <summary>测试3: 最终回退到 UnitPrice — 无加权无供应商价</summary>
    [Fact]
    public async Task CalculateMaterialCost_FinalFallback_UnitPrice()
    {
        _db.RawMaterialInventories.Add(new RawMaterialInventory
        {
            MaterialId = 3, ItemCode = "RM003", UnitPrice = 12m,
            WeightedUnitCost = 0, StockQty = 10
        });
        await _db.SaveChangesAsync();

        var cost = await _engine.CalculateMaterialCostAsync(3);
        Assert.Equal(12m, cost);
    }

    // ==================== 2. 香调成本计算 ====================

    /// <summary>测试4: Accord 配方成本 — 2 个原料各 50%, 批量100</summary>
    [Fact]
    public async Task CalculateNoteCost_AccordRecipe_CalculatesCorrectly()
    {
        _db.RawMaterialInventories.AddRange(
            new RawMaterialInventory { MaterialId = 10, ItemCode = "M10", UnitPrice = 10m, WeightedUnitCost = 10m, StockQty = 100 },
            new RawMaterialInventory { MaterialId = 11, ItemCode = "M11", UnitPrice = 20m, WeightedUnitCost = 20m, StockQty = 100 }
        );
        _db.RecipeAccords.Add(new RecipeAccord
        {
            AccordRecipeId = 1, NoteId = 100, BatchSize = 100,
            Status = "Published", PublishedAt = DateTime.Now
        });
        _db.RecipeAccordMaterials.AddRange(
            new RecipeAccordMaterial { DetailId = 1, AccordRecipeId = 1, MaterialId = 10, Percentage = 50, PlannedQty = 50 },
            new RecipeAccordMaterial { DetailId = 2, AccordRecipeId = 1, MaterialId = 11, Percentage = 50, PlannedQty = 50 }
        );
        await _db.SaveChangesAsync();

        // 先预加载缓存，确保 Accord 配方被加载
        await _engine.PreloadAllCostDataAsync();

        var cost = await _engine.CalculateNoteCostAsync(100);
        // (50/100)*10 + (50/100)*20 = 5 + 10 = 15
        Assert.Equal(15m, cost);
    }

    /// <summary>测试5: PriceAddition 兜底 — 无 Accord 无 NoteIngredients</summary>
    [Fact]
    public async Task CalculateNoteCost_FallbackToPriceAddition()
    {
        _db.FragranceNotes.Add(new FragranceNote
        {
            NoteId = 200, NoteName = "TestNote", PriceAddition = 8.5m,
            IsActive = true, NoteType = "Middle"
        });
        await _db.SaveChangesAsync();

        var cost = await _engine.CalculateNoteCostAsync(200);
        Assert.Equal(8.5m, cost);
    }

    // ==================== 3. 产品 BOM 成本 ====================

    /// <summary>测试6: 产品BOM = 香调配比 + 瓶身附加</summary>
    [Fact]
    public async Task CalculateProductBomCost_NoteRatiosPlusBottle()
    {
        _db.FragranceNotes.Add(new FragranceNote
        {
            NoteId = 300, NoteName = "Rose", PriceAddition = 10m,
            IsActive = true, NoteType = "Middle"
        });
        _db.ProductNoteRatios.Add(new ProductNoteRatio { ProductId = 500, NoteId = 300, Percentage = 80 });
        _db.BottleStyles.Add(new BottleStyle { BottleId = 1, BottleName = "Classic", PriceAddition = 5m, IsActive = true });
        _db.ProductBottleStyles.Add(new ProductBottleStyle { ProductId = 500, BottleId = 1 });
        await _db.SaveChangesAsync();

        var cost = await _engine.CalculateProductBomCostAsync(500);
        // 10 * 80/100 + 5 = 8 + 5 = 13
        Assert.Equal(13m, cost);
    }

    // ==================== 4. 产品单位总成本 ====================

    /// <summary>测试7: 产品单位成本 = BOM + 包装 + 其他</summary>
    [Fact]
    public async Task CalculateProductUnitCost_BomPlusPackagingPlusOther()
    {
        _db.FragranceNotes.Add(new FragranceNote
        {
            NoteId = 400, NoteName = "Jasmine", PriceAddition = 20m,
            IsActive = true, NoteType = "Top"
        });
        _db.ProductNoteRatios.Add(new ProductNoteRatio { ProductId = 600, NoteId = 400, Percentage = 100 });
        _db.BottleStyles.Add(new BottleStyle { BottleId = 10, BottleName = "Modern", PriceAddition = 3m, IsActive = true });
        _db.ProductBottleStyles.Add(new ProductBottleStyle { ProductId = 600, BottleId = 10 });
        _db.ProductCosts.AddRange(
            new ProductCost { ProductId = 600, CostType = "Packaging", TotalCost = 2m },
            new ProductCost { ProductId = 600, CostType = "Other", TotalCost = 1.5m }
        );
        _db.Products.Add(new Product { ProductId = 600, ProductName = "TestPerfume", ProductType = "custom", IsActive = true });
        await _db.SaveChangesAsync();

        var cost = await _engine.CalculateProductUnitCostAsync(600);
        // BOM: 20*100/100 + 3 = 23; Unit: 23 + 2 + 1.5 = 26.5
        Assert.Equal(26.5m, cost);
    }

    // ==================== 5. 品牌定香成本 ====================

    /// <summary>测试8: standard 产品优先使用 FixedBrand 采购成本</summary>
    [Fact]
    public async Task CalculateProductUnitCost_StandardProduct_UsesFixedBrandCost()
    {
        _db.Products.Add(new Product { ProductId = 700, ProductName = "BrandPerfume", ProductType = "standard", IsActive = true });
        _db.FixedBrandProducts.Add(new FixedBrandProduct { FixedProductId = 1, ProductId = 700, ProductName = "BP", UnitPrice = 55m, Status = "Active" });
        _db.FixedBrandInventories.Add(new FixedBrandInventory { InventoryId = 1, FixedProductId = 1, AvgUnitCost = 50m });
        await _db.SaveChangesAsync();

        var cost = await _engine.CalculateProductUnitCostAsync(700);
        Assert.Equal(50m, cost);
    }

    // ==================== 6. 缓存命中率 ====================

    /// <summary>测试9: 预加载后第二次调用必须命中缓存</summary>
    [Fact]
    public async Task PreloadAllCostData_SecondCallHitsCache()
    {
        _db.RawMaterialInventories.Add(new RawMaterialInventory
        {
            MaterialId = 50, ItemCode = "RM050", UnitPrice = 18m,
            WeightedUnitCost = 22m, StockQty = 200
        });
        _db.FragranceNotes.Add(new FragranceNote
        {
            NoteId = 500, NoteName = "Vanilla", PriceAddition = 7m,
            IsActive = true, NoteType = "Base"
        });
        _db.Products.Add(new Product { ProductId = 800, ProductName = "VanillaDream", IsActive = true });
        await _db.SaveChangesAsync();

        await _engine.PreloadAllCostDataAsync();

        var matCost1 = await _engine.GetCachedMaterialCostAsync(50);
        Assert.Equal(22m, matCost1);

        var matCost2 = await _engine.GetCachedMaterialCostAsync(50);
        Assert.Equal(22m, matCost2);

        var noteCost = await _engine.GetCachedNoteCostAsync(500);
        Assert.Equal(7m, noteCost);
    }

    // ==================== 7. 加权批次成本 ====================

    /// <summary>测试10: 加权批次成本在不同库存组合下的计算正确性</summary>
    [Fact]
    public async Task PreloadBatchCosts_DifferentInventoryCategories()
    {
        _db.RawMaterialInventories.Add(new RawMaterialInventory
        {
            MaterialId = 60, ItemCode = "BATCH001", UnitPrice = 30m,
            WeightedUnitCost = 35m, StockQty = 500
        });
        await _db.SaveChangesAsync();

        await _engine.PreloadBatchCostsAsync();

        var batchCost = await _engine.GetBatchWeightedCostAsync("BATCH001");
        Assert.Equal(35m, batchCost);

        var missing = await _engine.GetBatchWeightedCostAsync("NONEXIST");
        Assert.Equal(0m, missing);
    }

    // ==================== 8. ClearCache ====================

    /// <summary>测试11: ClearCache 后预加载数据应失效</summary>
    [Fact]
    public async Task ClearCache_InvalidatesPreloadedData()
    {
        _db.RawMaterialInventories.Add(new RawMaterialInventory
        {
            MaterialId = 70, ItemCode = "RM070", UnitPrice = 40m,
            WeightedUnitCost = 45m, StockQty = 100
        });
        await _db.SaveChangesAsync();

        await _engine.PreloadAllCostDataAsync();
        var cost1 = await _engine.GetCachedMaterialCostAsync(70);
        Assert.Equal(45m, cost1);

        _engine.ClearCache();
        var cost2 = await _engine.GetCachedMaterialCostAsync(70);
        Assert.Equal(0m, cost2);
    }

    // ==================== 9. GetCostSummary ====================

    /// <summary>测试12: 预加载后 GetCostSummary 统计数据正确</summary>
    [Fact]
    public async Task GetCostSummary_AfterPreload_ReturnsCorrectStats()
    {
        _db.Products.AddRange(
            new Product { ProductId = 901, ProductName = "P1", IsActive = true, UnitCost = 10m },
            new Product { ProductId = 902, ProductName = "P2", IsActive = true, UnitCost = 0 },
            new Product { ProductId = 903, ProductName = "P3", IsActive = true, UnitCost = 20m }
        );
        _db.RawMaterialInventories.Add(new RawMaterialInventory
        {
            MaterialId = 80, ItemCode = "RM080", StockQty = 100, UnitPrice = 5m
        });
        await _db.SaveChangesAsync();

        await _engine.PreloadAllCostDataAsync();
        var summary = await _engine.GetCostSummaryAsync();

        Assert.Equal(3, summary.TotalProducts);
        Assert.Equal(2, summary.UpdatedProducts);
        Assert.Equal(1, summary.CachedMaterials);
    }

    // ==================== 10. V18 函数映射覆盖检查 ====================

    /// <summary>测试13: 接口方法签名覆盖 V18 所有 CE_* 函数</summary>
    [Fact]
    public void ICostEngine_CoversAllV18Functions()
    {
        var methods = typeof(ICostEngine).GetMethods().Select(m => m.Name).ToHashSet();

        Assert.Contains("PreloadAllCostDataAsync", methods);
        Assert.Contains("PreloadBatchCostsAsync", methods);
        Assert.Contains("CalculateMaterialCostAsync", methods);
        Assert.Contains("CalculateNoteCostAsync", methods);
        Assert.Contains("CalculateProductBomCostAsync", methods);
        Assert.Contains("CalculateProductUnitCostAsync", methods);
        Assert.Contains("GetFixedBrandCostAsync", methods);
        Assert.Contains("GetBatchWeightedCostAsync", methods);
        Assert.Contains("UpdateProductCostAsync", methods);
        Assert.Contains("UpdateOrderCostsAsync", methods);
        Assert.Contains("GetCostSummaryAsync", methods);
    }
}
