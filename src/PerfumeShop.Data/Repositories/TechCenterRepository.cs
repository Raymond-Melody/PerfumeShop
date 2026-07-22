using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 技术中心仓储实现 — 对应 V18 dal_techcenter.asp 全部 DAL 方法
/// </summary>
public class TechCenterRepository : ITechCenterRepository
{
    private readonly PerfumeShopContext _context;

    public TechCenterRepository(PerfumeShopContext context)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
    }

    // ========== 配方 Recipes ==========

    public async Task<IEnumerable<Recipe>> GetActiveRecipesAsync(CancellationToken ct = default)
        => await _context.Recipes.AsNoTracking()
            .Where(r => r.IsActive == true)
            .OrderBy(r => r.RecipeCode)
            .ToListAsync(ct);

    public async Task<Recipe?> GetRecipeByIdAsync(int recipeId, CancellationToken ct = default)
        => await _context.Recipes.FindAsync(new object[] { recipeId }, ct);

    public async Task<Recipe> CreateRecipeAsync(Recipe recipe, CancellationToken ct = default)
    {
        recipe.CreatedAt = DateTime.Now;
        var entry = await _context.Recipes.AddAsync(recipe, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task UpdateRecipeAsync(Recipe recipe, CancellationToken ct = default)
    {
        recipe.UpdatedAt = DateTime.Now;
        _context.Recipes.Update(recipe);
        await _context.SaveChangesAsync(ct);
    }

    public async Task<int> CountActiveRecipesAsync(CancellationToken ct = default)
        => await _context.Recipes.CountAsync(r => r.IsActive == true, ct);

    // ========== 配方成分 RecipeIngredients ==========

    public async Task<int> CountIngredientsAsync(int recipeId, CancellationToken ct = default)
        => await _context.RecipeIngredients.CountAsync(i => i.RecipeId == recipeId, ct);

    public async Task<int> CountIngredientsByNoteAsync(int recipeId, int noteId, CancellationToken ct = default)
        => await _context.RecipeIngredients.CountAsync(i => i.RecipeId == recipeId && i.NoteId == noteId, ct);

    public async Task<IEnumerable<RecipeIngredient>> GetAccordIngredientsAsync(int recipeId, int noteId, CancellationToken ct = default)
        => await _context.RecipeIngredients.AsNoTracking()
            .Where(i => i.RecipeId == recipeId && i.NoteId == noteId)
            .OrderBy(i => i.Id)
            .ToListAsync(ct);

    public async Task AddIngredientAsync(RecipeIngredient ingredient, CancellationToken ct = default)
        => await _context.RecipeIngredients.AddAsync(ingredient, ct);

    public async Task DeleteIngredientAsync(int id, CancellationToken ct = default)
    {
        var entity = await _context.RecipeIngredients.FindAsync(new object[] { id }, ct);
        if (entity != null) _context.RecipeIngredients.Remove(entity);
    }

    public async Task SaveIngredientsAsync(CancellationToken ct = default)
        => await _context.SaveChangesAsync(ct);

    // ========== 配方香调 RecipeNotes ==========

    public async Task<IEnumerable<RecipeNote>> GetRecipeNotesForAccordAsync(int recipeId, CancellationToken ct = default)
        => await _context.RecipeNotes.AsNoTracking()
            .Where(rn => rn.RecipeId == recipeId)
            .OrderBy(rn => rn.Id)
            .ToListAsync(ct);

    public async Task<IEnumerable<RecipeNote>> GetRecipeNotesForProductAsync(int recipeId, CancellationToken ct = default)
        => await _context.RecipeNotes.AsNoTracking()
            .Where(rn => rn.RecipeId == recipeId)
            .OrderBy(rn => rn.Id)
            .ToListAsync(ct);

    public async Task<double> GetNotesPercentSumAsync(int recipeId, CancellationToken ct = default)
    {
        var sum = await _context.RecipeNotes
            .Where(rn => rn.RecipeId == recipeId)
            .SumAsync(rn => rn.Percentage ?? 0, ct);
        return (double)sum;
    }

    public async Task<int> CountNotesByRecipeAsync(int recipeId, CancellationToken ct = default)
        => await _context.RecipeNotes.CountAsync(rn => rn.RecipeId == recipeId, ct);

    public async Task AddRecipeNoteAsync(RecipeNote note, CancellationToken ct = default)
        => await _context.RecipeNotes.AddAsync(note, ct);

    public async Task DeleteRecipeNoteAsync(int id, CancellationToken ct = default)
    {
        var entity = await _context.RecipeNotes.FindAsync(new object[] { id }, ct);
        if (entity != null) _context.RecipeNotes.Remove(entity);
    }

    // ========== 香调 FragranceNotes ==========

    public async Task<IEnumerable<FragranceNote>> GetAllFragranceNotesAsync(CancellationToken ct = default)
        => await _context.FragranceNotes.AsNoTracking()
            .OrderBy(fn => fn.NoteType).ThenBy(fn => fn.NoteName)
            .ToListAsync(ct);

    public async Task<FragranceNote?> GetFragranceNoteByIdAsync(int noteId, CancellationToken ct = default)
        => await _context.FragranceNotes.FindAsync(new object[] { noteId }, ct);

    public async Task<FragranceNote> CreateFragranceNoteAsync(FragranceNote note, CancellationToken ct = default)
    {
        var entry = await _context.FragranceNotes.AddAsync(note, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task UpdateFragranceNoteAsync(FragranceNote note, CancellationToken ct = default)
    {
        _context.FragranceNotes.Update(note);
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeleteFragranceNoteAsync(int noteId, CancellationToken ct = default)
    {
        var entity = await _context.FragranceNotes.FindAsync(new object[] { noteId }, ct);
        if (entity != null) _context.FragranceNotes.Remove(entity);
        await _context.SaveChangesAsync(ct);
    }

    // ========== 基调 BaseNotes ==========

    public async Task<IEnumerable<BaseNote>> GetAllBaseNotesAsync(CancellationToken ct = default)
        => await _context.BaseNotes.AsNoTracking().OrderBy(b => b.BaseNoteName).ToListAsync(ct);

    public async Task<BaseNote?> GetBaseNoteByIdAsync(int baseNoteId, CancellationToken ct = default)
        => await _context.BaseNotes.FindAsync(new object[] { baseNoteId }, ct);

    public async Task<BaseNote> CreateBaseNoteAsync(BaseNote baseNote, CancellationToken ct = default)
    {
        baseNote.CreatedAt = DateTime.Now;
        var entry = await _context.BaseNotes.AddAsync(baseNote, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task UpdateBaseNoteAsync(BaseNote baseNote, CancellationToken ct = default)
    {
        _context.BaseNotes.Update(baseNote);
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeleteBaseNoteAsync(int baseNoteId, CancellationToken ct = default)
    {
        var entity = await _context.BaseNotes.FindAsync(new object[] { baseNoteId }, ct);
        if (entity != null) _context.BaseNotes.Remove(entity);
        await _context.SaveChangesAsync(ct);
    }

    // ========== 原料 RawMaterials ==========

    public async Task<IEnumerable<RawMaterialInventory>> GetAllRawMaterialsAsync(CancellationToken ct = default)
        => await _context.RawMaterialInventories.AsNoTracking()
            .OrderBy(r => r.ItemName)
            .ToListAsync(ct);

    public async Task<int> MatchRawMaterialAsync(string itemName, CancellationToken ct = default)
    {
        var material = await _context.RawMaterialInventories.AsNoTracking()
            .FirstOrDefaultAsync(r => r.ItemName == itemName, ct);
        return material?.MaterialId ?? 0;
    }

    // ========== Accord 发布 ==========

    public async Task<RecipeAccord> CreateAccordAsync(RecipeAccord accord, CancellationToken ct = default)
    {
        accord.Status = "Published";
        accord.PublishedAt = DateTime.Now;
        accord.CreatedAt = DateTime.Now;
        var entry = await _context.RecipeAccords.AddAsync(accord, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task CreateAccordMaterialAsync(RecipeAccordMaterial material, CancellationToken ct = default)
    {
        await _context.RecipeAccordMaterials.AddAsync(material, ct);
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeprecateAccordAsync(int accordRecipeId, int recipeId, int noteId, CancellationToken ct = default)
    {
        var items = await _context.RecipeAccords
            .Where(a => a.AccordRecipeId != accordRecipeId && a.RecipeId == recipeId && a.NoteId == noteId && a.Status == "Published")
            .ToListAsync(ct);
        foreach (var a in items) a.Status = "Deprecated";
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeprecateAccordSingleAsync(int accordRecipeId, CancellationToken ct = default)
    {
        var item = await _context.RecipeAccords.FirstOrDefaultAsync(a => a.AccordRecipeId == accordRecipeId, ct);
        if (item != null) { item.Status = "Deprecated"; await _context.SaveChangesAsync(ct); }
    }

    public async Task<IEnumerable<RecipeAccord>> GetPublishedAccordsAsync(CancellationToken ct = default)
        => await _context.RecipeAccords.AsNoTracking()
            .OrderByDescending(a => a.PublishedAt)
            .ToListAsync(ct);

    public async Task<int> CountPublishedAccordsAsync(CancellationToken ct = default)
        => await _context.RecipeAccords.CountAsync(a => a.Status == "Published", ct);

    public async Task<int> CountAccordsByRecipeAsync(int recipeId, CancellationToken ct = default)
        => await _context.RecipeAccords.CountAsync(a => a.RecipeId == recipeId && a.Status == "Published", ct);

    // ========== 产品配方发布 ==========

    public async Task<RecipeProduct> CreateProductAsync(RecipeProduct product, CancellationToken ct = default)
    {
        product.Status = "Published";
        product.PublishedAt = DateTime.Now;
        product.CreatedAt = DateTime.Now;
        var entry = await _context.RecipeProducts.AddAsync(product, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task CreateProductNoteAsync(RecipeProductNote note, CancellationToken ct = default)
    {
        await _context.RecipeProductNotes.AddAsync(note, ct);
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeprecateProductVersionsAsync(int productRecipeId, int recipeId, CancellationToken ct = default)
    {
        var items = await _context.RecipeProducts
            .Where(p => p.ProductRecipeId != productRecipeId && p.RecipeId == recipeId && p.Status == "Published")
            .ToListAsync(ct);
        foreach (var p in items) p.Status = "Deprecated";
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeprecateProductSingleAsync(int productRecipeId, CancellationToken ct = default)
    {
        var item = await _context.RecipeProducts.FirstOrDefaultAsync(p => p.ProductRecipeId == productRecipeId, ct);
        if (item != null) { item.Status = "Deprecated"; await _context.SaveChangesAsync(ct); }
    }

    public async Task<IEnumerable<RecipeProduct>> GetPublishedProductsAsync(CancellationToken ct = default)
        => await _context.RecipeProducts.AsNoTracking()
            .OrderByDescending(p => p.PublishedAt)
            .ToListAsync(ct);

    public async Task<int> CountPublishedProductsAsync(CancellationToken ct = default)
        => await _context.RecipeProducts.CountAsync(p => p.Status == "Published", ct);

    public async Task<int> CountProductsByRecipeAsync(int recipeId, CancellationToken ct = default)
        => await _context.RecipeProducts.CountAsync(p => p.RecipeId == recipeId && p.Status == "Published", ct);

    // ========== 发布日志 ==========

    public async Task CreatePublishLogAsync(RecipePublishLog log, CancellationToken ct = default)
    {
        log.PublishedAt = DateTime.Now;
        await _context.RecipePublishLogs.AddAsync(log, ct);
        await _context.SaveChangesAsync(ct);
    }

    public async Task<IEnumerable<RecipePublishLog>> GetRecentPublishLogsAsync(int topCount = 20, CancellationToken ct = default)
        => await _context.RecipePublishLogs.AsNoTracking()
            .OrderByDescending(l => l.PublishedAt)
            .Take(topCount)
            .ToListAsync(ct);

    public async Task<int> CountPublishLogsAsync(CancellationToken ct = default)
        => await _context.RecipePublishLogs.CountAsync(ct);

    // ========== 香调库存 NoteInventory ==========

    public async Task<IEnumerable<NoteInventory>> GetAllNoteInventoriesAsync(CancellationToken ct = default)
        => await _context.NoteInventories.AsNoTracking().ToListAsync(ct);

    public async Task UpdateNoteStockAsync(int noteId, int stockQty, int minStockLevel, CancellationToken ct = default)
    {
        var inv = await _context.NoteInventories.FirstOrDefaultAsync(n => n.NoteId == noteId, ct);
        if (inv != null)
        {
            inv.StockQuantity = stockQty;
            inv.MinStockLevel = minStockLevel;
            inv.UpdatedAt = DateTime.Now;
            await _context.SaveChangesAsync(ct);
        }
    }

    public async Task RestockNoteAsync(int noteId, int addQty, CancellationToken ct = default)
    {
        var inv = await _context.NoteInventories.FirstOrDefaultAsync(n => n.NoteId == noteId, ct);
        if (inv != null)
        {
            inv.StockQuantity = (inv.StockQuantity ?? 0) + addQty;
            inv.LastRestockDate = DateTime.Now;
            inv.UpdatedAt = DateTime.Now;
            await _context.SaveChangesAsync(ct);
        }
    }

    public async Task CreateInvTransactionAsync(InventoryTransaction tx, CancellationToken ct = default)
    {
        tx.CreatedAt = DateTime.Now;
        await _context.InventoryTransactions.AddAsync(tx, ct);
        await _context.SaveChangesAsync(ct);
    }

    // ========== 瓶子管理 ==========

    public async Task<IEnumerable<BottleStyle>> GetAllBottleStylesAsync(CancellationToken ct = default)
        => await _context.BottleStyles.AsNoTracking().OrderBy(b => b.BottleName).ToListAsync(ct);

    public async Task<BottleStyle?> GetBottleStyleByIdAsync(int bottleId, CancellationToken ct = default)
        => await _context.BottleStyles.FindAsync(new object[] { bottleId }, ct);

    public async Task<BottleStyle> CreateBottleStyleAsync(BottleStyle style, CancellationToken ct = default)
    {
        var entry = await _context.BottleStyles.AddAsync(style, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task UpdateBottleStyleAsync(BottleStyle style, CancellationToken ct = default)
    {
        style.UpdatedAt = DateTime.Now;
        _context.BottleStyles.Update(style);
        await _context.SaveChangesAsync(ct);
    }

    public async Task DeleteBottleStyleAsync(int bottleId, CancellationToken ct = default)
    {
        var entity = await _context.BottleStyles.FindAsync(new object[] { bottleId }, ct);
        if (entity != null) _context.BottleStyles.Remove(entity);
        await _context.SaveChangesAsync(ct);
    }

    // ========== 产品配置 ==========

    public async Task<IEnumerable<Product>> GetAllProductsAsync(CancellationToken ct = default)
        => await _context.Products.AsNoTracking().OrderBy(p => p.ProductName).ToListAsync(ct);

    public async Task<IEnumerable<ProductTypeConfig>> GetAllProductTypeConfigsAsync(CancellationToken ct = default)
        => await _context.ProductTypeConfigs.AsNoTracking()
            .OrderBy(t => t.DisplayOrder)
            .ToListAsync(ct);

    public async Task<ProductTypeConfig> CreateProductTypeConfigAsync(ProductTypeConfig config, CancellationToken ct = default)
    {
        config.CreatedAt = DateTime.Now;
        var entry = await _context.ProductTypeConfigs.AddAsync(config, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task UpdateProductTypeConfigAsync(ProductTypeConfig config, CancellationToken ct = default)
    {
        _context.ProductTypeConfigs.Update(config);
        await _context.SaveChangesAsync(ct);
    }

    public async Task<IEnumerable<Volume>> GetAllVolumesAsync(CancellationToken ct = default)
        => await _context.Volumes.AsNoTracking().OrderBy(v => v.VolumeMl).ToListAsync(ct);

    public async Task<IEnumerable<ProductVolumePrice>> GetVolumePricesByProductAsync(int productId, CancellationToken ct = default)
        => await _context.ProductVolumePrices.AsNoTracking()
            .Where(p => p.ProductId == productId)
            .ToListAsync(ct);

    public async Task<IEnumerable<ProductBottleStyle>> GetProductBottleStylesAsync(int productId, CancellationToken ct = default)
        => await _context.ProductBottleStyles.AsNoTracking()
            .Where(p => p.ProductId == productId)
            .ToListAsync(ct);

    // ========== 成分聚合 Ingredients ==========

    public async Task<IEnumerable<Ingredient>> GetAllIngredientsAsync(CancellationToken ct = default)
        => await _context.Ingredients.AsNoTracking()
            .OrderBy(i => i.IngredientName)
            .ToListAsync(ct);

    public async Task<Ingredient> CreateIngredientAsync(Ingredient ingredient, CancellationToken ct = default)
    {
        ingredient.CreatedAt = DateTime.Now;
        var entry = await _context.Ingredients.AddAsync(ingredient, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task UpdateIngredientAsync(Ingredient ingredient, CancellationToken ct = default)
    {
        _context.Ingredients.Update(ingredient);
        await _context.SaveChangesAsync(ct);
    }

    // ========== 持久化 ==========

    public async Task<int> SaveChangesAsync(CancellationToken ct = default)
        => await _context.SaveChangesAsync(ct);
}
