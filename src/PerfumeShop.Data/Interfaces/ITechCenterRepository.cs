using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Interfaces;

/// <summary>
/// 技术中心仓储接口 — 对应 ASP dal_techcenter.asp 全部方法
/// </summary>
public interface ITechCenterRepository
{
    // ========== 配方 Recipes ==========
    Task<IEnumerable<Recipe>> GetActiveRecipesAsync(CancellationToken ct = default);
    Task<Recipe?> GetRecipeByIdAsync(int recipeId, CancellationToken ct = default);
    Task<Recipe> CreateRecipeAsync(Recipe recipe, CancellationToken ct = default);
    Task UpdateRecipeAsync(Recipe recipe, CancellationToken ct = default);
    Task<int> CountActiveRecipesAsync(CancellationToken ct = default);

    // ========== 配方成分 RecipeIngredients ==========
    Task<int> CountIngredientsAsync(int recipeId, CancellationToken ct = default);
    Task<int> CountIngredientsByNoteAsync(int recipeId, int noteId, CancellationToken ct = default);
    Task<IEnumerable<RecipeIngredient>> GetAccordIngredientsAsync(int recipeId, int noteId, CancellationToken ct = default);
    Task AddIngredientAsync(RecipeIngredient ingredient, CancellationToken ct = default);
    Task DeleteIngredientAsync(int id, CancellationToken ct = default);
    Task SaveIngredientsAsync(CancellationToken ct = default);

    // ========== 配方香调 RecipeNotes ==========
    Task<IEnumerable<RecipeNote>> GetRecipeNotesForAccordAsync(int recipeId, CancellationToken ct = default);
    Task<IEnumerable<RecipeNote>> GetRecipeNotesForProductAsync(int recipeId, CancellationToken ct = default);
    Task<double> GetNotesPercentSumAsync(int recipeId, CancellationToken ct = default);
    Task<int> CountNotesByRecipeAsync(int recipeId, CancellationToken ct = default);
    Task AddRecipeNoteAsync(RecipeNote note, CancellationToken ct = default);
    Task DeleteRecipeNoteAsync(int id, CancellationToken ct = default);

    // ========== 香调 FragranceNotes ==========
    Task<IEnumerable<FragranceNote>> GetAllFragranceNotesAsync(CancellationToken ct = default);
    Task<FragranceNote?> GetFragranceNoteByIdAsync(int noteId, CancellationToken ct = default);
    Task<FragranceNote> CreateFragranceNoteAsync(FragranceNote note, CancellationToken ct = default);
    Task UpdateFragranceNoteAsync(FragranceNote note, CancellationToken ct = default);
    Task DeleteFragranceNoteAsync(int noteId, CancellationToken ct = default);

    // ========== 基调 BaseNotes ==========
    Task<IEnumerable<BaseNote>> GetAllBaseNotesAsync(CancellationToken ct = default);
    Task<BaseNote?> GetBaseNoteByIdAsync(int baseNoteId, CancellationToken ct = default);
    Task<BaseNote> CreateBaseNoteAsync(BaseNote baseNote, CancellationToken ct = default);
    Task UpdateBaseNoteAsync(BaseNote baseNote, CancellationToken ct = default);
    Task DeleteBaseNoteAsync(int baseNoteId, CancellationToken ct = default);

    // ========== 原料 RawMaterials ==========
    Task<IEnumerable<RawMaterialInventory>> GetAllRawMaterialsAsync(CancellationToken ct = default);
    Task<int> MatchRawMaterialAsync(string itemName, CancellationToken ct = default);

    // ========== Accord 发布 ==========
    Task<RecipeAccord> CreateAccordAsync(RecipeAccord accord, CancellationToken ct = default);
    Task CreateAccordMaterialAsync(RecipeAccordMaterial material, CancellationToken ct = default);
    Task DeprecateAccordAsync(int accordRecipeId, int recipeId, int noteId, CancellationToken ct = default);
    Task DeprecateAccordSingleAsync(int accordRecipeId, CancellationToken ct = default);
    Task<IEnumerable<RecipeAccord>> GetPublishedAccordsAsync(CancellationToken ct = default);
    Task<int> CountPublishedAccordsAsync(CancellationToken ct = default);
    Task<int> CountAccordsByRecipeAsync(int recipeId, CancellationToken ct = default);

    // ========== 产品配方发布 ==========
    Task<RecipeProduct> CreateProductAsync(RecipeProduct product, CancellationToken ct = default);
    Task CreateProductNoteAsync(RecipeProductNote note, CancellationToken ct = default);
    Task DeprecateProductVersionsAsync(int productRecipeId, int recipeId, CancellationToken ct = default);
    Task DeprecateProductSingleAsync(int productRecipeId, CancellationToken ct = default);
    Task<IEnumerable<RecipeProduct>> GetPublishedProductsAsync(CancellationToken ct = default);
    Task<int> CountPublishedProductsAsync(CancellationToken ct = default);
    Task<int> CountProductsByRecipeAsync(int recipeId, CancellationToken ct = default);

    // ========== 发布日志 ==========
    Task CreatePublishLogAsync(RecipePublishLog log, CancellationToken ct = default);
    Task<IEnumerable<RecipePublishLog>> GetRecentPublishLogsAsync(int topCount = 20, CancellationToken ct = default);
    Task<int> CountPublishLogsAsync(CancellationToken ct = default);

    // ========== 香调库存 NoteInventory ==========
    Task<IEnumerable<NoteInventory>> GetAllNoteInventoriesAsync(CancellationToken ct = default);
    Task UpdateNoteStockAsync(int noteId, int stockQty, int minStockLevel, CancellationToken ct = default);
    Task RestockNoteAsync(int noteId, int addQty, CancellationToken ct = default);
    Task CreateInvTransactionAsync(InventoryTransaction tx, CancellationToken ct = default);

    // ========== 瓶子管理 ==========
    Task<IEnumerable<BottleStyle>> GetAllBottleStylesAsync(CancellationToken ct = default);
    Task<BottleStyle?> GetBottleStyleByIdAsync(int bottleId, CancellationToken ct = default);
    Task<BottleStyle> CreateBottleStyleAsync(BottleStyle style, CancellationToken ct = default);
    Task UpdateBottleStyleAsync(BottleStyle style, CancellationToken ct = default);
    Task DeleteBottleStyleAsync(int bottleId, CancellationToken ct = default);

    // ========== 产品配置 ==========
    Task<IEnumerable<Product>> GetAllProductsAsync(CancellationToken ct = default);
    Task<IEnumerable<ProductTypeConfig>> GetAllProductTypeConfigsAsync(CancellationToken ct = default);
    Task<ProductTypeConfig> CreateProductTypeConfigAsync(ProductTypeConfig config, CancellationToken ct = default);
    Task UpdateProductTypeConfigAsync(ProductTypeConfig config, CancellationToken ct = default);
    Task<IEnumerable<Volume>> GetAllVolumesAsync(CancellationToken ct = default);
    Task<IEnumerable<ProductVolumePrice>> GetVolumePricesByProductAsync(int productId, CancellationToken ct = default);
    Task<IEnumerable<ProductBottleStyle>> GetProductBottleStylesAsync(int productId, CancellationToken ct = default);

    // ========== 成分聚合 Ingredients ==========
    Task<IEnumerable<Ingredient>> GetAllIngredientsAsync(CancellationToken ct = default);
    Task<Ingredient> CreateIngredientAsync(Ingredient ingredient, CancellationToken ct = default);
    Task UpdateIngredientAsync(Ingredient ingredient, CancellationToken ct = default);

    // ========== 持久化 ==========
    Task<int> SaveChangesAsync(CancellationToken ct = default);
}
