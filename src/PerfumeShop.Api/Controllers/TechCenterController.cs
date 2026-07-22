using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 技术中心 API — 配方导出、香调导入等操作
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class TechCenterController : ControllerBase
{
    private readonly ITechCenterRepository _repo;

    public TechCenterController(ITechCenterRepository repo)
    {
        _repo = repo ?? throw new ArgumentNullException(nameof(repo));
    }

    /// <summary>获取活跃配方列表</summary>
    [HttpGet("recipes")]
    public async Task<IActionResult> GetRecipes()
    {
        var recipes = await _repo.GetActiveRecipesAsync();
        return Ok(recipes.Select(r => new { r.RecipeId, r.RecipeName, r.RecipeCode, r.ProductType }));
    }

    /// <summary>获取配方详情</summary>
    [HttpGet("recipes/{id:int}")]
    public async Task<IActionResult> GetRecipe(int id)
    {
        var recipe = await _repo.GetRecipeByIdAsync(id);
        if (recipe == null) return NotFound();
        return Ok(recipe);
    }

    /// <summary>导出配方（JSON 格式）</summary>
    [HttpGet("recipes/{id:int}/export")]
    public async Task<IActionResult> ExportRecipe(int id)
    {
        var recipe = await _repo.GetRecipeByIdAsync(id);
        if (recipe == null) return NotFound();

        var ingredients = await _repo.GetAccordIngredientsAsync(id, 0);
        return Ok(new { Recipe = recipe, Ingredients = ingredients });
    }

    /// <summary>获取所有香调</summary>
    [HttpGet("notes")]
    public async Task<IActionResult> GetNotes()
    {
        var notes = await _repo.GetAllFragranceNotesAsync();
        return Ok(notes);
    }

    /// <summary>创建香调</summary>
    [HttpPost("notes")]
    public async Task<IActionResult> CreateNote([FromBody] FragranceNote note)
    {
        var created = await _repo.CreateFragranceNoteAsync(note);
        return CreatedAtAction(nameof(GetNotes), new { id = created.NoteId }, created);
    }

    /// <summary>获取所有原材料</summary>
    [HttpGet("materials")]
    public async Task<IActionResult> GetMaterials()
    {
        var materials = await _repo.GetAllRawMaterialsAsync();
        return Ok(materials);
    }

    /// <summary>获取已发布 Accord 列表</summary>
    [HttpGet("accords")]
    public async Task<IActionResult> GetAccords()
    {
        var accords = await _repo.GetPublishedAccordsAsync();
        return Ok(accords);
    }

    /// <summary>获取发布日志</summary>
    [HttpGet("logs")]
    public async Task<IActionResult> GetLogs([FromQuery] int top = 50)
    {
        var logs = await _repo.GetRecentPublishLogsAsync(top);
        return Ok(logs);
    }

    /// <summary>技术中心统计</summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetStats()
    {
        return Ok(new
        {
            ActiveRecipes = await _repo.CountActiveRecipesAsync(),
            PublishedAccords = await _repo.CountPublishedAccordsAsync(),
            PublishedProducts = await _repo.CountPublishedProductsAsync(),
            PublishLogs = await _repo.CountPublishLogsAsync()
        });
    }
}
