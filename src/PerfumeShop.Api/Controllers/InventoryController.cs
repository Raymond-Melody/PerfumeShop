using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class InventoryController : ControllerBase
{
    private readonly InventoryRepository _repo;
    public InventoryController(InventoryRepository repo) => _repo = repo;

    [HttpGet("products")]
    public async Task<IActionResult> GetProductInventories() => Ok(await _repo.GetProductInventoriesAsync());

    [HttpGet("materials")]
    public async Task<IActionResult> GetRawMaterials() => Ok(await _repo.GetRawMaterialInventoriesAsync());

    [HttpGet("bottles")]
    public async Task<IActionResult> GetBottles() => Ok(await _repo.GetBottleInventoriesAsync());

    [HttpGet("packaging")]
    public async Task<IActionResult> GetPackaging() => Ok(await _repo.GetPackagingInventoriesAsync());

    [HttpGet("movements")]
    public async Task<IActionResult> GetMovements(int page = 1, int pageSize = 20, string? itemType = null)
    {
        var (items, total) = await _repo.GetStockMovementsAsync(page, pageSize, itemType);
        return Ok(new { items, total });
    }

    [HttpGet("alerts/products")]
    public async Task<IActionResult> GetLowStockProducts(int threshold = 10) => Ok(await _repo.GetLowStockProductsAsync(threshold));

    [HttpGet("alerts/materials")]
    public async Task<IActionResult> GetLowStockMaterials(decimal threshold = 50) => Ok(await _repo.GetLowStockMaterialsAsync(threshold));
}
