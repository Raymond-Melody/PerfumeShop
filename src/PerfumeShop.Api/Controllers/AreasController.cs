using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 地区数据 API — 对应 V18 api/get_areas.asp
/// </summary>
[ApiController]
[Route("api/v2/areas")]
public class AreasController : ControllerBase
{
    private readonly PerfumeShopContext _db;

    public AreasController(PerfumeShopContext db)
    {
        _db = db;
    }

    /// <summary>GET /api/v2/areas — 获取地区列表</summary>
    [HttpGet]
    public async Task<IActionResult> GetAreas(
        [FromQuery] int parentId = 0,
        [FromQuery] string? parentName = null)
    {
        try
        {
            int effectiveParentId = parentId;

            // If parent_name provided, look up its ID first
            if (!string.IsNullOrEmpty(parentName) && _db.Set<Area>().Any())
            {
                var parent = await _db.Set<Area>()
                    .FirstOrDefaultAsync(a => a.AreaName == parentName);
                if (parent != null)
                    effectiveParentId = parent.AreaId;
            }

            var areas = await _db.Set<Area>()
                .Where(a => a.ParentId == effectiveParentId)
                .OrderBy(a => a.AreaId)
                .Select(a => new { areaId = a.AreaId, areaName = a.AreaName })
                .ToListAsync();

            return Ok(areas);
        }
        catch (Exception ex)
        {
            // Table might not exist — return empty array
            return Ok(new { error = ex.Message, areas = Array.Empty<object>() });
        }
    }
}
