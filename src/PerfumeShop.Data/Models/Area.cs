namespace PerfumeShop.Data.Models;

/// <summary>
/// 地区数据 — 对应 V18 api/get_areas.asp
/// </summary>
public class Area
{
    public int AreaId { get; set; }
    public string AreaName { get; set; } = null!;
    public int ParentId { get; set; }
}
