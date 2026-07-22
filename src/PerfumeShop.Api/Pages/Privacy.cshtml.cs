using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "StaticContent")]
public class PrivacyModel : PageModel
{
    public string Lang { get; set; } = "zh-CN";
    public string CurrentVersion { get; set; } = "2025-01-01";

    public List<VersionItem> Versions { get; set; } = new()
    {
        new VersionItem { Version = "2025-01-01", Summary = "更新数据保留政策和用户权利说明" },
        new VersionItem { Version = "2024-06-15", Summary = "增加 Cookie 使用和第三方服务说明" },
        new VersionItem { Version = "2024-01-01", Summary = "初始版本发布" },
    };

    public class VersionItem
    {
        public string Version { get; set; } = "";
        public string Summary { get; set; } = "";
    }

    public void OnGet(string? lang = null)
    {
        Lang = lang ?? Request.Cookies["PERFUME_LANG"] ?? "zh-CN";
    }
}
