using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "StaticContent")]
public class FragranceQuizModel : PageModel
{
    public void OnGet()
    {
        // 静态页面，香氛测试为前端交互
    }
}
