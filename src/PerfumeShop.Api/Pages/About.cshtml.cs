using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "StaticContent")]
public class AboutModel : PageModel
{
    public void OnGet() { }
}
