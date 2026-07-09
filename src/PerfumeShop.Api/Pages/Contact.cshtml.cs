using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "StaticContent")]
public class ContactModel : PageModel
{
    public void OnGet() { }
}
