using Microsoft.AspNetCore.Mvc.RazorPages;

namespace PerfumeShop.Api.Pages.User;

public class AccountDeleteModel : PageModel
{
    public string? Message { get; set; }
    public string? Error { get; set; }
    public bool ShowForm { get; set; } = true;

    public void OnGet()
    {
    }
}
