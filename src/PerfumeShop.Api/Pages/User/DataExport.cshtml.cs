using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class DataExportModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public DataExportModel(PerfumeShopContext db) => _db = db;

    public string? ExportType { get; set; }
    public DateTime? DateFrom { get; set; }
    public DateTime? DateTo { get; set; }
    public string? Message { get; set; }

    public async Task OnGetAsync()
    {
        await Task.CompletedTask;
    }
}
