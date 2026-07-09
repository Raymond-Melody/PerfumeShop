using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

public class CustomizeModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public CustomizeModel(PerfumeShopContext db) => _db = db;

    public List<FragranceNote> TopNotes { get; set; } = new();
    public List<FragranceNote> MiddleNotes { get; set; } = new();
    public List<FragranceNote> BaseNotes { get; set; } = new();

    public async Task OnGetAsync()
    {
        var notes = await _db.FragranceNotes
            .AsNoTracking()
            .Where(n => n.IsActive == true)
            .OrderBy(n => n.NoteName)
            .ToListAsync();

        TopNotes = notes.Where(n => n.NoteType == "Top").ToList();
        MiddleNotes = notes.Where(n => n.NoteType == "Middle").ToList();
        BaseNotes = notes.Where(n => n.NoteType == "Base").ToList();
    }
}
