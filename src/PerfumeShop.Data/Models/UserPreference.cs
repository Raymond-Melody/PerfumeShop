using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class UserPreference
{
    public DateTime? CreatedAt { get; set; }

    public int PreferenceId { get; set; }

    public string? PreferredBaseNotes { get; set; }

    public string? PreferredCategories { get; set; }

    public string? PreferredMiddleNotes { get; set; }

    public string? PreferredTopNotes { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public int UserId { get; set; }
}
