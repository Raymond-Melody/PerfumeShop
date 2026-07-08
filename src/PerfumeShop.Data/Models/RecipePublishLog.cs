using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RecipePublishLog
{
    public string? Ipaddress { get; set; }

    public int LogId { get; set; }

    public DateTime? PublishedAt { get; set; }

    public string? PublishedBy { get; set; }

    public string? PublishType { get; set; }

    public int? RecipeId { get; set; }

    public int? TargetRecipeId { get; set; }
}
