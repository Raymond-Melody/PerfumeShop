using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class SiteSetting
{
    public string? Description { get; set; }

    public string? SettingKey { get; set; }

    public string? SettingName { get; set; }

    public string? SettingValue { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public int? SecurityPasswordMinLength { get; set; }

    public int? SecuritySessionTimeout { get; set; }

    public int? SecurityLoginMaxAttempts { get; set; }

    public bool? SecurityMfaenabled { get; set; }

    public int? SecurityLockoutMinutes { get; set; }
}
