using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class RegistrationAttempt
{
    public int AttemptId { get; set; }

    public string Ipaddress { get; set; } = null!;

    public string? DeviceFingerprint { get; set; }

    public bool? Success { get; set; }

    public string? TokenHash { get; set; }

    public DateTime? AttemptedAt { get; set; }
}
