using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class Volume
{
    public bool? IsActive { get; set; }

    public double? PriceMultiplier { get; set; }

    public int VolumeId { get; set; }

    public int VolumeMl { get; set; }

    public string VolumeName { get; set; } = null!;
}
