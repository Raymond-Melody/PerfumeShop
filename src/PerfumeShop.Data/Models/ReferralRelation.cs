using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class ReferralRelation
{
    public int RelationId { get; set; }

    public int AncestorUserId { get; set; }

    public int DescendantUserId { get; set; }

    public int Depth { get; set; }

    public DateTime? CreatedAt { get; set; }
}
