using System;
using System.Collections.Generic;

namespace PerfumeShop.Data.Models;

public partial class GroupBuyParticipant
{
    public int ParticipantId { get; set; }

    public int GroupId { get; set; }

    public int UserId { get; set; }

    public int? OrderId { get; set; }

    public bool IsInitiator { get; set; }

    public int Status { get; set; }

    public DateTime JoinedAt { get; set; }
}
