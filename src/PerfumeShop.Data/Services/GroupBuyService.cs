using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

public class GroupBuyService : IGroupBuyService
{
    private readonly PerfumeShopContext _db;
    public GroupBuyService(PerfumeShopContext db) => _db = db;

    public async Task<List<GroupBuyPlanDto>> GetActivePlansAsync()
    {
        var now = DateTime.Now;
        var plans = await (from gp in _db.GroupBuyPlans
                           join p in _db.Products on gp.ProductId equals p.ProductId
                           where gp.IsActive && now >= gp.StartTime && now <= gp.EndTime
                           orderby gp.SortOrder, gp.EndTime
                           select new GroupBuyPlanDto
                           {
                               PlanId = gp.PlanId,
                               ProductId = gp.ProductId,
                               ProductName = p.ProductName ?? "",
                               ImageUrl = p.ImageUrl,
                               GroupPrice = gp.GroupPrice,
                               BasePrice = p.BasePrice,
                               TeamSize = gp.TeamSize,
                               StartTime = gp.StartTime,
                               EndTime = gp.EndTime,
                               DurationHours = gp.DurationHours,
                               Description = p.Description
                           }).ToListAsync();

        // 补充统计
        foreach (var plan in plans)
        {
            plan.OpenGroupCount = await _db.GroupBuyOrders.CountAsync(g => g.PlanId == plan.PlanId && g.Status == 0);
            plan.SuccessGroupCount = await _db.GroupBuyOrders.CountAsync(g => g.PlanId == plan.PlanId && g.Status == 1);
        }
        return plans;
    }

    public async Task<List<OpenGroupDto>> GetOpenGroupsAsync(int planId)
    {
        return await (from g in _db.GroupBuyOrders
                      join u in _db.Users on g.InitiatorId equals u.UserId into uJoin
                      from u in uJoin.DefaultIfEmpty()
                      where g.PlanId == planId && g.Status == 0
                      orderby g.CurrentSize descending, g.CreatedAt
                      select new OpenGroupDto
                      {
                          GroupId = g.GroupId,
                          GroupSn = g.GroupSn,
                          CurrentSize = g.CurrentSize,
                          TargetSize = g.TargetSize,
                          InitiatorName = u != null ? (u.Username ?? "") : "",
                          CreatedAt = g.CreatedAt,
                          HoursPassed = (int)EF.Functions.DateDiffHour(g.CreatedAt, DateTime.Now)
                      }).ToListAsync();
    }

    public async Task<GroupBuyStartResult> StartGroupAsync(int planId, int userId)
    {
        var plan = await _db.GroupBuyPlans.FirstOrDefaultAsync(p => p.PlanId == planId && p.IsActive);
        if (plan == null) return new() { Success = false, Message = "拼团计划不存在或已关闭" };

        var now = DateTime.Now;
        if (now < plan.StartTime || now > plan.EndTime)
            return new() { Success = false, Message = "拼团活动未开始或已结束" };

        var groupSn = $"GB{now:yyyyMMddHHmmss}{new Random().Next(1000, 9999)}";
        var group = new GroupBuyOrder
        {
            PlanId = planId,
            GroupSn = groupSn,
            InitiatorId = userId,
            Status = 0,
            CurrentSize = 1,
            TargetSize = plan.TeamSize,
            CreatedAt = now
        };
        _db.GroupBuyOrders.Add(group);
        await _db.SaveChangesAsync();

        // 添加发起人为第一个参与者
        _db.GroupBuyParticipants.Add(new GroupBuyParticipant
        {
            GroupId = group.GroupId,
            UserId = userId,
            IsInitiator = true,
            Status = 1,
            JoinedAt = now
        });
        await _db.SaveChangesAsync();

        return new() { Success = true, Message = "开团成功", GroupId = group.GroupId, GroupSn = groupSn };
    }

    public async Task<GroupBuyJoinResult> JoinGroupAsync(int groupId, int userId)
    {
        var group = await _db.GroupBuyOrders.FirstOrDefaultAsync(g => g.GroupId == groupId);
        if (group == null) return new() { Success = false, Message = "团不存在" };
        if (group.Status != 0) return new() { Success = false, Message = "该团已结束" };

        // 检查重复参团
        var exists = await _db.GroupBuyParticipants.AnyAsync(p => p.GroupId == groupId && p.UserId == userId);
        if (exists) return new() { Success = false, Message = "您已参加过该团" };

        if (group.CurrentSize >= group.TargetSize)
            return new() { Success = false, Message = "团已满" };

        _db.GroupBuyParticipants.Add(new GroupBuyParticipant
        {
            GroupId = groupId,
            UserId = userId,
            IsInitiator = false,
            Status = 1,
            JoinedAt = DateTime.Now
        });

        group.CurrentSize++;
        var isComplete = group.CurrentSize >= group.TargetSize;
        if (isComplete)
        {
            group.Status = 1; // 已成团
            group.CompletedAt = DateTime.Now;
        }
        await _db.SaveChangesAsync();

        return new() { Success = true, Message = isComplete ? "拼团成功！" : "参团成功", IsGroupComplete = isComplete };
    }

    public async Task<GroupDetailDto?> GetGroupDetailAsync(int groupId)
    {
        var group = await (from g in _db.GroupBuyOrders
                           join gp in _db.GroupBuyPlans on g.PlanId equals gp.PlanId
                           join p in _db.Products on gp.ProductId equals p.ProductId
                           where g.GroupId == groupId
                           select new { g, ProductName = p.ProductName ?? "" }).FirstOrDefaultAsync();
        if (group == null) return null;

        var participants = await (from gp in _db.GroupBuyParticipants
                                  join u in _db.Users on gp.UserId equals u.UserId into uJoin
                                  from u in uJoin.DefaultIfEmpty()
                                  where gp.GroupId == groupId
                                  orderby gp.IsInitiator descending, gp.JoinedAt
                                  select new GroupParticipantDto
                                  {
                                      ParticipantId = gp.ParticipantId,
                                      UserId = gp.UserId,
                                      Username = u != null ? (u.Username ?? "") : "",
                                      IsInitiator = gp.IsInitiator,
                                      Status = gp.Status,
                                      JoinedAt = gp.JoinedAt
                                  }).ToListAsync();

        return new GroupDetailDto
        {
            GroupId = group.g.GroupId,
            GroupSn = group.g.GroupSn,
            PlanId = group.g.PlanId,
            ProductName = group.ProductName,
            CurrentSize = group.g.CurrentSize,
            TargetSize = group.g.TargetSize,
            Status = group.g.Status,
            CreatedAt = group.g.CreatedAt,
            CompletedAt = group.g.CompletedAt,
            Participants = participants
        };
    }

    public async Task<GroupBuyStats> GetPlanStatsAsync(int planId)
    {
        return new GroupBuyStats
        {
            OpenGroups = await _db.GroupBuyOrders.CountAsync(g => g.PlanId == planId && g.Status == 0),
            SuccessGroups = await _db.GroupBuyOrders.CountAsync(g => g.PlanId == planId && g.Status == 1),
            TotalParticipants = await _db.GroupBuyParticipants
                .CountAsync(p => _db.GroupBuyOrders.Any(g => g.GroupId == p.GroupId && g.PlanId == planId))
        };
    }

    public async Task<int> SavePlanAsync(GroupBuyPlan entity)
    {
        if (entity.PlanId > 0)
        {
            var existing = await _db.GroupBuyPlans.FirstOrDefaultAsync(x => x.PlanId == entity.PlanId);
            if (existing == null) return 0;
            existing.ProductId = entity.ProductId;
            existing.TeamSize = entity.TeamSize;
            existing.GroupPrice = entity.GroupPrice;
            existing.MinUnit = entity.MinUnit;
            existing.MaxUnit = entity.MaxUnit;
            existing.StartTime = entity.StartTime;
            existing.EndTime = entity.EndTime;
            existing.DurationHours = entity.DurationHours;
            existing.SortOrder = entity.SortOrder;
            await _db.SaveChangesAsync();
            return existing.PlanId;
        }
        else
        {
            entity.CreatedAt = DateTime.Now;
            entity.IsActive = true;
            _db.GroupBuyPlans.Add(entity);
            await _db.SaveChangesAsync();
            return entity.PlanId;
        }
    }

    public async Task<bool> DeletePlanAsync(int planId)
    {
        var entity = await _db.GroupBuyPlans.FirstOrDefaultAsync(x => x.PlanId == planId);
        if (entity == null) return false;
        _db.GroupBuyPlans.Remove(entity);
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<bool> TogglePlanActiveAsync(int planId)
    {
        return await _db.GroupBuyPlans
            .Where(x => x.PlanId == planId)
            .ExecuteUpdateAsync(s => s.SetProperty(e => e.IsActive, e => !e.IsActive)) > 0;
    }
}
