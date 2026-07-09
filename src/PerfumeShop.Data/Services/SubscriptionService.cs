using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

public class SubscriptionService : ISubscriptionService
{
    private readonly PerfumeShopContext _db;
    public SubscriptionService(PerfumeShopContext db) => _db = db;

    public async Task<List<SubscriptionPlan>> GetActivePlansAsync()
    {
        return await _db.SubscriptionPlans
            .Where(p => p.IsActive)
            .OrderBy(p => p.SortOrder)
            .ToListAsync();
    }

    public async Task<List<SubscriptionPlan>> GetAllPlansAsync()
    {
        return await _db.SubscriptionPlans
            .OrderBy(p => p.SortOrder)
            .ToListAsync();
    }

    public async Task<UserSubscriptionDto?> GetUserSubscriptionAsync(int userId)
    {
        return await (from us in _db.UserSubscriptions
                      join sp in _db.SubscriptionPlans on us.PlanId equals sp.PlanId
                      where us.UserId == userId
                      orderby us.CreatedAt descending
                      select new UserSubscriptionDto
                      {
                          SubscriptionId = us.SubscriptionId,
                          PlanId = us.PlanId,
                          PlanName = sp.PlanName,
                          Period = sp.Period,
                          Price = sp.Price,
                          Status = us.Status,
                          StartDate = us.StartDate,
                          EndDate = us.EndDate,
                          AutoRenew = us.AutoRenew,
                          CreatedAt = us.CreatedAt,
                          DeliveryCount = _db.SubscriptionDeliveries.Count(d => d.SubscriptionId == us.SubscriptionId)
                      }).FirstOrDefaultAsync();
    }

    public async Task<SubscribeResult> SubscribeAsync(int userId, int planId, bool autoRenew = true)
    {
        var plan = await _db.SubscriptionPlans.FirstOrDefaultAsync(p => p.PlanId == planId && p.IsActive);
        if (plan == null) return new() { Success = false, Message = "订阅计划不存在或已关闭" };

        // 检查是否已有活跃订阅
        var existing = await _db.UserSubscriptions
            .AnyAsync(s => s.UserId == userId && s.Status != "cancelled" && s.Status != "expired");
        if (existing) return new() { Success = false, Message = "您已有活跃订阅，请先取消或等其到期" };

        var sub = new UserSubscription
        {
            UserId = userId,
            PlanId = planId,
            Status = "active",
            StartDate = DateTime.Now.Date,
            AutoRenew = autoRenew,
            CreatedAt = DateTime.Now
        };
        _db.UserSubscriptions.Add(sub);
        await _db.SaveChangesAsync();

        return new() { Success = true, Message = "订阅成功", SubscriptionId = sub.SubscriptionId };
    }

    public async Task<bool> CancelSubscriptionAsync(int subscriptionId, int userId)
    {
        var sub = await _db.UserSubscriptions
            .FirstOrDefaultAsync(s => s.SubscriptionId == subscriptionId && s.UserId == userId);
        if (sub == null) return false;

        sub.Status = "cancelled";
        sub.EndDate = DateTime.Now;
        sub.AutoRenew = false;
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<bool> ToggleAutoRenewAsync(int subscriptionId, int userId, bool autoRenew)
    {
        var sub = await _db.UserSubscriptions
            .FirstOrDefaultAsync(s => s.SubscriptionId == subscriptionId && s.UserId == userId);
        if (sub == null) return false;

        sub.AutoRenew = autoRenew;
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<List<SubscriptionDelivery>> GetDeliveryHistoryAsync(int subscriptionId)
    {
        return await _db.SubscriptionDeliveries
            .Where(d => d.SubscriptionId == subscriptionId)
            .OrderByDescending(d => d.DeliveryDate)
            .ToListAsync();
    }

    public async Task<int> SavePlanAsync(SubscriptionPlan entity)
    {
        if (entity.PlanId > 0)
        {
            var existing = await _db.SubscriptionPlans.FirstOrDefaultAsync(x => x.PlanId == entity.PlanId);
            if (existing == null) return 0;
            existing.PlanName = entity.PlanName;
            existing.Period = entity.Period;
            existing.Price = entity.Price;
            existing.SampleCount = entity.SampleCount;
            existing.FullSizeCount = entity.FullSizeCount;
            existing.FreeShipping = entity.FreeShipping;
            existing.CancellationFee = entity.CancellationFee;
            existing.Description = entity.Description;
            existing.SortOrder = entity.SortOrder;
            await _db.SaveChangesAsync();
            return existing.PlanId;
        }
        else
        {
            entity.CreatedAt = DateTime.Now;
            entity.IsActive = true;
            _db.SubscriptionPlans.Add(entity);
            await _db.SaveChangesAsync();
            return entity.PlanId;
        }
    }

    public async Task<bool> DeletePlanAsync(int planId)
    {
        var entity = await _db.SubscriptionPlans.FirstOrDefaultAsync(x => x.PlanId == planId);
        if (entity == null) return false;
        _db.SubscriptionPlans.Remove(entity);
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<bool> TogglePlanActiveAsync(int planId)
    {
        return await _db.SubscriptionPlans
            .Where(x => x.PlanId == planId)
            .ExecuteUpdateAsync(s => s.SetProperty(e => e.IsActive, e => !e.IsActive)) > 0;
    }
}
