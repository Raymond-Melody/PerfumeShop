using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class SubscriptionModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public SubscriptionModel(PerfumeShopContext db) => _db = db;

    public List<SubItem> ActiveSubscriptions { get; set; } = new();
    public List<SubItem> HistorySubscriptions { get; set; } = new();
    public List<DeliveryItem> Deliveries { get; set; } = new();
    public string? Message { get; set; }
    public string? Error { get; set; }
    public bool HasActive => ActiveSubscriptions.Any();

    public class SubItem
    {
        public int SubscriptionId { get; set; }
        public int PlanId { get; set; }
        public string Status { get; set; } = "";
        public DateTime StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public bool AutoRenew { get; set; }
        public string PlanName { get; set; } = "";
        public string Period { get; set; } = "";
        public decimal Price { get; set; }
        public int SampleCount { get; set; }
        public int FullSizeCount { get; set; }
        public bool FreeShipping { get; set; }
        public decimal CancellationFee { get; set; }
        public string PeriodLabel => Period?.ToLower() switch
        {
            "monthly" => "月度",
            "quarterly" => "季度",
            "yearly" => "年度",
            _ => Period ?? ""
        };
        public string StatusLabel => Status switch
        {
            "Active" => "活跃中",
            "Paused" => "已暂停",
            "Cancelled" => "已取消",
            "Expired" => "已过期",
            _ => Status
        };
        public string StatusBadge => Status switch
        {
            "Active" => "background:#e8f5e9;color:#2e7d32;",
            "Paused" => "background:#fff3e0;color:#e65100;",
            "Cancelled" => "background:#fce4ec;color:#c62828;",
            "Expired" => "background:#f5f5f5;color:#999;",
            _ => ""
        };
    }

    public class DeliveryItem
    {
        public int DeliveryId { get; set; }
        public int SubscriptionId { get; set; }
        public DateTime DeliveryDate { get; set; }
        public string Status { get; set; } = "";
        public string? TrackingNo { get; set; }
        public string StatusLabel => Status switch
        {
            "Pending" => "待配送",
            "Shipped" => "已发货",
            "Delivered" => "已签收",
            "Skipped" => "已跳过",
            _ => Status
        };
        public string StatusBadge => Status switch
        {
            "Pending" => "background:#e3f2fd;color:#1565c0;",
            "Shipped" => "background:#f3e5f5;color:#6a1b9a;",
            "Delivered" => "background:#e8f5e9;color:#2e7d32;",
            "Skipped" => "background:#fafafa;color:#bbb;",
            _ => ""
        };
    }

    public async Task OnGetAsync()
    {
        var activeSubs = await (from s in _db.UserSubscriptions
                                join p in _db.SubscriptionPlans on s.PlanId equals p.PlanId
                                where s.Status == "Active" || s.Status == "Paused"
                                orderby s.Status ascending, s.StartDate descending
                                select new SubItem
                                {
                                    SubscriptionId = s.SubscriptionId,
                                    PlanId = s.PlanId,
                                    Status = s.Status,
                                    StartDate = s.StartDate,
                                    EndDate = s.EndDate,
                                    AutoRenew = s.AutoRenew,
                                    PlanName = p.PlanName,
                                    Period = p.Period,
                                    Price = p.Price,
                                    SampleCount = p.SampleCount,
                                    FullSizeCount = p.FullSizeCount,
                                    FreeShipping = p.FreeShipping,
                                    CancellationFee = p.CancellationFee,
                                }).ToListAsync();
        ActiveSubscriptions = activeSubs;

        HistorySubscriptions = await (from s in _db.UserSubscriptions
                                      join p in _db.SubscriptionPlans on s.PlanId equals p.PlanId
                                      where s.Status == "Cancelled" || s.Status == "Expired"
                                      orderby s.EndDate descending
                                      select new SubItem
                                      {
                                          SubscriptionId = s.SubscriptionId,
                                          PlanId = s.PlanId,
                                          Status = s.Status,
                                          StartDate = s.StartDate,
                                          EndDate = s.EndDate,
                                          PlanName = p.PlanName,
                                          Period = p.Period,
                                      }).ToListAsync();

        // 获取所有活跃订阅的配送记录
        var activeSubIds = activeSubs.Select(s => s.SubscriptionId).ToList();
        if (activeSubIds.Any())
        {
            Deliveries = await _db.SubscriptionDeliveries
                .Where(d => activeSubIds.Contains(d.SubscriptionId))
                .OrderByDescending(d => d.DeliveryDate)
                .Select(d => new DeliveryItem
                {
                    DeliveryId = d.DeliveryId,
                    SubscriptionId = d.SubscriptionId,
                    DeliveryDate = d.DeliveryDate,
                    Status = d.Status,
                    TrackingNo = d.TrackingNo,
                }).ToListAsync();
        }
    }
}
