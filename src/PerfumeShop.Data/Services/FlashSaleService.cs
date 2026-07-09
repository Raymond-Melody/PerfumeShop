using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

public class FlashSaleService : IFlashSaleService
{
    private readonly PerfumeShopContext _db;
    public FlashSaleService(PerfumeShopContext db) => _db = db;

    public async Task<(List<FlashSaleDto> Items, int Total)> GetActiveFlashSalesAsync(int page = 1, int pageSize = 12)
    {
        var now = DateTime.Now;
        var query = from fs in _db.FlashSales
                    join p in _db.Products on fs.ProductId equals p.ProductId
                    where fs.IsActive && now >= fs.StartTime && now <= fs.EndTime && fs.Stock > fs.SoldCount
                    orderby fs.SortOrder, fs.EndTime
                    select new FlashSaleDto
                    {
                        FlashSaleId = fs.FlashSaleId,
                        ProductId = fs.ProductId,
                        ProductName = p.ProductName ?? "",
                        ImageUrl = p.ImageUrl,
                        FlashPrice = fs.FlashPrice,
                        BasePrice = p.BasePrice,
                        Stock = fs.Stock,
                        SoldCount = fs.SoldCount,
                        LimitPerUser = fs.LimitPerUser,
                        StartTime = fs.StartTime,
                        EndTime = fs.EndTime,
                        Category = p.Category,
                        Description = p.Description
                    };

        var total = await query.CountAsync();
        var items = await query.Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }

    public async Task<List<FlashSaleDto>> GetUpcomingFlashSalesAsync(int top = 6)
    {
        var now = DateTime.Now;
        return await (from fs in _db.FlashSales
                      join p in _db.Products on fs.ProductId equals p.ProductId
                      where fs.IsActive && now < fs.StartTime && fs.Stock > 0
                      orderby fs.StartTime
                      select new FlashSaleDto
                      {
                          FlashSaleId = fs.FlashSaleId,
                          ProductId = fs.ProductId,
                          ProductName = p.ProductName ?? "",
                          ImageUrl = p.ImageUrl,
                          FlashPrice = fs.FlashPrice,
                          BasePrice = p.BasePrice,
                          Stock = fs.Stock,
                          SoldCount = fs.SoldCount,
                          LimitPerUser = fs.LimitPerUser,
                          StartTime = fs.StartTime,
                          EndTime = fs.EndTime
                      }).Take(top).ToListAsync();
    }

    public async Task<FlashSaleDto?> GetFlashSaleByIdAsync(int flashSaleId)
    {
        return await (from fs in _db.FlashSales
                      join p in _db.Products on fs.ProductId equals p.ProductId
                      where fs.FlashSaleId == flashSaleId
                      select new FlashSaleDto
                      {
                          FlashSaleId = fs.FlashSaleId,
                          ProductId = fs.ProductId,
                          ProductName = p.ProductName ?? "",
                          ImageUrl = p.ImageUrl,
                          FlashPrice = fs.FlashPrice,
                          BasePrice = p.BasePrice,
                          Stock = fs.Stock,
                          SoldCount = fs.SoldCount,
                          LimitPerUser = fs.LimitPerUser,
                          StartTime = fs.StartTime,
                          EndTime = fs.EndTime,
                          Category = p.Category,
                          Description = p.Description
                      }).FirstOrDefaultAsync();
    }

    public async Task<FlashSalePurchaseResult> PurchaseAsync(int flashSaleId, int userId, int quantity)
    {
        var fs = await _db.FlashSales.FirstOrDefaultAsync(x => x.FlashSaleId == flashSaleId);
        if (fs == null) return new() { Success = false, Message = "秒杀活动不存在" };

        var now = DateTime.Now;
        if (!fs.IsActive || now < fs.StartTime || now > fs.EndTime)
            return new() { Success = false, Message = "秒杀活动未开始或已结束" };
        if (fs.Stock - fs.SoldCount < quantity)
            return new() { Success = false, Message = "库存不足" };
        if (quantity > fs.LimitPerUser)
            return new() { Success = false, Message = $"每人限购{fs.LimitPerUser}件" };

        // 乐观库存扣减
        fs.SoldCount += quantity;
        await _db.SaveChangesAsync();

        return new() { Success = true, Message = "抢购成功，请尽快完成支付" };
    }

    public async Task<int> SaveFlashSaleAsync(FlashSale entity)
    {
        if (entity.FlashSaleId > 0)
        {
            var existing = await _db.FlashSales.FirstOrDefaultAsync(x => x.FlashSaleId == entity.FlashSaleId);
            if (existing == null) return 0;
            existing.ProductId = entity.ProductId;
            existing.FlashPrice = entity.FlashPrice;
            existing.Stock = entity.Stock;
            existing.LimitPerUser = entity.LimitPerUser;
            existing.StartTime = entity.StartTime;
            existing.EndTime = entity.EndTime;
            existing.SortOrder = entity.SortOrder;
            await _db.SaveChangesAsync();
            return existing.FlashSaleId;
        }
        else
        {
            entity.CreatedAt = DateTime.Now;
            entity.SoldCount = 0;
            entity.IsActive = true;
            _db.FlashSales.Add(entity);
            await _db.SaveChangesAsync();
            return entity.FlashSaleId;
        }
    }

    public async Task<bool> DeleteFlashSaleAsync(int flashSaleId)
    {
        var entity = await _db.FlashSales.FirstOrDefaultAsync(x => x.FlashSaleId == flashSaleId);
        if (entity == null) return false;
        _db.FlashSales.Remove(entity);
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<bool> ToggleActiveAsync(int flashSaleId)
    {
        return await _db.FlashSales
            .Where(x => x.FlashSaleId == flashSaleId)
            .ExecuteUpdateAsync(s => s.SetProperty(e => e.IsActive, e => !e.IsActive)) > 0;
    }

    public async Task<FlashSaleAdminStats> GetAdminStatsAsync()
    {
        var now = DateTime.Now;
        return new FlashSaleAdminStats
        {
            Total = await _db.FlashSales.CountAsync(),
            Active = await _db.FlashSales.CountAsync(x => x.IsActive && now >= x.StartTime && now <= x.EndTime),
            Upcoming = await _db.FlashSales.CountAsync(x => x.IsActive && now < x.StartTime),
            Expired = await _db.FlashSales.CountAsync(x => now > x.EndTime)
        };
    }
}
