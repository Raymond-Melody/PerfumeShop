using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class InventoryRepository
{
    private readonly PerfumeShopContext _context;
    public InventoryRepository(PerfumeShopContext context) => _context = context;

    public async Task<List<ProductInventory>> GetProductInventoriesAsync() =>
        await _context.ProductInventories.AsNoTracking().ToListAsync();
    public async Task<List<RawMaterialInventory>> GetRawMaterialInventoriesAsync() =>
        await _context.RawMaterialInventories.AsNoTracking().ToListAsync();
    public async Task<List<BottleInventory>> GetBottleInventoriesAsync() =>
        await _context.BottleInventories.AsNoTracking().ToListAsync();
    public async Task<List<PackagingInventory>> GetPackagingInventoriesAsync() =>
        await _context.PackagingInventories.AsNoTracking().ToListAsync();
    public async Task<(List<StockMovement> Items, int Total)> GetStockMovementsAsync(int page, int pageSize, string? itemType = null)
    {
        var q = _context.StockMovements.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(itemType)) q = q.Where(m => m.ItemType == itemType);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(m => m.MovementId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<List<InventoryBatch>> GetInventoryBatchesAsync() =>
        await _context.InventoryBatches.AsNoTracking().ToListAsync();
    public async Task<(List<InventoryTransaction> Items, int Total)> GetInventoryTransactionsAsync(int page, int pageSize)
    {
        var total = await _context.InventoryTransactions.CountAsync();
        var items = await _context.InventoryTransactions.AsNoTracking()
            .OrderByDescending(t => t.TransactionId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<List<ProductInventory>> GetLowStockProductsAsync(int threshold = 10) =>
        await _context.ProductInventories.AsNoTracking().Where(p => p.StockQty <= threshold).ToListAsync();
    public async Task<List<RawMaterialInventory>> GetLowStockMaterialsAsync(decimal threshold = 50) =>
        await _context.RawMaterialInventories.AsNoTracking().Where(m => m.StockQty <= (double)threshold).ToListAsync();
}
