using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 半成品仓储实现 — V19 M4-C
/// </summary>
public class SemiFinishedRepository : ISemiFinishedRepository
{
    private readonly PerfumeShopContext _context;

    public SemiFinishedRepository(PerfumeShopContext context) => _context = context;

    // ===== AccordProduction =====

    public async Task<(IEnumerable<AccordProduction> Items, int TotalCount)> GetAccordProductionsAsync(
        string? status, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.AccordProductions.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) query = query.Where(a => a.Status == status);
        int total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(a => a.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task<AccordProduction?> GetAccordProductionDetailAsync(int productionId, CancellationToken ct = default)
        => await _context.AccordProductions.AsNoTracking()
            .FirstOrDefaultAsync(a => a.ProductionId == productionId, ct);

    public async Task<IEnumerable<AccordProductionDetail>> GetAccordProductionDetailsAsync(int productionId, CancellationToken ct = default)
        => await _context.AccordProductionDetails.AsNoTracking()
            .Where(d => d.ProductionId == productionId).ToListAsync(ct);

    public async Task<AccordProduction> CreateBlendRecordAsync(AccordProduction record, CancellationToken ct = default)
    {
        record.CreatedAt = DateTime.Now;
        record.UpdatedAt = DateTime.Now;
        var entry = await _context.AccordProductions.AddAsync(record, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task<bool> UpdateBlendStatusAsync(int productionId, string newStatus, CancellationToken ct = default)
    {
        var rec = await _context.AccordProductions.FirstOrDefaultAsync(a => a.ProductionId == productionId, ct);
        if (rec == null) return false;
        rec.Status = newStatus;
        rec.UpdatedAt = DateTime.Now;
        if (newStatus == "Completed") rec.CompletedAt = DateTime.Now;
        if (newStatus == "InProgress") rec.StartedAt = DateTime.Now;
        await _context.SaveChangesAsync(ct);
        return true;
    }

    // ===== NoteInventory / RawMaterialInventory =====

    public async Task<IEnumerable<NoteInventory>> GetNoteInventoryAsync(CancellationToken ct = default)
        => await _context.NoteInventories.AsNoTracking().OrderBy(n => n.NoteId).ToListAsync(ct);

    public async Task<IEnumerable<RawMaterialInventory>> GetRawMaterialInventoryAsync(CancellationToken ct = default)
        => await _context.RawMaterialInventories.AsNoTracking().OrderBy(r => r.MaterialId).ToListAsync(ct);

    public async Task<IEnumerable<object>> GetInventoryAlertsAsync(CancellationToken ct = default)
    {
        var noteAlerts = await _context.NoteInventories.AsNoTracking()
            .Where(n => n.StockQuantity < n.MinStockLevel)
            .Select(n => new { Type = "香料", Id = (int)n.NoteId, Name = $"Note #{n.NoteId}", Stock = (double)(n.StockQuantity ?? 0), MinStock = (double)(n.MinStockLevel ?? 0) })
            .ToListAsync(ct);

        var rawAlerts = await _context.RawMaterialInventories.AsNoTracking()
            .Where(r => r.StockQty < r.SafetyStock)
            .Select(r => new { Type = "原料", Id = r.MaterialId, Name = r.ItemName ?? "", Stock = r.StockQty ?? 0, MinStock = r.SafetyStock ?? 0 })
            .ToListAsync(ct);

        return noteAlerts.Cast<object>().Concat(rawAlerts.Cast<object>()).ToList();
    }

    // ===== WorkshopTransfer =====

    public async Task<(IEnumerable<WorkshopTransfer> Items, int TotalCount)> GetWorkshopTransfersAsync(
        string? status, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _context.WorkshopTransfers.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) query = query.Where(t => t.Status == status);
        int total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(t => t.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task<WorkshopTransfer> TransferSemiFinishedAsync(WorkshopTransfer transfer, CancellationToken ct = default)
    {
        transfer.CreatedAt = DateTime.Now;
        transfer.RequestedAt = DateTime.Now;
        transfer.Status = "Pending";
        transfer.TransferNo = $"TF-{DateTime.Now:yyyyMMddHHmmss}-{new Random().Next(100, 999)}";
        var entry = await _context.WorkshopTransfers.AddAsync(transfer, ct);
        await _context.SaveChangesAsync(ct);
        return entry.Entity;
    }

    public async Task<bool> FulfillTransferAsync(int transferId, CancellationToken ct = default)
    {
        var t = await _context.WorkshopTransfers.FirstOrDefaultAsync(x => x.TransferId == transferId, ct);
        if (t == null) return false;
        t.Status = "Completed";
        t.FulfilledAt = DateTime.Now;
        await _context.SaveChangesAsync(ct);
        return true;
    }

    public async Task<IEnumerable<BaseNote>> GetBaseNotesAsync(CancellationToken ct = default)
        => await _context.BaseNotes.AsNoTracking().OrderBy(b => b.BaseNoteId).ToListAsync(ct);

    public async Task<IEnumerable<AccordProduction>> GetReportDataAsync(DateTime start, DateTime end, CancellationToken ct = default)
        => await _context.AccordProductions.AsNoTracking()
            .Where(a => a.CreatedAt >= start && a.CreatedAt <= end)
            .OrderBy(a => a.CreatedAt).ToListAsync(ct);
}
