using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 半成品仓储实现 — V19 M4-C
/// </summary>
public class SemiFinishedRepository : ISemiFinishedRepository
{
    private readonly PerfumeShopContext _context;
    private readonly IInventoryLedger _ledger;

    public SemiFinishedRepository(PerfumeShopContext context, IInventoryLedger ledger)
    {
        _context = context;
        _ledger = ledger;
    }

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
        // V21: 全品类安全库存预警——原料/香调/成品/瓶身/包材
        var noteAlerts = await _context.NoteInventories.AsNoTracking()
            .Where(n => n.StockQuantity < n.MinStockLevel)
            .Select(n => new { Type = "香料", Id = (int)n.NoteId, Name = $"Note #{n.NoteId}", Stock = (double)(n.StockQuantity ?? 0), MinStock = (double)(n.MinStockLevel ?? 0) })
            .ToListAsync(ct);

        var rawAlerts = await _context.RawMaterialInventories.AsNoTracking()
            .Where(r => r.StockQty < r.SafetyStock)
            .Select(r => new { Type = "原料", Id = r.MaterialId, Name = r.ItemName ?? "", Stock = r.StockQty ?? 0, MinStock = r.SafetyStock ?? 0 })
            .ToListAsync(ct);

        var productAlerts = await _context.ProductInventories.AsNoTracking()
            .Where(p => p.ProductId != null && p.StockQty < p.SafetyStock)
            .Select(p => new { Type = "成品", Id = p.ProductId ?? 0, Name = $"Product #{p.ProductId}", Stock = (double)(p.StockQty ?? 0), MinStock = (double)(p.SafetyStock ?? 0) })
            .ToListAsync(ct);

        var bottleAlerts = await _context.BottleStyles.AsNoTracking()
            .Where(b => b.SafetyStock != null && b.StockQty < b.SafetyStock)
            .Select(b => new { Type = "瓶身", Id = b.BottleId, Name = b.BottleName, Stock = (double)(b.StockQty ?? 0), MinStock = (double)(b.SafetyStock ?? 0) })
            .ToListAsync(ct);

        var packagingAlerts = await _context.PackagingInventories.AsNoTracking()
            .Where(p => p.SafetyStock != null && p.StockQty < p.SafetyStock)
            .Select(p => new { Type = "包材", Id = p.PackagingId, Name = p.ItemName ?? "", Stock = (double)(p.StockQty ?? 0), MinStock = (double)(p.SafetyStock ?? 0) })
            .ToListAsync(ct);

        return noteAlerts.Cast<object>()
            .Concat(rawAlerts.Cast<object>())
            .Concat(productAlerts.Cast<object>())
            .Concat(bottleAlerts.Cast<object>())
            .Concat(packagingAlerts.Cast<object>())
            .ToList();
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

    // ===== V21: 香调生产（一步完成：校验→扣原料→产香调→成本结转→双向流水）=====
    // 对标 V18 admin/semifinished/accord_production.asp start_production

    public record AccordProductionResult(bool Success, string Message, int ProductionId, string BatchNo);

    /// <summary>
    /// 启动并完成香调生产：按 RecipeAccordMaterials 配比扣原料(OUT流水)，产出香调入 NoteInventory，
    /// 结转香调单位成本 = 本批原料总成本 / 产出量 → NoteInventory.WeightedUnitCost(IN流水)。
    /// </summary>
    public async Task<AccordProductionResult> StartAccordProductionAsync(
        int accordRecipeId, decimal plannedQty, string? notes, string? operatorName, CancellationToken ct = default)
    {
        if (accordRecipeId <= 0 || plannedQty <= 0)
            return new(false, "请选择配方和数量", 0, "");

        // 读香调配方头(RecipeID/NoteID/BatchSize)
        var accord = await _context.RecipeAccords.AsNoTracking()
            .FirstOrDefaultAsync(a => a.AccordRecipeId == accordRecipeId && a.Status == "Published", ct);
        if (accord == null || (accord.NoteId ?? 0) <= 0)
            return new(false, "无效的香调配方", 0, "");
        int noteId = accord.NoteId!.Value;
        decimal batchSize = (decimal)(accord.BatchSize > 0 ? accord.BatchSize!.Value : 100);

        // 读配方原料及当前成本/库存
        var mats = await _context.Database.SqlQueryRaw<AccordMatRow>(
            @"SELECT ram.MaterialID AS MaterialId, COALESCE(ram.MaterialName,'') AS MaterialName,
                     COALESCE(ram.Percentage,0) AS Percentage,
                     COALESCE(rmi.StockQty,0) AS StockQty,
                     COALESCE(rmi.WeightedUnitCost, COALESCE(rmi.UnitPrice,0)) AS MatCost
              FROM RecipeAccordMaterials ram
              LEFT JOIN RawMaterialInventory rmi ON ram.MaterialID = rmi.MaterialID
              WHERE ram.AccordRecipeID = {0}", accordRecipeId).ToListAsync(ct);
        if (mats.Count == 0)
            return new(false, "该配方无原料明细", 0, "");

        // 校验库存是否充足
        var shortage = new List<string>();
        foreach (var m in mats)
        {
            var need = (m.Percentage / batchSize) * plannedQty;
            if (need > m.StockQty) shortage.Add($"{m.MaterialName}(需{need:F1},存{m.StockQty:F1})");
        }
        if (shortage.Count > 0)
            return new(false, "原料库存不足: " + string.Join(" ", shortage), 0, "");

        var batchNo = $"ACP{DateTime.Now:yyyyMMddHHmm}";
        var noteName = await _context.FragranceNotes.AsNoTracking()
            .Where(f => f.NoteId == noteId).Select(f => f.NoteName).FirstOrDefaultAsync(ct) ?? "";

        using var tx = await _context.Database.BeginTransactionAsync(ct);
        try
        {
            // 建生产单头
            var prod = new AccordProduction
            {
                AccordRecipeId = accordRecipeId, NoteId = noteId, NoteName = noteName,
                BatchNo = batchNo, PlannedQty = (double)plannedQty, ActualQty = (double)plannedQty,
                Status = "Completed", WorkCenter = "SEMI", Notes = notes,
                StartedAt = DateTime.Now, CompletedAt = DateTime.Now,
                CreatedAt = DateTime.Now, UpdatedAt = DateTime.Now
            };
            await _context.AccordProductions.AddAsync(prod, ct);
            await _context.SaveChangesAsync(ct);

            // 扣原料 + OUT 流水 + 生产明细，累加本批原料总成本
            decimal totalMatCost = 0;
            foreach (var m in mats)
            {
                var need = (m.Percentage / batchSize) * plannedQty;
                if (need <= 0) continue;
                totalMatCost += need * m.MatCost;

                await _context.Database.ExecuteSqlInterpolatedAsync(
                    $"UPDATE RawMaterialInventory SET StockQty = COALESCE(StockQty,0) - {need}, UpdatedAt = {DateTime.Now} WHERE MaterialID = {m.MaterialId}", ct);

                await _ledger.WriteTransactionAsync(new InvTxn(
                    NoteId: noteId, MaterialId: m.MaterialId, ProductId: null, Quantity: -need,
                    TransactionType: "香调生产消耗", Direction: "OUT", ReferenceType: "AccordProduction",
                    ReferenceOrderId: prod.ProductionId, UnitCost: m.MatCost,
                    Notes: $"批次{batchNo}消耗[{m.MaterialName}]", CreatedBy: operatorName ?? "SYSTEM"), ct);

                await _context.Database.ExecuteSqlInterpolatedAsync($@"
INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty)
VALUES ({prod.ProductionId}, {m.MaterialId}, {m.MaterialName}, {need}, {need})", ct);
            }

            // 产出香调入库（若无行则创建）+ 结转单位成本
            var noteUnitCost = plannedQty > 0 ? totalMatCost / plannedQty : 0;
            var nowTs = DateTime.Now;
            var updated = await _context.Database.ExecuteSqlInterpolatedAsync(
                $"UPDATE NoteInventory SET StockQuantity = COALESCE(StockQuantity,0) + {plannedQty}, WeightedUnitCost = {noteUnitCost}, LastRestockDate = {nowTs}, UpdatedAt = {nowTs} WHERE NoteID = {noteId}", ct);
            if (updated == 0)
                await _context.Database.ExecuteSqlInterpolatedAsync(
                    $"INSERT INTO NoteInventory (NoteID, StockQuantity, MinStockLevel, WeightedUnitCost, LastRestockDate, UpdatedAt) VALUES ({noteId}, {plannedQty}, 50, {noteUnitCost}, {nowTs}, {nowTs})", ct);

            await _ledger.WriteTransactionAsync(new InvTxn(
                NoteId: noteId, MaterialId: null, ProductId: null, Quantity: plannedQty,
                TransactionType: "香调生产产出", Direction: "IN", ReferenceType: "AccordProduction",
                ReferenceOrderId: prod.ProductionId, UnitCost: noteUnitCost,
                Notes: $"批次{batchNo}产出香调[{noteName}]", CreatedBy: operatorName ?? "SYSTEM"), ct);

            await tx.CommitAsync(ct);
            return new(true, $"香调生产完成！批次：{batchNo}", prod.ProductionId, batchNo);
        }
        catch (Exception ex)
        {
            await tx.RollbackAsync(ct);
            return new(false, "香调生产失败: " + ex.Message, 0, "");
        }
    }

    private sealed class AccordMatRow
    {
        public int MaterialId { get; set; }
        public string MaterialName { get; set; } = "";
        public decimal Percentage { get; set; }
        public decimal StockQty { get; set; }
        public decimal MatCost { get; set; }
    }

    // ===== V21: 领料出库（扣原料 + OUT 流水）=====
    // 对标 V18 admin/semifinished/material_outbound.asp create_outbound

    public record OutboundLine(int MaterialId, decimal Qty, decimal? UnitPrice);
    public record OutboundResult(bool Success, string Message, string OutboundNo);

    /// <summary>
    /// 创建领料出库单：逐条按加权成本扣 RawMaterialInventory，写明细与 OUT 流水。
    /// </summary>
    public async Task<OutboundResult> CreateMaterialOutboundAsync(
        string outboundType, int? referenceId, string? referenceType, string? notes,
        List<OutboundLine> lines, string? operatorName, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(outboundType) || lines.Count == 0)
            return new(false, "参数错误", "");

        var outboundNo = $"OUT{DateTime.Now:yyyyMMddHHmmss}";
        using var tx = await _context.Database.BeginTransactionAsync(ct);
        try
        {
            // 建出库单头（keyless，原生 SQL + 取回 OutboundID）
            var nowTs = DateTime.Now;
            var idRows = await _context.Database.SqlQueryRaw<int>(
                @"INSERT INTO MaterialOutbound (OutboundNo, OutboundType, ReferenceID, ReferenceType, RequestedBy, OutboundDate, Status, Notes, CreatedAt)
                  VALUES ({0}, {1}, {2}, {3}, {4}, {6}, 'Confirmed', {5}, {6});
                  SELECT CAST(SCOPE_IDENTITY() AS INT) AS Value;",
                outboundNo, outboundType, (object?)referenceId ?? DBNull.Value, (object?)referenceType ?? DBNull.Value,
                (object?)(operatorName ?? "SYSTEM") ?? DBNull.Value, (object?)notes ?? DBNull.Value, nowTs).ToListAsync(ct);
            var outboundId = idRows.FirstOrDefault();
            if (outboundId <= 0) { await tx.RollbackAsync(ct); return new(false, "创建出库单失败", ""); }

            foreach (var line in lines)
            {
                if (line.MaterialId <= 0 || line.Qty <= 0) continue;
                // 出库单价优先用传入值，否则取加权成本
                var price = line.UnitPrice ?? 0;
                if (price <= 0)
                {
                    var costRows = await _context.Database.SqlQueryRaw<decimal>(
                        "SELECT CAST(COALESCE(WeightedUnitCost, COALESCE(UnitPrice,0)) AS decimal(19,4)) AS Value FROM RawMaterialInventory WHERE MaterialID = {0}",
                        line.MaterialId).ToListAsync(ct);
                    price = costRows.FirstOrDefault();
                }

                await _context.Database.ExecuteSqlInterpolatedAsync($@"
INSERT INTO MaterialOutboundDetails (OutboundID, MaterialID, RequestedQty, ActualQty, UnitPrice, TotalAmount)
VALUES ({outboundId}, {line.MaterialId}, {line.Qty}, {line.Qty}, {price}, {line.Qty * price})", ct);

                await _context.Database.ExecuteSqlInterpolatedAsync(
                    $"UPDATE RawMaterialInventory SET StockQty = COALESCE(StockQty,0) - {line.Qty}, UpdatedAt = {DateTime.Now} WHERE MaterialID = {line.MaterialId}", ct);

                await _ledger.WriteTransactionAsync(new InvTxn(
                    NoteId: 0, MaterialId: line.MaterialId, ProductId: null, Quantity: -line.Qty,
                    TransactionType: outboundType, Direction: "OUT", ReferenceType: "MaterialOutbound",
                    ReferenceOrderId: referenceId, UnitCost: price,
                    Notes: $"领料出库 {outboundNo}", CreatedBy: operatorName ?? "SYSTEM"), ct);
            }

            await tx.CommitAsync(ct);
            return new(true, $"出库成功！单号：{outboundNo}", outboundNo);
        }
        catch (Exception ex)
        {
            await tx.RollbackAsync(ct);
            return new(false, "出库处理失败: " + ex.Message, "");
        }
    }
}
