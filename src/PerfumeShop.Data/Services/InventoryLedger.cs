using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// V21 库存流水统一写入器 — 对标 V18 各页面内散落的 InventoryTransactions/StockMovements INSERT。
/// 设计要点：
///  1. InventoryTransactions / StockMovements 均为原生 SQL 写入（前者 keyless，无法经 EF 跟踪保存；
///     StockMovements 亦用原生 SQL 保持一致，且不触发 SaveChanges，避免过早提交调用方未存更改）。
///  2. 不自开事务：所有方法在调用方(仓储)的 EF 事务/同一连接内执行，与库存增减保持原子。
///  3. StockMovements 为审计快照，写失败仅吞异常不阻断主库存事务（表/列可能不齐）。
/// </summary>
public interface IInventoryLedger
{
    /// <summary>写库存交易流水 (InventoryTransactions)。Direction: IN/OUT；数量按方向传正/负值由调用方决定。</summary>
    Task WriteTransactionAsync(InvTxn txn, CancellationToken ct = default);

    /// <summary>写库存移动快照 (StockMovements，含前后量)。</summary>
    Task WriteMovementAsync(StockMove move, CancellationToken ct = default);
}

/// <summary>库存交易流水参数。NoteId 不可空(非香调交易传 0，对齐 V18)；MaterialId/ProductId 视品类而定。</summary>
public sealed record InvTxn(
    int NoteId,
    int? MaterialId,
    int? ProductId,
    decimal Quantity,
    string TransactionType,
    string Direction,
    string? ReferenceType,
    int? ReferenceOrderId,
    decimal? UnitCost,
    string? Notes,
    string? CreatedBy);

/// <summary>库存移动快照参数。</summary>
public sealed record StockMove(
    string ItemType,
    int ItemId,
    string? ItemName,
    string? ItemCode,
    string MovementType,
    decimal Quantity,
    decimal BeforeQty,
    decimal AfterQty,
    string? Unit,
    string? ReferenceNo,
    string? Notes,
    string? CreatedBy);

public class InventoryLedger : IInventoryLedger
{
    private readonly PerfumeShopContext _db;

    public InventoryLedger(PerfumeShopContext db) => _db = db;

    public async Task WriteTransactionAsync(InvTxn t, CancellationToken ct = default)
    {
        var now = DateTime.Now;
        await _db.Database.ExecuteSqlInterpolatedAsync($@"
INSERT INTO InventoryTransactions
    (NoteID, MaterialID, ProductID, Quantity, TransactionType, TransactionDirection,
     ReferenceType, ReferenceOrderID, UnitCost, Notes, CreatedBy, CreatedAt)
VALUES
    ({t.NoteId}, {t.MaterialId}, {t.ProductId}, {t.Quantity}, {t.TransactionType}, {t.Direction},
     {t.ReferenceType}, {t.ReferenceOrderId}, {t.UnitCost}, {t.Notes}, {t.CreatedBy}, {now})", ct);
    }

    public async Task WriteMovementAsync(StockMove m, CancellationToken ct = default)
    {
        try
        {
            var now = DateTime.Now;
            await _db.Database.ExecuteSqlInterpolatedAsync($@"
INSERT INTO StockMovements
    (ItemType, ItemID, ItemName, ItemCode, MovementType, Quantity, BeforeQty, AfterQty,
     Unit, ReferenceNo, Notes, CreatedBy, CreatedAt)
VALUES
    ({m.ItemType}, {m.ItemId}, {m.ItemName}, {m.ItemCode}, {m.MovementType}, {m.Quantity}, {m.BeforeQty}, {m.AfterQty},
     {m.Unit}, {m.ReferenceNo}, {m.Notes}, {m.CreatedBy}, {now})", ct);
        }
        catch
        {
            // StockMovements 为审计快照，写失败不阻断主库存事务
        }
    }
}
