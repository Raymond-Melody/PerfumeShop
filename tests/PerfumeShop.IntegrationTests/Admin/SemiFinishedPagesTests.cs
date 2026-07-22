using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Repositories;

namespace PerfumeShop.IntegrationTests.Admin;

/// <summary>
/// M4-C 半成品管理测试
/// 覆盖：调配记录CRUD、库存查询、预警、车间转移、原料出库
/// </summary>
public class SemiFinishedPagesTests : IDisposable
{
    private readonly ProductionTestContext _db;
    private readonly SemiFinishedRepository _repo;

    public SemiFinishedPagesTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"SemiTests_{Guid.NewGuid()}")
            .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning))
            .Options;
        _db = new ProductionTestContext(options);
        _repo = new SemiFinishedRepository(_db);
    }

    public void Dispose() => _db.Dispose();

    // ===== 调配记录 =====

    [Fact]
    public async Task GetAccordProductions_ReturnsPaginated()
    {
        SeedAccordProductions(20);
        var (items, total) = await _repo.GetAccordProductionsAsync(null, 1, 10);
        Assert.Equal(20, total);
        Assert.Equal(10, items.Count());
    }

    [Fact]
    public async Task GetAccordProductions_FilterByStatus()
    {
        SeedAccordProductions(9);
        var (items, total) = await _repo.GetAccordProductionsAsync("Completed", 1, 20);
        Assert.Equal(3, total);
    }

    [Fact]
    public async Task GetAccordProductionDetail_ReturnsRecord()
    {
        SeedAccordProductions(1);
        var detail = await _repo.GetAccordProductionDetailAsync(1);
        Assert.NotNull(detail);
        Assert.Equal("BATCH-001", detail.BatchNo);
    }

    [Fact]
    public async Task CreateBlendRecord_InsertsAndReturns()
    {
        var record = new AccordProduction
        {
            BatchNo = "NEW-BATCH",
            NoteName = "Rose Accord",
            PlannedQty = 100,
            Status = "Pending"
        };
        var created = await _repo.CreateBlendRecordAsync(record);
        Assert.True(created.ProductionId > 0);
        Assert.Equal("NEW-BATCH", created.BatchNo);

        var count = await _db.AccordProductions.CountAsync();
        Assert.Equal(1, count);
    }

    [Fact]
    public async Task UpdateBlendStatus_ChangesState()
    {
        SeedAccordProductions(1);
        var ok = await _repo.UpdateBlendStatusAsync(1, "InProgress");
        Assert.True(ok);

        var rec = await _db.AccordProductions.FirstAsync();
        Assert.Equal("InProgress", rec.Status);
        Assert.NotNull(rec.StartedAt);
    }

    [Fact]
    public async Task UpdateBlendStatus_Completed_SetsCompletedAt()
    {
        SeedAccordProductions(1);
        await _repo.UpdateBlendStatusAsync(1, "Completed");
        var rec = await _db.AccordProductions.FirstAsync();
        Assert.Equal("Completed", rec.Status);
        Assert.NotNull(rec.CompletedAt);
    }

    // ===== 库存查询 =====

    [Fact]
    public async Task GetNoteInventory_ReturnsAll()
    {
        _db.NoteInventories.AddRange(
            new NoteInventory { NoteId = 1, StockQuantity = 50, MinStockLevel = 10 },
            new NoteInventory { NoteId = 2, StockQuantity = 5, MinStockLevel = 10 }
        );
        _db.SaveChanges();

        var items = await _repo.GetNoteInventoryAsync();
        Assert.Equal(2, items.Count());
    }

    [Fact]
    public async Task GetRawMaterialInventory_ReturnsAll()
    {
        _db.RawMaterialInventories.AddRange(
            new RawMaterialInventory { MaterialId = 1, ItemName = "Ethanol", StockQty = 1000, SafetyStock = 100, Unit = "L" },
            new RawMaterialInventory { MaterialId = 2, ItemName = "Bergamot Oil", StockQty = 50, SafetyStock = 100, Unit = "ml" }
        );
        _db.SaveChanges();

        var items = await _repo.GetRawMaterialInventoryAsync();
        Assert.Equal(2, items.Count());
    }

    // ===== 库存预警 =====

    [Fact]
    public async Task GetInventoryAlerts_ReturnsLowStock()
    {
        _db.NoteInventories.AddRange(
            new NoteInventory { NoteId = 1, StockQuantity = 5, MinStockLevel = 10 },
            new NoteInventory { NoteId = 2, StockQuantity = 20, MinStockLevel = 10 }
        );
        _db.RawMaterialInventories.Add(
            new RawMaterialInventory { MaterialId = 1, ItemName = "Low Item", StockQty = 10, SafetyStock = 50 }
        );
        _db.SaveChanges();

        var alerts = await _repo.GetInventoryAlertsAsync();
        Assert.Equal(2, alerts.Count()); // 1 note + 1 raw
    }

    // ===== 车间转移 =====

    [Fact]
    public async Task TransferSemiFinished_CreatesRecord()
    {
        var transfer = new WorkshopTransfer
        {
            FromWorkshop = "调配车间",
            ToWorkshop = "灌装车间",
            NoteId = 1,
            RequestQty = 50.0,
            RequestedBy = "张三"
        };

        var created = await _repo.TransferSemiFinishedAsync(transfer);
        Assert.True(created.TransferId > 0);
        Assert.Equal("Pending", created.Status);
        Assert.NotNull(created.TransferNo);
        Assert.StartsWith("TF-", created.TransferNo);
    }

    [Fact]
    public async Task GetWorkshopTransfers_ReturnsPaginated()
    {
        for (int i = 1; i <= 5; i++)
        {
            _db.WorkshopTransfers.Add(new WorkshopTransfer
            {
                TransferId = i,
                TransferNo = $"TF-{i:D4}",
                FromWorkshop = "A",
                ToWorkshop = "B",
                Status = "Pending",
                CreatedAt = DateTime.Now
            });
        }
        _db.SaveChanges();

        var (items, total) = await _repo.GetWorkshopTransfersAsync(null, 1, 3);
        Assert.Equal(5, total);
        Assert.Equal(3, items.Count());
    }

    [Fact]
    public async Task FulfillTransfer_CompletesTransfer()
    {
        _db.WorkshopTransfers.Add(new WorkshopTransfer
        {
            TransferId = 1,
            TransferNo = "TF-001",
            FromWorkshop = "A",
            ToWorkshop = "B",
            Status = "Pending",
            CreatedAt = DateTime.Now
        });
        _db.SaveChanges();

        var ok = await _repo.FulfillTransferAsync(1);
        Assert.True(ok);

        var t = await _db.WorkshopTransfers.FirstAsync();
        Assert.Equal("Completed", t.Status);
        Assert.NotNull(t.FulfilledAt);
    }

    [Fact]
    public async Task FulfillTransfer_NonExisting_ReturnsFalse()
    {
        var ok = await _repo.FulfillTransferAsync(999);
        Assert.False(ok);
    }

    // ===== 基础香料 =====

    [Fact]
    public async Task GetBaseNotes_ReturnsAll()
    {
        _db.BaseNotes.AddRange(
            new BaseNote { BaseNoteId = 1, BaseNoteName = "Sandalwood" },
            new BaseNote { BaseNoteId = 2, BaseNoteName = "Musk" }
        );
        _db.SaveChanges();

        var notes = await _repo.GetBaseNotesAsync();
        Assert.Equal(2, notes.Count());
    }

    // ===== 报表 =====

    [Fact]
    public async Task GetReportData_ReturnsDateRange()
    {
        _db.AccordProductions.AddRange(
            new AccordProduction { ProductionId = 1, BatchNo = "R1", CreatedAt = DateTime.Now.AddDays(-5) },
            new AccordProduction { ProductionId = 2, BatchNo = "R2", CreatedAt = DateTime.Now.AddDays(-2) },
            new AccordProduction { ProductionId = 3, BatchNo = "R3", CreatedAt = DateTime.Now.AddDays(-30) }
        );
        _db.SaveChanges();

        var data = await _repo.GetReportDataAsync(DateTime.Now.AddDays(-10), DateTime.Now);
        Assert.Equal(2, data.Count());
    }

    // ===== Helper =====

    private void SeedAccordProductions(int count)
    {
        for (int i = 1; i <= count; i++)
        {
            _db.AccordProductions.Add(new AccordProduction
            {
                ProductionId = i,
                BatchNo = $"BATCH-{i:D3}",
                NoteName = $"Accord #{i}",
                PlannedQty = 100.0,
                Status = i % 3 == 0 ? "Completed" : i % 3 == 1 ? "Pending" : "InProgress",
                CreatedAt = DateTime.Now.AddDays(-count + i)
            });
        }
        _db.SaveChanges();
    }
}
