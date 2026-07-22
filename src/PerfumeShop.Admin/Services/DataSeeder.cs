using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Admin.Services;

/// <summary>
/// 测试数据种子服务 - 一次性注入成本/分摊/财务/采购测试数据
/// </summary>
public class DataSeeder
{
    private readonly PerfumeShopContext _db;
    public DataSeeder(PerfumeShopContext db) => _db = db;

    public async Task<string> SeedAllAsync()
    {
        var log = new System.Text.StringBuilder();
        try
        {
            log.AppendLine(await SeedProductCostsAsync());
            log.AppendLine(await SeedRawMaterialCostsAsync());
            log.AppendLine(await SeedFragranceNoteCostsAsync());
            log.AppendLine(await SeedProductCostDetailsAsync());
            log.AppendLine(await SeedOrderCostAllocationAsync());
            log.AppendLine(await SeedExpenseRecordsAsync());
            log.AppendLine(await SeedAccountsAsync());
            log.AppendLine(await SeedBudgetPlansAsync());
            log.AppendLine(await SeedFundAccountsAsync());
            log.AppendLine(await SeedPurchaseDataAsync());
            log.AppendLine(await SeedSupplierPricesAsync());
            await _db.SaveChangesAsync();
            log.AppendLine("=== 全部种子数据注入完成 ===");
        }
        catch (Exception ex) { log.AppendLine($"错误: {ex.Message}"); }
        return log.ToString();
    }

    async Task<string> SeedProductCostsAsync()
    {
        var cnt = 0;
        var products = await _db.Products.Take(10).ToListAsync();
        foreach (var p in products)
        {
            if (p.UnitCost == null || p.UnitCost == 0)
            {
                p.UnitCost = (decimal)(p.ProductId * 15.5 + 30);
                cnt++;
            }
        }
        return $"产品成本更新: {cnt} 个字段";
    }

    async Task<string> SeedRawMaterialCostsAsync()
    {
        var cnt = 0;
        var materials = await _db.RawMaterialInventories.ToListAsync();
        foreach (var m in materials)
        {
            if (m.UnitPrice == null || m.UnitPrice == 0)
            {
                m.UnitPrice = (decimal)(m.MaterialId * 8.8 + 50);
                cnt++;
            }
            if (m.WeightedUnitCost == null || m.WeightedUnitCost == 0)
            {
                m.WeightedUnitCost = (decimal)(m.MaterialId * 7.5 + 40);
                cnt++;
            }
            if (m.StockQty == null || m.StockQty == 0) m.StockQty = 500 + m.MaterialId * 20;
            if (m.SafetyStock == null || m.SafetyStock == 0) m.SafetyStock = 100;
        }
        return $"原料成本/库存更新: {cnt} 个字段";
    }

    async Task<string> SeedFragranceNoteCostsAsync()
    {
        var cnt = 0;
        var notes = await _db.FragranceNotes.ToListAsync();
        foreach (var n in notes)
        {
            if (n.PriceAddition == null || n.PriceAddition == 0)
            {
                n.PriceAddition = (decimal)(n.NoteId * 5.5 + 15);
                cnt++;
            }
        }
        return $"香调价格更新: {cnt} 个";
    }

    async Task<string> SeedProductCostDetailsAsync()
    {
        var cnt = 0;
        var existing = await _db.ProductCosts.AnyAsync();
        if (!existing)
        {
            var products = await _db.Products.Take(8).ToListAsync();
            var costTypes = new[] { "BOM", "Purchase", "Packaging", "Labor", "Shipping", "Other" };
            int i = 0;
            foreach (var p in products)
            {
                foreach (var ct in costTypes.Take(3 + i % 3))
                {
                    _db.ProductCosts.Add(new ProductCost
                    {
                        ProductId = p.ProductId,
                        CostType = ct,
                        CostName = $"{ct}成本-{p.ProductName}",
                        UnitCost = (decimal)(p.ProductId * 2.5 + i * 10 + 5),
                        Quantity = 10 + i * 5,
                        TotalCost = (decimal)((p.ProductId * 2.5 + i * 10 + 5) * (10 + i * 5)),
                        EffectiveDate = DateTime.Today.AddDays(-30)
                    });
                    cnt++;
                }
                i++;
            }
        }
        return $"ProductCosts: 新增 {cnt} 条";
    }

    async Task<string> SeedOrderCostAllocationAsync()
    {
        var cnt = 0;
        var existing = await _db.OrderCostAllocations.AnyAsync();
        if (!existing)
        {
            var orders = await _db.Orders.Where(o => o.Status == "Completed" || o.Status == "completed").Take(5).ToListAsync();
            var types = new[] { "Shipping", "Packaging", "Platform", "Marketing", "Labor" };
            int i = 0;
            foreach (var o in orders)
            {
                foreach (var t in types.Take(2 + i % 3))
                {
                    _db.OrderCostAllocations.Add(new OrderCostAllocation
                    {
                        OrderId = o.OrderId,
                        CostType = t,
                        ItemName = $"{t}费用",
                        TotalCost = (decimal)(o.TotalAmount * 0.05m + i * 3),
                        AllocatedAt = DateTime.Today.AddDays(-i * 5)
                    });
                    cnt++;
                }
                i++;
            }
        }
        return $"OrderCostAllocation: 新增 {cnt} 条";
    }

    async Task<string> SeedExpenseRecordsAsync()
    {
        var cnt = 0;
        var existing = await _db.ExpenseRecords.AnyAsync();
        if (!existing)
        {
            var expenses = new[]
            {
                ("运费分摊", "Shipping", 350.00m),
                ("运费分摊-2月", "Shipping", 280.50m),
                ("推广费用-抖音", "Marketing", 1200.00m),
                ("推广费用-小红书", "Marketing", 800.00m),
                ("平台佣金", "Platform", 450.00m),
                ("平台服务费", "Platform", 199.00m),
                ("包装耗材", "Packaging", 180.00m),
                ("人工成本", "Labor", 2500.00m),
                ("仓储租金", "Storage", 1800.00m),
                ("物流配送", "Logistics", 620.00m),
            };
            int i = 0;
            foreach (var (name, type, amount) in expenses)
            {
                _db.ExpenseRecords.Add(new ExpenseRecord
                {
                    ExpenseName = name,
                    ExpenseType = type,
                    Amount = amount,
                    CreatedAt = DateTime.Now,
                    Period = DateTime.Today.ToString("yyyy-MM")
                });
                cnt++; i++;
            }
        }
        return $"ExpenseRecords: 新增 {cnt} 条";
    }

    async Task<string> SeedAccountsAsync()
    {
        var cnt = 0;
        if (!await _db.AccountsPayables.AnyAsync())
        {
            _db.AccountsPayables.AddRange(
                new AccountsPayable { SupplierName = "香料供应商A", Amount = 2500.00m, DueDate = DateTime.Today.AddDays(30), Status = "未付", CreatedAt = DateTime.Now },
                new AccountsPayable { SupplierName = "瓶子供应商B", Amount = 1800.00m, DueDate = DateTime.Today.AddDays(15), Status = "未付", CreatedAt = DateTime.Now },
                new AccountsPayable { SupplierName = "包装供应商C", Amount = 950.00m, DueDate = DateTime.Today.AddDays(-5), Status = "已付", PaidAmount = 950.00m, CreatedAt = DateTime.Now }
            );
            cnt += 3;
        }
        if (!await _db.AccountsReceivables.AnyAsync())
        {
            _db.AccountsReceivables.AddRange(
                new AccountsReceivable { CustomerName = "批发客户D", Amount = 3200.00m, DueDate = DateTime.Today.AddDays(45), Status = "未收", CreatedAt = DateTime.Now },
                new AccountsReceivable { CustomerName = "企业客户E", Amount = 1500.00m, DueDate = DateTime.Today.AddDays(20), Status = "未收", CreatedAt = DateTime.Now }
            );
            cnt += 2;
        }
        return $"应收应付: 新增 {cnt} 条";
    }

    async Task<string> SeedBudgetPlansAsync()
    {
        var cnt = 0;
        if (!await _db.BudgetPlans.AnyAsync())
        {
            var sql = @"INSERT INTO BudgetPlans (BudgetName, Category, BudgetAmount, ActualAmount, Period, Status) VALUES
('2026Q1 运营预算', '运营', 50000, 32000, '2026Q1', '执行中'),
('2026Q2 市场推广', '推广', 30000, 18500, '2026Q2', '执行中'),
('2026年度研发', '研发', 80000, 25000, '2026', '执行中')";
            cnt = await _db.Database.ExecuteSqlRawAsync(sql);
        }
        return $"BudgetPlans: 新增 {cnt} 条";
    }

    async Task<string> SeedFundAccountsAsync()
    {
        var cnt = 0;
        if (!await _db.FundAccounts.AnyAsync())
        {
            _db.FundAccounts.AddRange(
                new FundAccount { AccountName = "基本户-工商银行", AccountType = "银行", TotalBalance = 50000, AlertThreshold = 5000, AvailableBalance = 48000, CreatedAt = DateTime.Now, IsActive = true },
                new FundAccount { AccountName = "支付宝商户", AccountType = "第三方支付", TotalBalance = 25000, AlertThreshold = 3000, AvailableBalance = 22000, CreatedAt = DateTime.Now, IsActive = true },
                new FundAccount { AccountName = "微信商户", AccountType = "第三方支付", TotalBalance = 18000, AlertThreshold = 2000, AvailableBalance = 16500, CreatedAt = DateTime.Now, IsActive = true }
            );
            cnt = 3;
        }
        return $"FundAccounts: 新增 {cnt} 条";
    }

    async Task<string> SeedPurchaseDataAsync()
    {
        var cnt = 0;
        if (!await _db.FixedBrandCostAllocations.AnyAsync())
        {
            _db.FixedBrandCostAllocations.AddRange(
                new FixedBrandCostAllocation { OrderId = 1, PurchaseId = 1, FixedProductId = 1, ProductName = "品牌定香产品A", CostPerUnit = 450.00m, Quantity = 10, TotalCost = 4500.00m, SalePrice = 680.00m, ProfitAmount = 2300.00m, ProfitRate = 0.51m, AllocatedAt = DateTime.Today.AddDays(-20) },
                new FixedBrandCostAllocation { OrderId = 2, PurchaseId = 2, FixedProductId = 2, ProductName = "品牌定香产品B", CostPerUnit = 320.00m, Quantity = 5, TotalCost = 1600.00m, SalePrice = 480.00m, ProfitAmount = 800.00m, ProfitRate = 0.50m, AllocatedAt = DateTime.Today.AddDays(-15) }
            );
            cnt += 2;
        }
        if (!await _db.PurchaseReceipts.AnyAsync())
        {
            var sql = @"INSERT INTO PurchaseReceipts (PurchaseId, ReceiptNo, TotalReceivedQty, ReceiptDate, Status, CreatedAt) VALUES
(1, 'RCV-20260701', 100, GETDATE()-10, '已入库', GETDATE()),
(2, 'RCV-20260702', 50, GETDATE()-5, '已入库', GETDATE())";
            cnt += await _db.Database.ExecuteSqlRawAsync(sql);
        }
        return $"采购数据: 新增 {cnt} 条";
    }

    async Task<string> SeedSupplierPricesAsync()
    {
        var cnt = 0;
        if (!await _db.SupplierPrices.AnyAsync())
        {
            _db.SupplierPrices.AddRange(
                new SupplierPrice { SupplierId = 1, ItemName = "玫瑰精油", UnitPrice = 85.50m, EffectiveDate = DateTime.Today.AddDays(-30), PriceType = "采购价" },
                new SupplierPrice { SupplierId = 1, ItemName = "茉莉净油", UnitPrice = 120.00m, EffectiveDate = DateTime.Today.AddDays(-30), PriceType = "采购价" },
                new SupplierPrice { SupplierId = 2, ItemName = "玻璃瓶50ml", UnitPrice = 12.50m, EffectiveDate = DateTime.Today.AddDays(-20), PriceType = "采购价" },
                new SupplierPrice { SupplierId = 2, ItemName = "玻璃瓶100ml", UnitPrice = 18.00m, EffectiveDate = DateTime.Today.AddDays(-20), PriceType = "采购价" }
            );
            cnt = 4;
        }
        return $"SupplierPrices: 新增 {cnt} 条";
    }
}
