-- ============================================================
-- V19 成本数据修复种子脚本（路线A + 路线B源数据）
-- 所有插入行以 CreatedBy='seed_v19' 标记，便于回滚：
--   DELETE FROM ProductCosts WHERE CreatedBy='seed_v19';
-- 可重复执行（先清理旧种子再插入）
-- ============================================================
SET NOCOUNT ON;
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET XACT_ABORT ON;
BEGIN TRAN;

-- ============ 路线B · 源数据补齐（让成本引擎可算出非零 BOM）============
-- 1) 原料单价（元/克或元/ml），仅填补缺失(=0)的
UPDATE RawMaterialInventory
SET UnitPrice = CASE ItemCode
        WHEN 'RM-ROSE-001'  THEN 2.80
        WHEN 'RM-JASM-001'  THEN 3.50
        WHEN 'RM-BERG-001'  THEN 1.20
        WHEN 'RM-LEMON-001' THEN 0.90
        WHEN 'RM-SAND-001'  THEN 4.20
        ELSE 1.50 END
WHERE ISNULL(UnitPrice,0) = 0;

UPDATE RawMaterialInventory
SET WeightedUnitCost = UnitPrice
WHERE ISNULL(WeightedUnitCost,0) = 0 AND ISNULL(UnitPrice,0) > 0;

-- 2) 香调 PriceAddition 兜底（每 ml 成本），对无值的活跃香调设置
UPDATE FragranceNotes
SET PriceAddition = CAST(0.50 + (ABS(NoteID) % 10) * 0.15 AS DECIMAL(19,4))
WHERE IsActive = 1 AND ISNULL(PriceAddition,0) = 0;

-- 3) 基香单价兜底
UPDATE BaseNotes
SET UnitPrice = CAST(0.80 + (ABS(BaseNoteID) % 8) * 0.20 AS DECIMAL(19,4))
WHERE ISNULL(UnitPrice,0) = 0;

-- ============ 路线A · ProductCosts 录入 + Products 回写 ============
-- 清理旧种子（可重跑）
DELETE FROM ProductCosts WHERE CreatedBy = 'seed_v19';

-- 本月成本项：Purchase / Packaging / Other / BOM（按 BasePrice 比例）
INSERT INTO ProductCosts (ProductID, CostType, CostName, UnitCost, Quantity, TotalCost, EffectiveDate, CreatedBy, CreatedAt)
SELECT ProductID, 'Purchase',  N'原料采购',      CAST(BasePrice*0.06 AS DECIMAL(19,4)), 1, CAST(BasePrice*0.06 AS DECIMAL(19,4)), CAST(GETDATE() AS DATE), 'seed_v19', GETDATE() FROM Products WHERE IsActive=1
UNION ALL
SELECT ProductID, 'Packaging', N'瓶身与包材',    CAST(BasePrice*0.08 AS DECIMAL(19,4)), 1, CAST(BasePrice*0.08 AS DECIMAL(19,4)), CAST(GETDATE() AS DATE), 'seed_v19', GETDATE() FROM Products WHERE IsActive=1
UNION ALL
SELECT ProductID, 'Other',     N'人工与管理分摊', CAST(BasePrice*0.04 AS DECIMAL(19,4)), 1, CAST(BasePrice*0.04 AS DECIMAL(19,4)), CAST(GETDATE() AS DATE), 'seed_v19', GETDATE() FROM Products WHERE IsActive=1
UNION ALL
SELECT ProductID, 'BOM',       N'配方料体',      CAST(BasePrice*0.22 AS DECIMAL(19,4)), 1, CAST(BasePrice*0.22 AS DECIMAL(19,4)), CAST(GETDATE() AS DATE), 'seed_v19', GETDATE() FROM Products WHERE IsActive=1;

-- 回写 Products（BOM=0.22×价, UnitCost=BOM+Packaging+Other=0.34×价）
-- 注：设置 UnitCost 使引擎对 standard 型产品保留该值（existingUnitCost 分支）
UPDATE Products
SET BOMCost  = CAST(BasePrice*0.22 AS DECIMAL(19,4)),
    UnitCost = CAST(BasePrice*0.34 AS DECIMAL(19,4))
WHERE IsActive = 1;

-- ============ 订单成本与利润（与 CE_UpdateOrderCosts 一致的公式）============
-- 成本 = Σ(明细数量 × 产品单位成本)；利润 = 金额 - 成本 - 运费（下限 0）
UPDATE o
SET CostAmount = ISNULL(x.Cost,0),
    ProfitAmount = CASE WHEN (ISNULL(o.TotalAmount,0) - ISNULL(x.Cost,0) - ISNULL(o.ShippingFee,0)) < 0
                        THEN 0 ELSE (ISNULL(o.TotalAmount,0) - ISNULL(x.Cost,0) - ISNULL(o.ShippingFee,0)) END
FROM Orders o
LEFT JOIN (
    SELECT od.OrderID, SUM(od.Quantity * ISNULL(p.UnitCost,0)) AS Cost
    FROM OrderDetails od LEFT JOIN Products p ON od.ProductID = p.ProductID
    GROUP BY od.OrderID
) x ON o.OrderID = x.OrderID
WHERE o.Status NOT IN ('Pending','Cancelled');

-- 重建订单成本分摊明细（测试库安全全量重建）
DELETE FROM OrderCostAllocation;
INSERT INTO OrderCostAllocation (OrderID, OrderNo, CostType, ItemCode, ItemName, UnitCost, Quantity, TotalCost, AllocatedAt, CreatedAt)
SELECT o.OrderID, o.OrderNo, 'Product', CAST(od.ProductID AS NVARCHAR(50)), p.ProductName,
       ISNULL(p.UnitCost,0), od.Quantity, od.Quantity * ISNULL(p.UnitCost,0), GETDATE(), GETDATE()
FROM Orders o
JOIN OrderDetails od ON o.OrderID = od.OrderID
LEFT JOIN Products p ON od.ProductID = p.ProductID
WHERE o.Status NOT IN ('Pending','Cancelled') AND ISNULL(p.UnitCost,0) > 0;

-- 计价方式默认值（供顶部横幅显示）
IF NOT EXISTS (SELECT 1 FROM SiteSettings WHERE SettingKey='CostCalculationMethod')
    INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('CostCalculationMethod','WEIGHTED');

COMMIT TRAN;

-- ============ 结果校验 ============
PRINT '=== 种子结果 ===';
SELECT 'ProductCosts(seed)' AS Item, COUNT(*) AS Cnt FROM ProductCosts WHERE CreatedBy='seed_v19'
UNION ALL SELECT 'Products.UnitCost>0', COUNT(*) FROM Products WHERE IsActive=1 AND ISNULL(UnitCost,0)>0
UNION ALL SELECT 'Products.BOMCost>0', COUNT(*) FROM Products WHERE IsActive=1 AND ISNULL(BOMCost,0)>0
UNION ALL SELECT 'Materials.priced', COUNT(*) FROM RawMaterialInventory WHERE ISNULL(WeightedUnitCost,0)>0
UNION ALL SELECT 'Notes.PriceAdd>0', COUNT(*) FROM FragranceNotes WHERE IsActive=1 AND ISNULL(PriceAddition,0)>0;
