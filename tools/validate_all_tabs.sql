SET NOCOUNT ON;
PRINT '========== TAB 1: 商品成本维护 ==========';
SELECT COUNT(*) AS Products, SUM(CASE WHEN ISNULL(UnitCost,0)>0 THEN 1 ELSE 0 END) AS WithUnit,
       SUM(CASE WHEN ISNULL(BOMCost,0)>0 THEN 1 ELSE 0 END) AS WithBOM FROM Products WHERE IsActive=1;
SELECT TOP 3 p.ProductID, p.ProductName, p.BOMCost,
       (SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID=p.ProductID AND CostType='Purchase') AS Purchase,
       (SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID=p.ProductID AND CostType='Packaging') AS Packaging,
       (SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID=p.ProductID AND CostType='Other') AS Other,
       p.UnitCost FROM Products p WHERE IsActive=1 ORDER BY p.ProductID;

PRINT '========== TAB 2: 成本异动监控（本月 vs 上月）==========';
DECLARE @cm VARCHAR(6)=CONVERT(VARCHAR(6),GETDATE(),112), @lm VARCHAR(6)=CONVERT(VARCHAR(6),DATEADD(MONTH,-1,GETDATE()),112);
SELECT TOP 6 p.ProductID,
   (SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID=p.ProductID AND CONVERT(VARCHAR(6),CreatedAt,112)=@lm) AS LastM,
   (SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID=p.ProductID AND CONVERT(VARCHAR(6),CreatedAt,112)=@cm) AS CurrM
FROM Products p WHERE IsActive=1 ORDER BY p.ProductID;
SELECT SUM(CASE WHEN CONVERT(VARCHAR(6),CreatedAt,112)=@cm THEN 1 ELSE 0 END) AS ThisMonthRows,
       SUM(CASE WHEN CONVERT(VARCHAR(6),CreatedAt,112)=@lm THEN 1 ELSE 0 END) AS LastMonthRows FROM ProductCosts;

PRINT '========== TAB 3: 成本变更历史 ==========';
SELECT COUNT(*) AS HistoryRows, COUNT(DISTINCT ProductID) AS Products,
       MIN(CreatedAt) AS Earliest, MAX(CreatedAt) AS Latest FROM ProductCosts;

PRINT '========== TAB 4: 成本传导链（源数据）==========';
SELECT 'ActiveNotes' AS Item, COUNT(*) AS Cnt FROM FragranceNotes WHERE IsActive=1
UNION ALL SELECT 'NotesPriced', COUNT(*) FROM FragranceNotes WHERE IsActive=1 AND ISNULL(PriceAddition,0)>0
UNION ALL SELECT 'MaterialsPriced', COUNT(*) FROM RawMaterialInventory WHERE ISNULL(WeightedUnitCost,0)>0
UNION ALL SELECT 'ProductsWithCost', COUNT(*) FROM Products WHERE IsActive=1 AND ISNULL(UnitCost,0)>0;

PRINT '========== 订单成本/利润/分摊 ==========';
SELECT COUNT(*) AS EligibleOrders,
       SUM(CASE WHEN ISNULL(CostAmount,0)>0 THEN 1 ELSE 0 END) AS WithCost,
       SUM(CASE WHEN ISNULL(ProfitAmount,0)>0 THEN 1 ELSE 0 END) AS WithProfit
FROM Orders WHERE Status NOT IN ('Pending','Cancelled');
SELECT COUNT(*) AS AllocationRows FROM OrderCostAllocation;
