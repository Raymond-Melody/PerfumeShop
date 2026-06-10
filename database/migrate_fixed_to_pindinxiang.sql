-- ============================================
-- 迁移脚本: 品牌定香产品管理流程重构
-- 将 "固定品牌" 重命名为 "品牌定香"
-- 日期: 2026-06-02
-- ============================================

-- 更新 ProductTypeConfig 表中 TypeCode='Fixed' 的记录
-- 此脚本适用于 SQL Server
UPDATE ProductTypeConfig 
SET DisplayName = N'品牌定香',
    NavName = N'品牌定香',
    Description = N'品牌方预设的固定香型成品，通过采购模块进行入库管理，区别于用户自定香型组合',
    Icon = N'fas fa-box'
WHERE TypeCode = 'Fixed';

-- 验证更新结果
SELECT ConfigID, TypeCode, DisplayName, NavName, Description, Icon, IsActive, RequiresRatio, RequiresReview
FROM ProductTypeConfig
WHERE TypeCode = 'Fixed';

-- ============================================
-- 数据一致性检查
-- ============================================

-- 1. 检查 FixedBrandProducts 表的 ProductID 字段是否存在
-- 如果以下查询报错，说明 ProductID 字段缺失，需执行 ALTER TABLE
-- ALTER TABLE FixedBrandProducts ADD ProductID INT NULL;
SELECT COLUMN_NAME, DATA_TYPE 
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'FixedBrandProducts' AND COLUMN_NAME = 'ProductID';

-- 2. 检查 FixedBrandProducts 中 ProductID 为 NULL 的记录（孤立记录）
-- 这些记录有 FixedBrandProducts 但没有对应的 Products 行
SELECT FBP.FixedProductID, FBP.ProductName, FBP.BasePrice, FBP.ProductID
FROM FixedBrandProducts FBP
WHERE FBP.Status = 'Active' AND (FBP.ProductID IS NULL OR FBP.ProductID = 0);

-- 3. 手动修复孤立记录：为没有 ProductID 的 FixedBrandProducts 创建 Products 行
-- 取消注释以下代码以执行自动修复（请先备份数据库）
/*
DECLARE @fpid INT, @pname NVARCHAR(200), @baseprice DECIMAL(10,2), @isactive BIT, @desc NVARCHAR(MAX)
DECLARE cur CURSOR FOR 
    SELECT FixedProductID, ProductName, ISNULL(BasePrice,0), 
           CASE WHEN Status='Active' THEN 1 ELSE 0 END,
           ISNULL(Description,'品牌定香产品')
    FROM FixedBrandProducts 
    WHERE Status='Active' AND (ProductID IS NULL OR ProductID = 0)
OPEN cur
FETCH NEXT FROM cur INTO @fpid, @pname, @baseprice, @isactive, @desc
WHILE @@FETCH_STATUS = 0
BEGIN
    DECLARE @newPid INT
    INSERT INTO Products (ProductName, ProductType, BasePrice, BaseIngredients, IsActive, Description, CreatedAt, UpdatedAt)
    VALUES (@pname, 'Fixed', @baseprice, N'品牌定香成品', @isactive, @desc, GETDATE(), GETDATE())
    SET @newPid = SCOPE_IDENTITY()
    UPDATE FixedBrandProducts SET ProductID = @newPid WHERE FixedProductID = @fpid
    FETCH NEXT FROM cur INTO @fpid, @pname, @baseprice, @isactive, @desc
END
CLOSE cur
DEALLOCATE cur
*/

-- 4. 检查 Products 表中 Fixed 类型产品与 FixedBrandProducts 的关联一致性
SELECT 
    P.ProductID, P.ProductName, P.UnitCost, P.IsActive,
    FBP.FixedProductID, FBP.ProductName AS FBPName,
    FBI.AvgUnitCost AS InventoryCost
FROM Products P
LEFT JOIN FixedBrandProducts FBP ON P.ProductID = FBP.ProductID AND FBP.Status = 'Active'
LEFT JOIN FixedBrandInventory FBI ON FBP.FixedProductID = FBI.FixedProductID
WHERE P.ProductType = 'Fixed'
ORDER BY P.ProductID;

-- 5. 检查成本不一致：Products.UnitCost 与 FixedBrandInventory.AvgUnitCost 的偏差
SELECT 
    P.ProductID, P.ProductName, P.UnitCost AS ProductsCost,
    FBI.AvgUnitCost AS InventoryCost,
    ABS(ISNULL(P.UnitCost,0) - ISNULL(FBI.AvgUnitCost,0)) AS CostDiff
FROM Products P
JOIN FixedBrandProducts FBP ON P.ProductID = FBP.ProductID AND FBP.Status = 'Active'
JOIN FixedBrandInventory FBI ON FBP.FixedProductID = FBI.FixedProductID
WHERE ABS(ISNULL(P.UnitCost,0) - ISNULL(FBI.AvgUnitCost,0)) > 0.01
ORDER BY CostDiff DESC;

-- ============================================
-- 迁移完成标记
-- ============================================
PRINT N'品牌定香迁移脚本执行完成。请检查以上输出确认所有数据一致性。';
