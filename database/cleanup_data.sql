-- ============================================
-- 全流程模拟 - 数据清理脚本
-- 清空所有业务流水，保留基础配置
-- ============================================
USE PerfumeShop;
GO

PRINT '===== 开始数据清理 =====';
GO

-- 1. 订单相关（按依赖顺序）
PRINT '清理订单相关数据...';
DELETE FROM OrderIngredients;
DELETE FROM OrderDetailNoteSelections;
DELETE FROM OrderDetails;
DELETE FROM Orders;
DELETE FROM CartNoteSelections;
DELETE FROM Cart;
GO

-- 2. 生产相关
PRINT '清理生产相关数据...';
DELETE FROM ProductManufacturingDetails;
DELETE FROM ProductManufacturing;
DELETE FROM ProductionOrders;
DELETE FROM ProductionLogs;
DELETE FROM AccordProductionDetails;
DELETE FROM AccordProductions;
DELETE FROM AccordQCReports;
GO

-- 3. 采购流水
PRINT '清理采购流水数据...';
DELETE FROM PurchaseReceiptDetails;
DELETE FROM PurchaseReceipts;
DELETE FROM PurchaseOrderDetails;
DELETE FROM PurchaseOrders;
DELETE FROM PurchaseCostReview;
GO

-- 4. 库存变动
PRINT '清理库存变动数据...';
DELETE FROM MaterialOutboundDetails;
DELETE FROM MaterialOutbound;
DELETE FROM InventoryTransactions;
DELETE FROM WorkshopTransfer;
GO

-- 5. 财务流水
PRINT '清理财务流水数据...';
DELETE FROM ExpenseRecords;
DELETE FROM PaymentRecords;
DELETE FROM GLTransactions;
DELETE FROM AccountsPayable;
DELETE FROM AccountsReceivable;
DELETE FROM RefundRecords;
DELETE FROM ReconciliationLogs;
DELETE FROM DailyStatistics;
GO

-- 6. 营销数据
PRINT '清理营销数据...';
DELETE FROM MarketingCampaigns;
DELETE FROM Coupons;
GO

-- 7. 用户数据（保留管理员 Ray88）
PRINT '清理用户数据（保留管理员）...';
DELETE FROM UserFavorites;
DELETE FROM ProductReviews;
DELETE FROM PointTransactions;
DELETE FROM UserPoints;
DELETE FROM UserPreferences;
DELETE FROM UserAddresses;
-- 保留 AdminUsers 中的 Ray88，删除前端用户
DELETE FROM Users WHERE Username <> 'Ray88';
GO

-- 8. 库存重置
PRINT '重置库存...';
UPDATE ProductInventory SET StockQty = 0;
UPDATE NoteInventory SET StockQuantity = 0;
UPDATE RawMaterialInventory SET StockQty = 0;
GO

-- 9. 清理其他可能残留的数据
PRINT '清理其他残留数据...';
DELETE FROM AdminLogs WHERE ModuleCode NOT IN ('SYSTEM','SECURITY');
DELETE FROM RecipePublishLog;
DELETE FROM RecipePopularity;
GO

PRINT '===== 数据清理完成 =====';
GO

-- ============================================
-- 验证
-- ============================================
PRINT '===== 验证清理结果 =====';

SELECT 'Orders' AS TableName, COUNT(*) AS RowCount FROM Orders
UNION ALL
SELECT 'OrderDetails', COUNT(*) FROM OrderDetails
UNION ALL
SELECT 'Cart', COUNT(*) FROM Cart
UNION ALL
SELECT 'ProductionOrders', COUNT(*) FROM ProductionOrders
UNION ALL
SELECT 'AccordProductions', COUNT(*) FROM AccordProductions
UNION ALL
SELECT 'ProductManufacturing', COUNT(*) FROM ProductManufacturing
UNION ALL
SELECT 'PurchaseOrders', COUNT(*) FROM PurchaseOrders
UNION ALL
SELECT 'PurchaseReceipts', COUNT(*) FROM PurchaseReceipts
UNION ALL
SELECT 'MaterialOutbound', COUNT(*) FROM MaterialOutbound
UNION ALL
SELECT 'InventoryTransactions', COUNT(*) FROM InventoryTransactions
UNION ALL
SELECT 'PaymentRecords', COUNT(*) FROM PaymentRecords
UNION ALL
SELECT 'GLTransactions', COUNT(*) FROM GLTransactions
UNION ALL
SELECT 'ExpenseRecords', COUNT(*) FROM ExpenseRecords
UNION ALL
SELECT 'MarketingCampaigns', COUNT(*) FROM MarketingCampaigns
UNION ALL
SELECT 'Coupons', COUNT(*) FROM Coupons
UNION ALL
SELECT 'Users', COUNT(*) FROM Users
UNION ALL
SELECT 'Products', COUNT(*) FROM Products
UNION ALL
SELECT 'FragranceNotes', COUNT(*) FROM FragranceNotes
UNION ALL
SELECT 'BaseNotes', COUNT(*) FROM BaseNotes
UNION ALL
SELECT 'Suppliers', COUNT(*) FROM Suppliers
UNION ALL
SELECT 'BottleStyles', COUNT(*) FROM BottleStyles
UNION ALL
SELECT 'Volumes', COUNT(*) FROM Volumes;
GO
