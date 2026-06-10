USE PerfumeShop;

-- Create Purchase Receipt
DECLARE @receiptId INT;
INSERT INTO PurchaseReceipts (PurchaseID, SupplierID, ReceiptNo, ReceiptDate, Status, ReceivedBy, TotalReceivedQty, Notes, CreatedAt)
VALUES (14, 1, 'REC-20260517-001', GETDATE(), '已收货', 'Ray88', 5500, '模拟收货入库', GETDATE());
SET @receiptId = SCOPE_IDENTITY();
PRINT 'Receipt ID: ' + CAST(@receiptId AS NVARCHAR);

-- Create Receipt Details for each item
INSERT INTO PurchaseReceiptDetails (ReceiptID, PurchaseDetailID, MaterialID, ReceivedQty, AcceptedQty, RejectedQty, UnitPrice)
SELECT @receiptId, DetailID, NULL, Quantity, Quantity, 0, UnitPrice
FROM PurchaseOrderDetails WHERE PurchaseID = 14;

-- Update PO Status to completed
UPDATE PurchaseOrders SET Status = '已完成' WHERE PurchaseID = 14;

-- Update RawMaterialInventory - add stock for purchased items
-- First, ensure items exist in inventory
IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-001')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt)
    VALUES ('BN-001', '薰衣草精华', 0, 2.5, 'ML', 'RAW', GETDATE());
IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-002')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt)
    VALUES ('BN-002', '玫瑰原精', 0, 3.5, 'ML', 'RAW', GETDATE());
IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-003')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt)
    VALUES ('BN-003', '檀香木油', 0, 8.0, 'ML', 'RAW', GETDATE());
IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-004')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt)
    VALUES ('BN-004', '白麝香', 0, 6.0, 'ML', 'RAW', GETDATE());
IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-005')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt)
    VALUES ('BN-005', '佛手柑精油', 0, 1.8, 'ML', 'RAW', GETDATE());

-- Update stock quantities
UPDATE RawMaterialInventory SET StockQty = StockQty + 1000, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-001';
UPDATE RawMaterialInventory SET StockQty = StockQty + 1000, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-002';
UPDATE RawMaterialInventory SET StockQty = StockQty + 1000, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-003';
UPDATE RawMaterialInventory SET StockQty = StockQty + 500, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-004';
UPDATE RawMaterialInventory SET StockQty = StockQty + 2000, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-005';

PRINT 'Inventory updated successfully';

-- Verify
SELECT ItemCode, ItemName, StockQty, UnitPrice FROM RawMaterialInventory WHERE ItemCode LIKE 'BN-%';
