USE PerfumeShop;
GO

-- Fix PO status
UPDATE PurchaseOrders SET Status = '已审批' WHERE PurchaseID = 14;
GO

-- Create Purchase Receipt
INSERT INTO PurchaseReceipts (PurchaseID, SupplierID, ReceiptNo, ReceiptDate, Status, ReceivedBy, TotalReceivedQty, Notes, CreatedAt)
VALUES (14, 1, 'REC-001', GETDATE(), '已收货', 'Ray88', 5500, '模拟收货', GETDATE());
GO

-- Receipt details
INSERT INTO PurchaseReceiptDetails (ReceiptID, PurchaseDetailID, ReceivedQty, AcceptedQty, RejectedQty, UnitPrice)
SELECT (SELECT MAX(ReceiptID) FROM PurchaseReceipts), DetailID, Quantity, Quantity, 0, UnitPrice
FROM PurchaseOrderDetails WHERE PurchaseID = 14;
GO

-- Update PO
UPDATE PurchaseOrders SET Status = '已完成' WHERE PurchaseID = 14;
GO

-- Add inventory items
IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-001')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt) VALUES ('BN-001', 'XunYiCao', 1000, 2.5, 'ML', 'RAW', GETDATE());
ELSE UPDATE RawMaterialInventory SET StockQty = StockQty + 1000 WHERE ItemCode = 'BN-001';
GO

IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-002')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt) VALUES ('BN-002', 'MeiGui', 1000, 3.5, 'ML', 'RAW', GETDATE());
ELSE UPDATE RawMaterialInventory SET StockQty = StockQty + 1000 WHERE ItemCode = 'BN-002';
GO

IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-003')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt) VALUES ('BN-003', 'TanXiang', 1000, 8.0, 'ML', 'RAW', GETDATE());
ELSE UPDATE RawMaterialInventory SET StockQty = StockQty + 1000 WHERE ItemCode = 'BN-003';
GO

IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-004')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt) VALUES ('BN-004', 'BaiShe', 500, 6.0, 'ML', 'RAW', GETDATE());
ELSE UPDATE RawMaterialInventory SET StockQty = StockQty + 500 WHERE ItemCode = 'BN-004';
GO

IF NOT EXISTS (SELECT 1 FROM RawMaterialInventory WHERE ItemCode = 'BN-005')
    INSERT INTO RawMaterialInventory (ItemCode, ItemName, StockQty, UnitPrice, Unit, CategoryCode, UpdatedAt) VALUES ('BN-005', 'FoShouGan', 2000, 1.8, 'ML', 'RAW', GETDATE());
ELSE UPDATE RawMaterialInventory SET StockQty = StockQty + 2000 WHERE ItemCode = 'BN-005';
GO

SELECT 'ALL DONE' AS Result;
GO
