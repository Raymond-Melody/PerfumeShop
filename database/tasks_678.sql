-- ============================================
-- Tasks 6-8: Semi-finished Production + Product Manufacturing + User Order
-- ============================================

-- Task 6: Accord Productions (semi-finished / fragrance note production)
-- Produce 7 fragrance notes using raw materials

-- Material outbound records
DECLARE @outId1 INT, @outId2 INT, @outId3 INT, @outId4 INT, @outId5 INT, @outId6 INT, @outId7 INT;

-- Outbound for Top Notes
INSERT INTO MaterialOutbound (OutboundNo, OutboundType, ReferenceType, Status, OutboundDate, RequestedBy, Notes, CreatedAt)
VALUES ('MO-001', 'Production', 'AccordProduction', 'Completed', GETDATE(), 'Ray88', 'Produce Top Notes', GETDATE());
SET @outId1 = SCOPE_IDENTITY();

INSERT INTO MaterialOutboundDetails (OutboundID, MaterialID, RequestedQty, ActualQty, UnitPrice, TotalAmount, ProductionOrderRef)
SELECT @outId1, MaterialID, 200, 200, UnitPrice, 200*UnitPrice, NULL 
FROM RawMaterialInventory WHERE ItemCode IN ('BN-001','BN-005');

-- Consume raw materials
UPDATE RawMaterialInventory SET StockQty = StockQty - 200, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-001';
UPDATE RawMaterialInventory SET StockQty = StockQty - 200, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-005';

-- Outbound for Middle Notes
INSERT INTO MaterialOutbound (OutboundNo, OutboundType, ReferenceType, Status, OutboundDate, RequestedBy, Notes, CreatedAt)
VALUES ('MO-002', 'Production', 'AccordProduction', 'Completed', GETDATE(), 'Ray88', 'Produce Middle Notes', GETDATE());
SET @outId2 = SCOPE_IDENTITY();

INSERT INTO MaterialOutboundDetails (OutboundID, MaterialID, RequestedQty, ActualQty, UnitPrice, TotalAmount, ProductionOrderRef)
SELECT @outId2, MaterialID, 300, 300, UnitPrice, 300*UnitPrice, NULL 
FROM RawMaterialInventory WHERE ItemCode IN ('BN-002','BN-005');

UPDATE RawMaterialInventory SET StockQty = StockQty - 300, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-002';
UPDATE RawMaterialInventory SET StockQty = StockQty - 300, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-005';

-- Outbound for Base Notes
INSERT INTO MaterialOutbound (OutboundNo, OutboundType, ReferenceType, Status, OutboundDate, RequestedBy, Notes, CreatedAt)
VALUES ('MO-003', 'Production', 'AccordProduction', 'Completed', GETDATE(), 'Ray88', 'Produce Base Notes', GETDATE());
SET @outId3 = SCOPE_IDENTITY();

INSERT INTO MaterialOutboundDetails (OutboundID, MaterialID, RequestedQty, ActualQty, UnitPrice, TotalAmount, ProductionOrderRef)
SELECT @outId3, MaterialID, 400, 400, UnitPrice, 400*UnitPrice, NULL 
FROM RawMaterialInventory WHERE ItemCode IN ('BN-003','BN-004');

UPDATE RawMaterialInventory SET StockQty = StockQty - 400, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-003';
UPDATE RawMaterialInventory SET StockQty = StockQty - 400, UpdatedAt = GETDATE() WHERE ItemCode = 'BN-004';

-- Create AccordProduction records
DECLARE @apId1 INT, @apId2 INT, @apId3 INT, @apId4 INT, @apId5 INT, @apId6 INT, @apId7 INT;

-- Top notes production
INSERT INTO AccordProductions (NoteID, NoteName, AccordRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (63, 'Citrus Fresh', 1, 'BATCH-T1', 1000, 1000, 'Completed', DATEADD(hour,-4,GETDATE()), DATEADD(hour,-2,GETDATE()), GETDATE(), 'WS-1');
SET @apId1 = SCOPE_IDENTITY();

INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty, UnitCost, TotalCost)
SELECT @apId1, MaterialID, ItemName, 200, 200, UnitPrice, 200*UnitPrice FROM RawMaterialInventory WHERE ItemCode IN ('BN-001','BN-005');

INSERT INTO AccordProductions (NoteID, NoteName, AccordRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (64, 'Green Leaf', 2, 'BATCH-T2', 1000, 1000, 'Completed', DATEADD(hour,-4,GETDATE()), DATEADD(hour,-2,GETDATE()), GETDATE(), 'WS-1');
SET @apId2 = SCOPE_IDENTITY();

INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty, UnitCost, TotalCost)
SELECT @apId2, MaterialID, ItemName, 150, 150, UnitPrice, 150*UnitPrice FROM RawMaterialInventory WHERE ItemCode IN ('BN-001','BN-005');

-- Middle notes production
INSERT INTO AccordProductions (NoteID, NoteName, AccordRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (65, 'Rose Floral', 3, 'BATCH-M1', 1500, 1500, 'Completed', DATEADD(hour,-4,GETDATE()), DATEADD(hour,-2,GETDATE()), GETDATE(), 'WS-2');
SET @apId3 = SCOPE_IDENTITY();

INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty, UnitCost, TotalCost)
SELECT @apId3, MaterialID, ItemName, 300, 300, UnitPrice, 300*UnitPrice FROM RawMaterialInventory WHERE ItemCode IN ('BN-002','BN-005');

INSERT INTO AccordProductions (NoteID, NoteName, AccordRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (66, 'Woody Warm', 4, 'BATCH-M2', 1000, 1000, 'Completed', DATEADD(hour,-4,GETDATE()), DATEADD(hour,-2,GETDATE()), GETDATE(), 'WS-2');
SET @apId4 = SCOPE_IDENTITY();

INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty, UnitCost, TotalCost)
SELECT @apId4, MaterialID, ItemName, 200, 200, UnitPrice, 200*UnitPrice FROM RawMaterialInventory WHERE ItemCode IN ('BN-003');

INSERT INTO AccordProductions (NoteID, NoteName, AccordRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (67, 'Fruity Sweet', 5, 'BATCH-M3', 800, 800, 'Completed', DATEADD(hour,-4,GETDATE()), DATEADD(hour,-2,GETDATE()), GETDATE(), 'WS-2');
SET @apId5 = SCOPE_IDENTITY();

INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty, UnitCost, TotalCost)
SELECT @apId5, MaterialID, ItemName, 150, 150, UnitPrice, 150*UnitPrice FROM RawMaterialInventory WHERE ItemCode IN ('BN-002','BN-005');

-- Base notes production
INSERT INTO AccordProductions (NoteID, NoteName, AccordRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (68, 'Oriental Amber', 6, 'BATCH-B1', 1200, 1200, 'Completed', DATEADD(hour,-4,GETDATE()), DATEADD(hour,-2,GETDATE()), GETDATE(), 'WS-3');
SET @apId6 = SCOPE_IDENTITY();

INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty, UnitCost, TotalCost)
SELECT @apId6, MaterialID, ItemName, 250, 250, UnitPrice, 250*UnitPrice FROM RawMaterialInventory WHERE ItemCode IN ('BN-003');

INSERT INTO AccordProductions (NoteID, NoteName, AccordRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (69, 'Woody Musk', 7, 'BATCH-B2', 1000, 1000, 'Completed', DATEADD(hour,-4,GETDATE()), DATEADD(hour,-2,GETDATE()), GETDATE(), 'WS-3');
SET @apId7 = SCOPE_IDENTITY();

INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty, UnitCost, TotalCost)
SELECT @apId7, MaterialID, ItemName, 400, 400, UnitPrice, 400*UnitPrice FROM RawMaterialInventory WHERE ItemCode IN ('BN-003','BN-004');

-- QC Reports for all accord productions
INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, Notes, TestDate, CreatedAt)
VALUES (@apId1, 'BATCH-T1', 'Pass', 'QC Pass', GETDATE(), GETDATE());
INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, Notes, TestDate, CreatedAt)
VALUES (@apId2, 'BATCH-T2', 'Pass', 'QC Pass', GETDATE(), GETDATE());
INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, Notes, TestDate, CreatedAt)
VALUES (@apId3, 'BATCH-M1', 'Pass', 'QC Pass', GETDATE(), GETDATE());
INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, Notes, TestDate, CreatedAt)
VALUES (@apId4, 'BATCH-M2', 'Pass', 'QC Pass', GETDATE(), GETDATE());
INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, Notes, TestDate, CreatedAt)
VALUES (@apId5, 'BATCH-M3', 'Pass', 'QC Pass', GETDATE(), GETDATE());
INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, Notes, TestDate, CreatedAt)
VALUES (@apId6, 'BATCH-B1', 'Pass', 'QC Pass', GETDATE(), GETDATE());
INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, Notes, TestDate, CreatedAt)
VALUES (@apId7, 'BATCH-B2', 'Pass', 'QC Pass', GETDATE(), GETDATE());

-- Update NoteInventory (add produced stock)
UPDATE NoteInventory SET StockQuantity = StockQuantity + 1000, UpdatedAt = GETDATE() WHERE NoteID = 63;
UPDATE NoteInventory SET StockQuantity = StockQuantity + 1000, UpdatedAt = GETDATE() WHERE NoteID = 64;
UPDATE NoteInventory SET StockQuantity = StockQuantity + 1500, UpdatedAt = GETDATE() WHERE NoteID = 65;
UPDATE NoteInventory SET StockQuantity = StockQuantity + 1000, UpdatedAt = GETDATE() WHERE NoteID = 66;
UPDATE NoteInventory SET StockQuantity = StockQuantity + 800, UpdatedAt = GETDATE() WHERE NoteID = 67;
UPDATE NoteInventory SET StockQuantity = StockQuantity + 1200, UpdatedAt = GETDATE() WHERE NoteID = 68;
UPDATE NoteInventory SET StockQuantity = StockQuantity + 1000, UpdatedAt = GETDATE() WHERE NoteID = 69;

-- ============================================
-- Task 7: Product Manufacturing
-- ============================================

-- Create product manufacturing records for all 3 products
DECLARE @pmId1 INT, @pmId2 INT, @pmId3 INT;

INSERT INTO ProductManufacturing (ProductID, ProductName, ProductRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (71, 'ChenXi', NULL, 'BATCH-P1', 100, 100, 'Completed', DATEADD(hour,-2,GETDATE()), DATEADD(hour,-1,GETDATE()), GETDATE(), 'WS-MAIN');
SET @pmId1 = SCOPE_IDENTITY();

INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId1, 63, 'Citrus Fresh', 30, 30);
INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId1, 65, 'Rose Floral', 40, 40);
INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId1, 69, 'Woody Musk', 30, 30);

INSERT INTO ProductManufacturing (ProductID, ProductName, ProductRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (72, 'SiXiang', NULL, 'BATCH-P2', 50, 50, 'Completed', DATEADD(hour,-2,GETDATE()), DATEADD(hour,-1,GETDATE()), GETDATE(), 'WS-MAIN');
SET @pmId2 = SCOPE_IDENTITY();

INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId2, 64, 'Green Leaf', 30, 30);
INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId2, 66, 'Woody Warm', 40, 40);
INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId2, 68, 'Oriental Amber', 30, 30);

INSERT INTO ProductManufacturing (ProductID, ProductName, ProductRecipeID, BatchNo, PlannedQty, ActualQty, Status, StartedAt, CompletedAt, CreatedAt, WorkCenter)
VALUES (73, 'YueYe', NULL, 'BATCH-P3', 80, 80, 'Completed', DATEADD(hour,-2,GETDATE()), DATEADD(hour,-1,GETDATE()), GETDATE(), 'WS-MAIN');
SET @pmId3 = SCOPE_IDENTITY();

INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId3, 63, 'Citrus Fresh', 25, 25);
INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId3, 65, 'Rose Floral', 45, 45);
INSERT INTO ProductManufacturingDetails (ManufacturingID, NoteID, NoteName, PlannedQty, ActualQty)
VALUES (@pmId3, 68, 'Oriental Amber', 30, 30);

-- Consume note inventory
UPDATE NoteInventory SET StockQuantity = StockQuantity - 30, UpdatedAt = GETDATE() WHERE NoteID = 63;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 40, UpdatedAt = GETDATE() WHERE NoteID = 65;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 30, UpdatedAt = GETDATE() WHERE NoteID = 69;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 30, UpdatedAt = GETDATE() WHERE NoteID = 64;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 40, UpdatedAt = GETDATE() WHERE NoteID = 66;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 30, UpdatedAt = GETDATE() WHERE NoteID = 68;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 25, UpdatedAt = GETDATE() WHERE NoteID = 63;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 45, UpdatedAt = GETDATE() WHERE NoteID = 65;
UPDATE NoteInventory SET StockQuantity = StockQuantity - 30, UpdatedAt = GETDATE() WHERE NoteID = 68;

-- Update ProductInventory
UPDATE ProductInventory SET StockQty = StockQty + 100, UpdatedAt = GETDATE() WHERE ProductID = 71;
UPDATE ProductInventory SET StockQty = StockQty + 50, UpdatedAt = GETDATE() WHERE ProductID = 72;
UPDATE ProductInventory SET StockQty = StockQty + 80, UpdatedAt = GETDATE() WHERE ProductID = 73;

-- ============================================
-- Task 8: User Order (Raymond buys 3 products)
-- ============================================

DECLARE @userId INT, @orderId INT;
SELECT @userId = UserID FROM Users WHERE Username = 'Raymond';

-- Create Order
INSERT INTO Orders (OrderNo, UserID, Status, TotalAmount, ShippingFee, PaymentMethod, ShippingAddress, ShippingName, ShippingPhone, ShippingStatus, CreatedAt, UpdatedAt, ChannelSource)
VALUES ('ORD-SIM-001', @userId, 'Paid', 847.00, 0, 'Simulated', 'Guangdong Shenzhen Nanshan Keji Rd 100', 'Raymond', '13800138000', 'Pending', GETDATE(), GETDATE(), 'Direct');

SET @orderId = SCOPE_IDENTITY();

-- Order Detail 1: Fixed Brand (晨曦之光 #71)
INSERT INTO OrderDetails (OrderID, ProductID, ProductName, Quantity, UnitPrice, Subtotal, VolumeML, VolumeName)
VALUES (@orderId, 71, 'ChenXi', 1, 99.00, 99.00, 50, '50ML');

-- Order Detail 2: Custom (私享时光 #72) with note selections
DECLARE @detailId2 INT;
INSERT INTO OrderDetails (OrderID, ProductID, ProductName, Quantity, UnitPrice, Subtotal, VolumeML, VolumeName, CustomLabel)
VALUES (@orderId, 72, 'SiXiang', 1, 399.00, 399.00, 50, '50ML', 'For My Love');
SET @detailId2 = SCOPE_IDENTITY();

INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage)
VALUES (@detailId2, 64, 'Top', 30);
INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage)
VALUES (@detailId2, 66, 'Middle', 40);
INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage)
VALUES (@detailId2, 68, 'Base', 30);

-- Order Detail 3: KOL (月夜玫瑰 #73) with preset notes
DECLARE @detailId3 INT;
INSERT INTO OrderDetails (OrderID, ProductID, ProductName, Quantity, UnitPrice, Subtotal, VolumeML, VolumeName)
VALUES (@orderId, 73, 'YueYe', 1, 349.00, 349.00, 50, '50ML');
SET @detailId3 = SCOPE_IDENTITY();

INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage)
VALUES (@detailId3, 63, 'Top', 25);
INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage)
VALUES (@detailId3, 65, 'Middle', 45);
INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage)
VALUES (@detailId3, 68, 'Base', 30);

-- Payment Record
INSERT INTO PaymentRecords (OrderID, OrderNo, Amount, PaymentMethod, TransactionNo, Status, TransactionType, CreatedAt)
VALUES (@orderId, 'ORD-SIM-001', 847.00, 'Simulated', 'TXN-SIM-001', 'Completed', 'Payment', GETDATE());

-- Create Production Orders for the order
DECLARE @poId1 INT, @poId2 INT, @poId3 INT;

INSERT INTO ProductionOrders (OrderID, DetailID, Status, Priority, Notes, CreatedAt, EstimatedDate, TotalBottles, WorkOrderNo)
VALUES (@orderId, (SELECT MIN(DetailID) FROM OrderDetails WHERE OrderID = @orderId), 'Completed', 1, 'Fixed Brand Production', GETDATE(), GETDATE(), 1, 'WO-SIM-001');
SET @poId1 = SCOPE_IDENTITY();

INSERT INTO ProductionOrders (OrderID, DetailID, Status, Priority, Notes, CreatedAt, EstimatedDate, TotalBottles, WorkOrderNo)
VALUES (@orderId, @detailId2, 'Completed', 1, 'Custom Production', GETDATE(), GETDATE(), 1, 'WO-SIM-002');
SET @poId2 = SCOPE_IDENTITY();

INSERT INTO ProductionOrders (OrderID, DetailID, Status, Priority, Notes, CreatedAt, EstimatedDate, TotalBottles, WorkOrderNo)
VALUES (@orderId, @detailId3, 'Completed', 1, 'KOL Production', GETDATE(), GETDATE(), 1, 'WO-SIM-003');
SET @poId3 = SCOPE_IDENTITY();

-- Create ProductionLogs
INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedAt, CreatedBy)
VALUES (@poId1, 'Completed', 'Production completed', GETDATE(), 'Ray88');
INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedAt, CreatedBy)
VALUES (@poId2, 'Completed', 'Production completed', GETDATE(), 'Ray88');
INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedAt, CreatedBy)
VALUES (@poId3, 'Completed', 'Production completed', GETDATE(), 'Ray88');

-- Update inventory (consume for order fulfillment)
UPDATE ProductInventory SET StockQty = StockQty - 1, UpdatedAt = GETDATE() WHERE ProductID = 71;
UPDATE ProductInventory SET StockQty = StockQty - 1, UpdatedAt = GETDATE() WHERE ProductID = 72;
UPDATE ProductInventory SET StockQty = StockQty - 1, UpdatedAt = GETDATE() WHERE ProductID = 73;

-- Update order to completed/shipped
UPDATE Orders SET Status = 'Completed', ShippingStatus = 'Delivered', ShippedAt = GETDATE(), DeliveredAt = GETDATE(), TrackingNumber = 'SF-SIM-001', ShippingCompany = 'SF Express' WHERE OrderID = @orderId;
UPDATE ProductionOrders SET Status = 'Completed', CompletedAt = GETDATE(), WarehouseInAt = GETDATE(), ShippedOutAt = GETDATE() WHERE OrderID = @orderId;

SELECT 'TASKS 6-8 DONE' AS Result;
