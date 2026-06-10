-- ============================================
-- Task 11: Clean Duplicates + Full Verification
-- ============================================

-- Step 1: Clean duplicates from first partial run
DELETE FROM ProductCosts WHERE CostID IN (4,5,6);
DELETE FROM MarketingCampaigns WHERE CampaignID = 3;
DELETE FROM Coupons WHERE CouponID = 3;
DELETE FROM ExpenseRecords WHERE ExpenseID = 17;
DELETE FROM GLTransactions WHERE GLID IN (3,4);
DELETE FROM DailyStatistics WHERE StatID = 2;
DELETE FROM ReconciliationLogs WHERE LogID = 2;

-- ============================================
-- V1: Inventory Consistency
-- ============================================
SELECT '=== V1: RawMaterial Inventory ===' AS Info;
SELECT MaterialID, ItemCode, ItemName, StockQty, UnitPrice
FROM RawMaterialInventory ORDER BY ItemCode;

SELECT '=== V1: Note Inventory ===' AS Info;
SELECT ni.NoteID, fn.NoteName, ni.StockQuantity
FROM NoteInventory ni
JOIN FragranceNotes fn ON ni.NoteID = fn.NoteID
WHERE ni.NoteID BETWEEN 63 AND 69
ORDER BY ni.NoteID;

SELECT '=== V1: Product Inventory ===' AS Info;
SELECT pi.ProductID, p.ProductName, pi.StockQty, pi.UnitCost
FROM ProductInventory pi
JOIN Products p ON pi.ProductID = p.ProductID
WHERE pi.ProductID IN (71,72,73)
ORDER BY pi.ProductID;

-- ============================================
-- V2: Financial Consistency
-- ============================================
SELECT '=== V2: Order Profit Check ===' AS Info;
SELECT OrderID, OrderNo, TotalAmount, CostAmount, ProfitAmount,
    CASE WHEN ProfitAmount = TotalAmount - CostAmount THEN 'OK' ELSE 'MISMATCH' END AS ProfitCheck
FROM Orders WHERE OrderNo = 'ORD-SIM-001';

-- ============================================
-- V3: Production Traceability
-- ============================================
SELECT '=== V3: Production Orders ===' AS Info;
SELECT ProductionID, OrderID, RecipeID, WorkOrderNo, PlannedQty, TotalBottles, Status
FROM ProductionOrders ORDER BY ProductionID;

SELECT '=== V3: Accord Productions ===' AS Info;
SELECT ProductionID, NoteID, BatchNo, PlannedQty, ActualQty, Status
FROM AccordProductions ORDER BY ProductionID;

SELECT '=== V3: Product Manufacturing ===' AS Info;
SELECT ManufacturingID, ProductID, BatchNo, PlannedQty, ActualQty, Status
FROM ProductManufacturing ORDER BY ManufacturingID;

-- ============================================
-- V4: Cost Chain Completeness
-- ============================================
SELECT '=== V4: Product BOM Cost ===' AS Info;
SELECT ProductID, ProductName, UnitCost, BOMCost, BasePrice
FROM Products WHERE ProductID IN (71,72,73);

SELECT '=== V4: ProductCosts Records ===' AS Info;
SELECT CostID, ProductID, CostName, CostType, UnitCost, TotalCost, Quantity
FROM ProductCosts WHERE ProductID IN (71,72,73);

-- ============================================
-- V5: Marketing, GL, Fund
-- ============================================
SELECT '=== V5: Marketing Campaign ===' AS Info;
SELECT CampaignID, CampaignName, CampaignType, DiscountValue, MinPurchase FROM MarketingCampaigns;
SELECT '=== V5: Coupons ===' AS Info;
SELECT CouponID, CouponCode, DiscountType, DiscountValue, MinPurchase FROM Coupons;
SELECT '=== V5: Expenses ===' AS Info;
SELECT ExpenseID, ExpenseName, Amount, AllocationMethod FROM ExpenseRecords;

SELECT '=== V5: GL Transactions ===' AS Info;
SELECT GLID, GLNo, AccountCode, AccountName, DebitAmount, CreditAmount, Description FROM GLTransactions;

SELECT '=== V5: Fund Accounts ===' AS Info;
SELECT AccountID, AccountName, AvailableBalance, TotalBalance FROM FundAccounts;

-- ============================================
-- V6: Payment & Reconciliation
-- ============================================
SELECT '=== V6: Payment Records ===' AS Info;
SELECT RecordID, OrderID, OrderNo, Amount, PaymentMethod, Status FROM PaymentRecords;

SELECT '=== V6: Reconciliation ===' AS Info;
SELECT LogID, OrderID, OrderNo, OrderAmount, PaymentAmount, Difference, Status FROM ReconciliationLogs;

-- ============================================
-- V7: Order Details (3 product lines)
-- ============================================
SELECT '=== V7: Order Details ===' AS Info;
SELECT DetailID, OrderID, ProductID, ProductName, Quantity, UnitPrice, Subtotal, VolumeML, TopNoteName, MiddleNoteName, BaseNoteName
FROM OrderDetails
WHERE OrderID = (SELECT MAX(OrderID) FROM Orders WHERE OrderNo = 'ORD-SIM-001')
ORDER BY DetailID;

-- ============================================
-- V8: User check
-- ============================================
SELECT '=== V8: Users ===' AS Info;
SELECT UserID, Username, Email, UserRole, IsActive FROM Users;

SELECT '=== FINAL: ALL VERIFICATIONS DONE ===' AS FinalResult;
