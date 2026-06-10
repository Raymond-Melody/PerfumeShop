-- ============================================
-- Task 9: Cost Chain Setup
-- ============================================

-- Update ProductInventory with BOM cost
-- Product 71 (Fixed Brand): Citrus 30% + Rose 40% + WoodyMusk 30%
-- Cost per ml: 0.3*2.15 + 0.4*2.65 + 0.3*7.0 = 0.645 + 1.06 + 2.1 = 3.805
UPDATE ProductInventory SET UnitCost = 3.805 WHERE ProductID = 71;
UPDATE Products SET UnitCost = 3.805, BOMCost = 3.805 WHERE ProductID = 71;

-- Product 72 (Custom): Green 30% + Woody 40% + Amber 30%
-- Cost per ml: 0.3*2.15 + 0.4*8.0 + 0.3*8.0 = 0.645 + 3.2 + 2.4 = 6.245
UPDATE ProductInventory SET UnitCost = 6.245 WHERE ProductID = 72;
UPDATE Products SET UnitCost = 6.245, BOMCost = 6.245 WHERE ProductID = 72;

-- Product 73 (KOL): Citrus 25% + Rose 45% + Amber 30%
-- Cost per ml: 0.25*2.15 + 0.45*2.65 + 0.3*8.0 = 0.5375 + 1.1925 + 2.4 = 4.13
UPDATE ProductInventory SET UnitCost = 4.13 WHERE ProductID = 73;
UPDATE Products SET UnitCost = 4.13, BOMCost = 4.13 WHERE ProductID = 73;

-- ProductCosts table
INSERT INTO ProductCosts (ProductID, CostName, CostType, UnitCost, TotalCost, Quantity, EffectiveDate, CreatedAt, CreatedBy)
VALUES (71, 'ChenXi BOM Cost', 'BOM', 3.805, 190.25, 50, GETDATE(), GETDATE(), 'Ray88');
INSERT INTO ProductCosts (ProductID, CostName, CostType, UnitCost, TotalCost, Quantity, EffectiveDate, CreatedAt, CreatedBy)
VALUES (72, 'SiXiang BOM Cost', 'BOM', 6.245, 312.25, 50, GETDATE(), GETDATE(), 'Ray88');
INSERT INTO ProductCosts (ProductID, CostName, CostType, UnitCost, TotalCost, Quantity, EffectiveDate, CreatedAt, CreatedBy)
VALUES (73, 'YueYe BOM Cost', 'BOM', 4.13, 206.50, 50, GETDATE(), GETDATE(), 'Ray88');

-- Update Order costs
-- Cost = product unit cost * quantity (50ml each, 1 each)
-- Total cost: 3.805*50 + 6.245*50 + 4.13*50 = 190.25 + 312.25 + 206.50 = 709.00
-- Profit: 847.00 - 709.00 = 138.00
DECLARE @orderId INT;
SELECT @orderId = MAX(OrderID) FROM Orders;

UPDATE Orders SET 
    CostAmount = 709.00,
    ProfitAmount = 138.00
WHERE OrderID = @orderId;

-- ============================================
-- Task 10: Marketing & Expense Amortization
-- ============================================

-- Marketing Campaign
INSERT INTO MarketingCampaigns (CampaignName, CampaignType, Description, DiscountValue, MinPurchase, StartDate, EndDate, IsActive, CreatedAt)
VALUES ('Summer Limited', 'Discount', 'Summer special promotion - 10% off', 10, 100, GETDATE(), DATEADD(month,3,GETDATE()), 1, GETDATE());

-- Coupon
INSERT INTO Coupons (CouponCode, DiscountType, DiscountValue, MinPurchase, StartDate, EndDate, UsageLimit, UsedCount, IsActive, CreatedAt)
VALUES ('SUMMER2026', 'Percentage', 10, 200, GETDATE(), DATEADD(month,3,GETDATE()), 100, 0, 1, GETDATE());

-- Marketing Expense
DECLARE @expId INT;
INSERT INTO ExpenseRecords (ExpenseName, ExpenseType, Amount, AllocationMethod, AllocationRatio, Period, OrderID, SourceOrderID, CreatedAt)
VALUES ('Summer Ad Campaign', 'Marketing', 5000.00, 'PerOrder', 100, '2026-05', @orderId, @orderId, GETDATE());

-- Ensure FundAccounts exists and update balance
DECLARE @fundId INT;
IF NOT EXISTS (SELECT 1 FROM FundAccounts)
BEGIN
    INSERT INTO FundAccounts (AccountName, AccountType, AvailableBalance, TotalBalance, IsActive, CreatedAt, UpdatedAt)
    VALUES ('Main Fund', 'Cash', 847.00 - 5000.00, 847.00 - 5000.00, 1, GETDATE(), GETDATE());
    SET @fundId = SCOPE_IDENTITY();
END
ELSE
BEGIN
    SELECT TOP 1 @fundId = AccountID FROM FundAccounts;
    UPDATE FundAccounts SET 
        AvailableBalance = ISNULL(AvailableBalance,0) + 847.00 - 5000.00, 
        TotalBalance = ISNULL(TotalBalance,0) + 847.00 - 5000.00, 
        UpdatedAt = GETDATE() 
    WHERE AccountID = @fundId;
END

-- GL Transactions (using correct schema)
-- Expense: Debit 5000 (money going out for marketing)
INSERT INTO GLTransactions (GLNo, TransactionDate, AccountCode, AccountName, DebitAmount, CreditAmount, Balance, RefType, RefID, RefNo, Description, CreatedBy, CreatedAt)
VALUES ('GL-2026-001', GETDATE(), '5201', 'Marketing Expense', 5000.00, 0, -5000.00, 'Expense', @orderId, 'ORD-SIM-001', 'Marketing expense - Summer Ad Campaign', 'Ray88', GETDATE());

-- Revenue: Credit 847 (money coming in from order)
INSERT INTO GLTransactions (GLNo, TransactionDate, AccountCode, AccountName, DebitAmount, CreditAmount, Balance, RefType, RefID, RefNo, Description, CreatedBy, CreatedAt)
VALUES ('GL-2026-002', GETDATE(), '6001', 'Sales Revenue', 0, 847.00, 847.00, 'Order', @orderId, 'ORD-SIM-001', 'Order ORD-SIM-001 payment', 'Ray88', GETDATE());

-- Daily Statistics
INSERT INTO DailyStatistics (StatDate, TotalOrders, TotalRevenue, NewUsers, TotalUsers, CreatedAt, UpdatedAt)
VALUES (CAST(GETDATE() AS DATE), 1, 847.00, 1, 1, GETDATE(), GETDATE());

-- Reconciliation Log
INSERT INTO ReconciliationLogs (OrderID, OrderNo, OrderAmount, PaymentAmount, Difference, Status, ReconcileDate, CreatedAt)
VALUES (@orderId, 'ORD-SIM-001', 847.00, 847.00, 0, 'Matched', GETDATE(), GETDATE());

SELECT 'TASKS 9-10 DONE' AS Result;
