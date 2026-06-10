SELECT '=== Products Cost ===' AS Info;
SELECT ProductID, ProductName, UnitCost, BOMCost, BasePrice FROM Products WHERE ProductID IN (71,72,73);

SELECT '=== ProductInventory Cost ===' AS Info;
SELECT ProductID, StockQty, UnitCost FROM ProductInventory WHERE ProductID IN (71,72,73);

SELECT '=== ProductCosts ===' AS Info;
SELECT CostID, ProductID, CostName, CostType, UnitCost, TotalCost, Quantity FROM ProductCosts WHERE ProductID IN (71,72,73);

SELECT '=== Order Cost ===' AS Info;
SELECT OrderID, OrderNo, TotalAmount, CostAmount, ProfitAmount FROM Orders WHERE OrderNo = 'ORD-SIM-001';

SELECT '=== MarketingCampaigns ===' AS Info;
SELECT CampaignID, CampaignName, CampaignType, DiscountValue, MinPurchase, IsActive FROM MarketingCampaigns;

SELECT '=== Coupons ===' AS Info;
SELECT CouponID, CouponCode, DiscountType, DiscountValue, MinPurchase, IsActive FROM Coupons;

SELECT '=== ExpenseRecords ===' AS Info;
SELECT ExpenseID, ExpenseName, ExpenseType, Amount, AllocationMethod, Period FROM ExpenseRecords;

SELECT '=== GLTransactions ===' AS Info;
SELECT GLID, GLNo, AccountCode, AccountName, DebitAmount, CreditAmount, Description FROM GLTransactions;

SELECT '=== FundAccounts ===' AS Info;
SELECT AccountID, AccountName, AccountType, AvailableBalance, TotalBalance FROM FundAccounts;

SELECT '=== DailyStatistics ===' AS Info;
SELECT StatID, StatDate, TotalOrders, TotalRevenue, NewUsers FROM DailyStatistics;

SELECT '=== ReconciliationLogs ===' AS Info;
SELECT LogID, OrderID, OrderNo, OrderAmount, PaymentAmount, Difference, Status FROM ReconciliationLogs;
