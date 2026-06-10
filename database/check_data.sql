SELECT OrderID, OrderNo, TotalAmount, CostAmount, ProfitAmount, Status FROM Orders WHERE OrderNo = 'ORD-SIM-001';
SELECT NoteID, StockQuantity FROM NoteInventory WHERE NoteID BETWEEN 63 AND 69 ORDER BY NoteID;
SELECT ProductID, StockQty, UnitCost FROM ProductInventory WHERE ProductID IN (71,72,73);
SELECT ProductID, ProductName, UnitCost, BOMCost, BasePrice FROM Products WHERE ProductID IN (71,72,73);
SELECT TOP 1 AccountID, AccountName, AccountType, AvailableBalance, TotalBalance FROM FundAccounts;
SELECT COUNT(*) AS CampaignCount FROM MarketingCampaigns;
SELECT COUNT(*) AS CouponCount FROM Coupons;
