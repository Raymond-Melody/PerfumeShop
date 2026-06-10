USE PerfumeShop;
GO
UPDATE PurchaseOrderDetails SET UnitPrice=2.5, TotalPrice=Quantity*2.5 WHERE ItemCode='BN-001';
UPDATE PurchaseOrderDetails SET UnitPrice=3.5, TotalPrice=Quantity*3.5 WHERE ItemCode='BN-002';
UPDATE PurchaseOrderDetails SET UnitPrice=8.0, TotalPrice=Quantity*8.0 WHERE ItemCode='BN-003';
UPDATE PurchaseOrderDetails SET UnitPrice=6.0, TotalPrice=Quantity*6.0 WHERE ItemCode='BN-004';
UPDATE PurchaseOrderDetails SET UnitPrice=1.8, TotalPrice=Quantity*1.8 WHERE ItemCode='BN-005';
GO
UPDATE PurchaseOrders SET TotalAmount=(SELECT SUM(TotalPrice) FROM PurchaseOrderDetails WHERE PurchaseID=14), Status='已审批', ApprovedAt=GETDATE(), ApprovedBy=2 WHERE PurchaseID=14;
GO
SELECT 'OK' AS Result;
GO
