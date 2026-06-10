UPDATE PurchaseOrders SET Status = '已审批' WHERE PurchaseID = 14;
SELECT @@ROWCOUNT AS Updated;
