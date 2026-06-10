USE PerfumeShop;
GO
DECLARE @poId INT;
SELECT @poId = MAX(PurchaseID) FROM PurchaseOrders;

UPDATE PurchaseOrderDetails 
SET UnitPrice = CASE ItemCode 
    WHEN 'BN-001' THEN 2.5 
    WHEN 'BN-002' THEN 3.5 
    WHEN 'BN-003' THEN 8.0 
    WHEN 'BN-004' THEN 6.0 
    WHEN 'BN-005' THEN 1.8 
END,
TotalPrice = Quantity * CASE ItemCode 
    WHEN 'BN-001' THEN 2.5 
    WHEN 'BN-002' THEN 3.5 
    WHEN 'BN-003' THEN 8.0 
    WHEN 'BN-004' THEN 6.0 
    WHEN 'BN-005' THEN 1.8 
END
WHERE PurchaseID = @poId;

UPDATE PurchaseOrders 
SET TotalAmount = (SELECT SUM(TotalPrice) FROM PurchaseOrderDetails WHERE PurchaseID = @poId),
    Status = '已审批',
    ApprovedAt = GETDATE(),
    ApprovedBy = 2
WHERE PurchaseID = @poId;

SELECT 'PO Updated' AS Status, @poId AS PO_ID;
GO
