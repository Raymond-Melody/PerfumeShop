USE PerfumeShop;
SELECT 'PO' AS T, PurchaseID, Status FROM PurchaseOrders WHERE PurchaseID=14;
SELECT 'Rec' AS T, COUNT(*) AS C FROM PurchaseReceipts;
SELECT 'Inv' AS T, COUNT(*) AS C FROM RawMaterialInventory WHERE ItemCode IN ('BN-001','BN-002','BN-003','BN-004','BN-005');
SELECT * FROM RawMaterialInventory WHERE ItemCode IN ('BN-001','BN-002','BN-003','BN-004','BN-005');
