USE PerfumeShop;
SELECT '---PO---' AS Section;
SELECT PurchaseID, Status, TotalAmount FROM PurchaseOrders WHERE PurchaseID=14;
SELECT '---Receipt---' AS Section;
SELECT COUNT(*) AS ReceiptCount FROM PurchaseReceipts;
SELECT '---Inventory---' AS Section;
SELECT ItemCode, ItemName, StockQty, UnitPrice FROM RawMaterialInventory WHERE ItemCode LIKE 'BN-%';
