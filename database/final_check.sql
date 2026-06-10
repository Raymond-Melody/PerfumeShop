-- Check OrderDetailNoteSelections for custom/KOL products
SELECT '=== OrderDetailNoteSelections ===' AS Info;
SELECT ns.SelectionID, ns.DetailID, ns.NoteID, ns.NoteType, ns.Percentage
FROM OrderDetailNoteSelections ns
ORDER BY ns.DetailID, ns.NoteType;

-- Supply chain trace
SELECT '=== Purchase Orders ===' AS Info;
SELECT PurchaseID, PurchaseNo, SupplierID, TotalAmount, Status FROM PurchaseOrders;

SELECT '=== Purchase Receipts ===' AS Info;
SELECT ReceiptID, PurchaseID, ReceiptNo, Status FROM PurchaseReceipts;

SELECT '=== Material Outbound ===' AS Info;
SELECT OutboundID, OutboundNo, Status FROM MaterialOutbound;

-- Profit margin
SELECT '=== Profit Analysis ===' AS Info;
SELECT OrderNo, TotalAmount, CostAmount, ProfitAmount,
    CASE WHEN TotalAmount > 0 THEN ROUND(ProfitAmount / TotalAmount * 100, 2) ELSE 0 END AS ProfitMarginPct
FROM Orders WHERE OrderNo = 'ORD-SIM-001';
