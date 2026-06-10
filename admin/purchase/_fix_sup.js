const fs = require('fs');
const file = 'f:/网站制作/网站/网站二/admin/purchase/supplier_management.asp';
let content = fs.readFileSync(file, 'utf8');

// Fix query to include PurchaseID
const oldQuery = 'SELECT TOP 5 PurchaseNo, OrderDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status FROM PurchaseOrders WHERE SupplierID=';
const newQuery = 'SELECT TOP 5 PurchaseID, PurchaseNo, OrderDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status FROM PurchaseOrders WHERE SupplierID=';
content = content.replace(oldQuery, newQuery);

// Fix the link to use PurchaseID
const oldLink1 = 'purchase_orders.asp?view=<%= rsSupplierOrders("PurchaseNo") %>';
const newLink1 = 'purchase_orders.asp?view=<%= rsSupplierOrders("PurchaseID") %>';
content = content.replace(oldLink1, newLink1);

// Also fix the link text that was showing PurchaseNo
const oldLinkText = '<%= Server.HTMLEncode(rsSupplierOrders("PurchaseNo") & "") %>';
// That's fine, keep it

fs.writeFileSync(file, content, 'utf8');
console.log('Fix applied');
