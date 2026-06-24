<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' V10: 确保批次追溯所需表存在
If Not TableExists_BD("PurchaseReceipts") Then Call EnsureTable_BD("PurchaseReceipts", _
    "ReceiptID INT IDENTITY(1,1) PRIMARY KEY, PurchaseID INT, ReceiptNo NVARCHAR(100), SupplierID INT, " & _
    "ReceivedBy NVARCHAR(100), ReceiptDate DATETIME, Status NVARCHAR(30), TotalReceivedQty FLOAT DEFAULT 0, " & _
    "Notes NVARCHAR(500), CreatedAt DATETIME DEFAULT GETDATE()")
If Not TableExists_BD("PurchaseBatches") Then Call EnsureTable_BD("PurchaseBatches", _
    "BatchID INT IDENTITY(1,1) PRIMARY KEY, PurchaseDetailID INT, PurchaseID INT, BatchNo NVARCHAR(100), " & _
    "ItemType NVARCHAR(30), ItemCode NVARCHAR(50), ItemName NVARCHAR(200), UnitPrice DECIMAL(19,4) DEFAULT 0, " & _
    "Quantity FLOAT DEFAULT 0, ReceivedQty FLOAT DEFAULT 0, RemainingQty FLOAT DEFAULT 0, " & _
    "ReceivedDate DATETIME, SupplierID INT, CreatedAt DATETIME DEFAULT GETDATE()")
If Not TableExists_BD("OrderCostAllocation") Then Call EnsureTable_BD("OrderCostAllocation", _
    "AllocationID INT IDENTITY(1,1) PRIMARY KEY, OrderID INT, OrderNo NVARCHAR(100), CostType NVARCHAR(30), " & _
    "ItemCode NVARCHAR(50), ItemName NVARCHAR(200), UnitCost DECIMAL(19,4) DEFAULT 0, Quantity FLOAT DEFAULT 0, " & _
    "TotalCost DECIMAL(19,4) DEFAULT 0, BatchID INT, AllocatedAt DATETIME DEFAULT GETDATE(), CreatedAt DATETIME DEFAULT GETDATE()")

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Dim receiptID
receiptID = SafeNum(Request.QueryString("receipt_id"))
If receiptID <= 0 Then
    Response.Write "<div style='text-align:center;color:#e74c3c;padding:20px;'>无效的收货单ID</div>"
    Response.End
End If

' 查询该收货单的采购批次
Dim sqlBatches, rsBatches
sqlBatches = "SELECT pb.BatchID, pb.BatchNo, pb.ItemType, pb.ItemCode, pb.ItemName, " & _
    "pb.UnitPrice, pb.Quantity, pb.ReceivedQty, pb.RemainingQty, pb.ReceivedDate, " & _
    "s.SupplierName, " & _
    "(SELECT COUNT(*) FROM OrderCostAllocation WHERE BatchID=pb.BatchID) AS AllocatedOrders " & _
    "FROM PurchaseBatches pb " & _
    "LEFT JOIN Suppliers s ON pb.SupplierID=s.SupplierID " & _
    "WHERE pb.PurchaseID=(SELECT PurchaseID FROM PurchaseReceipts WHERE ReceiptID=" & receiptID & ") " & _
    "ORDER BY pb.BatchNo"

Set rsBatches = conn.Execute(sqlBatches)

Function GetTypeColor(ct)
    Select Case ct
        Case "RawMaterial" : GetTypeColor = "#FF9800"
        Case "Packaging"   : GetTypeColor = "#2196F3"
        Case "Bottle"      : GetTypeColor = "#9C27B0"
        Case "Printing"    : GetTypeColor = "#00BCD4"
        Case "SprayHead"   : GetTypeColor = "#E91E63"
        Case Else          : GetTypeColor = "#888"
    End Select
End Function

Function GetTypeLabel(ct)
    Select Case ct
        Case "RawMaterial" : GetTypeLabel = "原料"
        Case "Packaging"   : GetTypeLabel = "包装"
        Case "Bottle"      : GetTypeLabel = "瓶子"
        Case "Printing"    : GetTypeLabel = "印刷品"
        Case "SprayHead"   : GetTypeLabel = "喷头"
        Case Else          : GetTypeLabel = ct
    End Select
End Function
%>
<div style="max-height:400px; overflow-y:auto;">
<%
If Not rsBatches Is Nothing Then
    Dim batchCount : batchCount = 0
    Do While Not rsBatches.EOF
        batchCount = batchCount + 1
        Dim batchNo, itemType, itemCode, itemName, unitPrice, qty, receivedQty, remainingQty, allocOrders
        batchNo = rsBatches("BatchNo") & ""
        itemType = rsBatches("ItemType") & ""
        itemCode = rsBatches("ItemCode") & ""
        itemName = rsBatches("ItemName") & ""
        unitPrice = SafeNum(rsBatches("UnitPrice"))
        qty = SafeNum(rsBatches("Quantity"))
        receivedQty = SafeNum(rsBatches("ReceivedQty"))
        remainingQty = SafeNum(rsBatches("RemainingQty"))
        allocOrders = SafeNum(rsBatches("AllocatedOrders"))
        Dim tColor : tColor = GetTypeColor(itemType)
        Dim usedQty : usedQty = receivedQty - remainingQty
        Dim usagePct : usagePct = 0
        If receivedQty > 0 Then usagePct = (usedQty / receivedQty) * 100
%>
    <div style="background:rgba(255,255,255,0.02); border-radius:8px; padding:14px; margin-bottom:10px; border-left:3px solid <%=tColor%>;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
            <div>
                <span class="type-badge" style="background:<%=tColor%>20; color:<%=tColor%>; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; margin-right:8px;"><%=GetTypeLabel(itemType)%></span>
                <strong style="font-size:14px;"><%=Server.HTMLEncode(batchNo)%></strong>
            </div>
            <span style="color:<%=tColor%>; font-weight:600; font-family:Consolas,monospace;">¥<%=FormatNumber(unitPrice, 4)%>/单位</span>
        </div>
        
        <div style="display:flex; gap:20px; font-size:13px; margin-bottom:6px;">
            <span style="color:#aaa;">物料: <span style="color:#ccc;"><%=Server.HTMLEncode(itemName)%> (<%=Server.HTMLEncode(itemCode)%>)</span></span>
            <span style="color:#aaa;">收货: <span style="color:#4CAF50;"><%=FormatNumber(receivedQty, 2)%></span></span>
            <span style="color:#aaa;">已分配: <span style="color:#FF9800;"><%=FormatNumber(usedQty, 2)%></span></span>
            <span style="color:#aaa;">剩余: <span style="color:<%=IIF(remainingQty>0,"#4CAF50","#888")%>;"><%=FormatNumber(remainingQty, 2)%></span></span>
            <% If allocOrders > 0 Then %>
            <span style="color:#aaa;">关联订单: <span style="color:#2196F3;"><%=allocOrders%> 个</span></span>
            <% End If %>
        </div>
        
        <!-- 使用进度条 -->
        <div style="background:rgba(255,255,255,0.05); border-radius:4px; height:6px; overflow:hidden;">
            <div style="width:<%=usagePct%>%; height:100%; background:<%=tColor%>; border-radius:4px; transition:width 0.3s;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:11px; color:#666; margin-top:4px;">
            <span>已使用 <%=FormatNumber(usagePct,1)%>%</span>
            <span>总价: ¥<%=FormatNumber(unitPrice * receivedQty, 2)%></span>
        </div>
        
        <% If remainingQty > 0 Then %>
        <div style="margin-top:8px; font-size:11px; color:#4CAF50;">
            <i class="fas fa-check-circle"></i> 该批次仍有库存剩余，可用于后续订单分摊
        </div>
        <% Else %>
        <div style="margin-top:8px; font-size:11px; color:#888;">
            <i class="fas fa-check-double"></i> 该批次已完全消耗
        </div>
        <% End If %>
    </div>
<%
        rsBatches.MoveNext
    Loop
    rsBatches.Close
    
    If batchCount = 0 Then
%>
    <div style="text-align:center;color:#888;padding:20px;">
        <i class="fas fa-info-circle"></i> 该收货单暂无批次记录（可能在V10升级前收货）
    </div>
<%
    End If
End If
Set rsBatches = Nothing
%>
</div>
<%
' ============================================
' 数据库辅助函数
' ============================================

Function TableExists_BD(tblName)
    Dim rs, exists
    exists = False
    On Error Resume Next
    Set rs = conn.Execute("SELECT 1 FROM sys.tables WHERE name=N'" & tblName & "'")
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then exists = True
            rs.Close
            Set rs = Nothing
        End If
    Else
        Err.Clear
        conn.Execute "SELECT TOP 1 1 FROM [" & tblName & "]"
        If Err.Number = 0 Then
            exists = True
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0
    TableExists_BD = exists
End Function

Sub EnsureTable_BD(tblName, colDefs)
    On Error Resume Next
    conn.Execute "CREATE TABLE " & tblName & " (" & colDefs & ")"
    If Err.Number <> 0 Then
        Session("DBSetupError") = Session("DBSetupError") & "[BD] Failed to create " & tblName & ": " & Err.Description & "<br>"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Call CloseConnection()
%>
