<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' V10: 自动创建成本分摊表
On Error Resume Next
conn.Execute "SELECT TOP 1 1 FROM OrderCostAllocation"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE OrderCostAllocation (AllocationID INT IDENTITY(1,1) PRIMARY KEY, OrderNo NVARCHAR(100), BatchID INT, CostType NVARCHAR(30), ItemCode NVARCHAR(50), ItemName NVARCHAR(200), UnitCost DECIMAL(19,4) DEFAULT 0, Quantity FLOAT DEFAULT 0, TotalCost DECIMAL(19,4) DEFAULT 0, InvBatchID INT, AllocatedAt DATETIME DEFAULT GETDATE())"
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Function GetTypeColor(ct)
    Select Case ct
        Case "RawMaterial" : GetTypeColor = "#FF9800"
        Case "Packaging"   : GetTypeColor = "#2196F3"
        Case "Bottle"      : GetTypeColor = "#9C27B0"
        Case "Printing"    : GetTypeColor = "#00BCD4"
        Case "SprayHead"   : GetTypeColor = "#E91E63"
        Case "Product"     : GetTypeColor = "#4CAF50"
        Case Else          : GetTypeColor = "#888"
    End Select
End Function

Function GetCostTypeLabel(ct)
    Select Case ct
        Case "RawMaterial" : GetCostTypeLabel = "原料"
        Case "Packaging"   : GetCostTypeLabel = "包装"
        Case "Bottle"      : GetCostTypeLabel = "瓶子"
        Case "Printing"    : GetCostTypeLabel = "印刷品"
        Case "SprayHead"   : GetCostTypeLabel = "喷头"
        Case "Product"     : GetCostTypeLabel = "产品"
        Case Else          : GetCostTypeLabel = ct
    End Select
End Function

Dim orderNo, sql, rs, ct
orderNo = Trim(Request.QueryString("orderno"))
If orderNo = "" Then
    Response.Write "<div style='text-align:center;color:#e74c3c;padding:10px;'>无效的订单号</div>"
    Response.End
End If

' 按CostType分组查询明细
sql = "SELECT CostType, ItemCode, ItemName, UnitCost, Quantity, TotalCost, BatchID, InvBatchID, AllocatedAt " & _
      "FROM OrderCostAllocation WHERE OrderNo='" & SafeSQL(orderNo) & "' ORDER BY CostType, ItemName"

Set rs = conn.Execute(sql)

' 按CostType分组汇总
Dim costTypes, costTotals, costItems
Set costTypes = CreateObject("Scripting.Dictionary")
Set costTotals = CreateObject("Scripting.Dictionary")
Set costItems = CreateObject("Scripting.Dictionary")

If Not rs Is Nothing Then
    Do While Not rs.EOF
        ct = rs("CostType") & ""
        If ct = "" Then ct = "Other"
        
        ' 汇总
        If costTotals.Exists(ct) Then
            costTotals(ct) = costTotals(ct) + SafeNum(rs("TotalCost"))
        Else
            costTotals.Add ct, SafeNum(rs("TotalCost"))
        End If
        
        ' 明细列表
        Dim itemInfo
        itemInfo = Array(SafeNum(rs("UnitCost")), SafeNum(rs("Quantity")), SafeNum(rs("TotalCost")), rs("ItemName") & "", rs("ItemCode") & "", SafeNum(rs("BatchID")), SafeNum(rs("InvBatchID")))
        If costItems.Exists(ct) Then
            costItems(ct) = costItems(ct) & "|" & Join(itemInfo, "^")
        Else
            costItems.Add ct, Join(itemInfo, "^")
        End If
        
        rs.MoveNext
    Loop
    rs.Close
End If
Set rs = Nothing

Call CloseConnection()

' 渲染HTML片段
Dim allKeys, k, i
allKeys = costTotals.Keys
%>
<div class="detail-grid">
<%
For i = 0 To UBound(allKeys)
    ct = allKeys(i)
    Dim tColor : tColor = GetTypeColor(ct)
    Dim tTotal : tTotal = costTotals(ct)
%>
    <div style="background:rgba(255,255,255,0.02); border-radius:8px; padding:12px 14px; border-left:3px solid <%=tColor%>;">
        <div style="font-weight:600; font-size:14px; color:<%=tColor%>; margin-bottom:8px;">
            <i class="fas fa-tag"></i> <%=GetCostTypeLabel(ct)%> 小计: ¥<%=FormatNumber(tTotal, 4)%>
        </div>
        <%
        Dim itemRows, rowIdx, rowData
        itemRows = Split(costItems(ct), "|")
        For rowIdx = 0 To UBound(itemRows)
            rowData = Split(itemRows(rowIdx), "^")
            If UBound(rowData) >= 4 Then
        %>
        <div style="display:flex; justify-content:space-between; align-items:center; padding:6px 0; border-bottom:1px solid rgba(255,255,255,0.03); font-size:13px;">
            <span style="flex:2;">
                <span style="color:#ccc;"><%=Server.HTMLEncode(rowData(3))%></span>
                <span style="color:#666;font-size:11px;">(<%=Server.HTMLEncode(rowData(4))%>)</span>
            </span>
            <span style="flex:1;text-align:center;color:#888;font-size:12px;">
                ¥<%=FormatNumber(CDbl(rowData(0)),4)%> × <%=FormatNumber(CDbl(rowData(1)),2)%>
            </span>
            <span style="flex:1;text-align:right;font-weight:600;color:<%=tColor%>;font-family:Consolas,monospace;">
                ¥<%=FormatNumber(CDbl(rowData(2)),4)%>
            </span>
        </div>
        <%
            End If
        Next
        %>
    </div>
<%
Next
%>
</div>
