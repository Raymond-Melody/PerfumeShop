<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<% Option Explicit %>
<!--#include file="../../includes/connection.asp"-->
<%
' ========== 历史采购产品查询 AJAX 接口 ==========
' 参数: ordertype - 采购类型 (RawMaterial/Packaging/Bottle/Printing/SprayHead)
'       search - 搜索关键词（名称/编码模糊匹配）
'       supplier - 供应商ID（可选过滤）
'       limit - 返回数量限制（默认50）

Response.ContentType = "application/json"
Response.Charset = "UTF-8"

Call OpenConnection()
conn.CommandTimeout = 30

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Dim orderType : orderType = Trim(Request.QueryString("ordertype"))
If orderType = "" Then orderType = Trim(Request.Form("ordertype"))
If orderType = "" Then orderType = "RawMaterial"

Dim searchKey : searchKey = Trim(Request.QueryString("search"))
If searchKey = "" Then searchKey = Trim(Request.Form("search"))

Dim supplierID : supplierID = SafeNum(Request.QueryString("supplier"))
If supplierID = 0 Then supplierID = SafeNum(Request.Form("supplier"))

Dim limitNum : limitNum = SafeNum(Request.QueryString("limit"))
If limitNum <= 0 Then limitNum = 50

' 查询历史采购记录：从 PurchaseOrderDetails 中按物料名称/编码去重
Dim sql, rs
sql = "SELECT DISTINCT pod.ItemName, pod.ItemCode, pod.Specification, pod.Unit, " & _
      "LAST_VALUE(pod.UnitPrice) OVER (PARTITION BY pod.ItemName, ISNULL(pod.ItemCode,'') ORDER BY po.OrderDate ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING) AS LastPrice, " & _
      "MAX(po.OrderDate) AS LastOrderDate, " & _
      "LAST_VALUE(s.SupplierName) OVER (PARTITION BY pod.ItemName, ISNULL(pod.ItemCode,'') ORDER BY po.OrderDate ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING) AS LastSupplier, " & _
      "LAST_VALUE(po.SupplierID) OVER (PARTITION BY pod.ItemName, ISNULL(pod.ItemCode,'') ORDER BY po.OrderDate ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING) AS LastSupplierID, " & _
      "COUNT(*) OVER (PARTITION BY pod.ItemName, ISNULL(pod.ItemCode,'')) AS PurchaseCount " & _
      "FROM PurchaseOrderDetails pod " & _
      "INNER JOIN PurchaseOrders po ON pod.PurchaseID = po.PurchaseID " & _
      "LEFT JOIN Suppliers s ON po.SupplierID = s.SupplierID " & _
      "WHERE po.OrderType = '" & SafeSQL(orderType) & "' AND pod.ItemName IS NOT NULL AND pod.ItemName <> '' "

If searchKey <> "" Then
    sql = sql & "AND (pod.ItemName LIKE '%" & SafeSQL(searchKey) & "%' OR pod.ItemCode LIKE '%" & SafeSQL(searchKey) & "%') "
End If

If supplierID > 0 Then
    sql = sql & "AND po.SupplierID = " & CLng(supplierID) & " "
End If

sql = sql & "ORDER BY LastOrderDate DESC"

On Error Resume Next
Set rs = conn.Execute("SELECT TOP " & CLng(limitNum) & " * FROM (" & sql & ") AS HistoryProducts")

' Fallback: 兼容不支持窗口函数的旧版 SQL Server
If Err.Number <> 0 Then
    Err.Clear
    sql = "SELECT pod.ItemName, pod.ItemCode, MAX(pod.Specification) AS Specification, MAX(pod.Unit) AS Unit, " & _
          "MAX(pod.UnitPrice) AS LastPrice, MAX(po.OrderDate) AS LastOrderDate, " & _
          "(SELECT TOP 1 s2.SupplierName FROM PurchaseOrderDetails pod2 INNER JOIN PurchaseOrders po2 ON pod2.PurchaseID=po2.PurchaseID LEFT JOIN Suppliers s2 ON po2.SupplierID=s2.SupplierID WHERE pod2.ItemName=pod.ItemName AND ISNULL(pod2.ItemCode,'')=ISNULL(pod.ItemCode,'') AND po2.OrderType='" & SafeSQL(orderType) & "' ORDER BY po2.OrderDate DESC) AS LastSupplier, " & _
          "(SELECT TOP 1 po3.SupplierID FROM PurchaseOrderDetails pod3 INNER JOIN PurchaseOrders po3 ON pod3.PurchaseID=po3.PurchaseID WHERE pod3.ItemName=pod.ItemName AND ISNULL(pod3.ItemCode,'')=ISNULL(pod.ItemCode,'') AND po3.OrderType='" & SafeSQL(orderType) & "' ORDER BY po3.OrderDate DESC) AS LastSupplierID, " & _
          "COUNT(DISTINCT po.PurchaseID) AS PurchaseCount " & _
          "FROM PurchaseOrderDetails pod INNER JOIN PurchaseOrders po ON pod.PurchaseID=po.PurchaseID " & _
          "WHERE po.OrderType='" & SafeSQL(orderType) & "' AND pod.ItemName IS NOT NULL AND pod.ItemName <> '' "

    If searchKey <> "" Then
        sql = sql & "AND (pod.ItemName LIKE '%" & SafeSQL(searchKey) & "%' OR pod.ItemCode LIKE '%" & SafeSQL(searchKey) & "%') "
    End If

    If supplierID > 0 Then
        sql = sql & "AND po.SupplierID=" & CLng(supplierID) & " "
    End If

    sql = sql & "GROUP BY pod.ItemName, pod.ItemCode ORDER BY MAX(po.OrderDate) DESC"
    Set rs = conn.Execute(sql)
    If Err.Number <> 0 Then Err.Clear
End If

Dim jsonParts : jsonParts = ""
Dim recCount : recCount = 0

If Not rs Is Nothing Then
    Do While Not rs.EOF And recCount < CLng(limitNum)
        recCount = recCount + 1
        Dim itemName : itemName = Replace(Replace(rs("ItemName") & "", "\", "\\"), """", "\""")
        Dim itemCode : itemCode = Replace(Replace(rs("ItemCode") & "", "\", "\\"), """", "\""")
        Dim spec : spec = Replace(Replace(rs("Specification") & "", "\", "\\"), """", "\""")
        Dim unit : unit = Replace(Replace(rs("Unit") & "", "\", "\\"), """", "\""")
        Dim lastPrice : lastPrice = SafeNum(rs("LastPrice"))
        Dim lastSupplier : lastSupplier = Replace(Replace(rs("LastSupplier") & "", "\", "\\"), """", "\""")
        Dim lastSupplierID : lastSupplierID = SafeNum(rs("LastSupplierID"))
        Dim purchaseCount : purchaseCount = SafeNum(rs("PurchaseCount"))
        Dim lastOrderDate : lastOrderDate = ""
        If Not IsNull(rs("LastOrderDate")) And IsDate(rs("LastOrderDate")) Then
            lastOrderDate = FormatDateTime(rs("LastOrderDate"), 2)
        End If
        
        If jsonParts <> "" Then jsonParts = jsonParts & ","
        jsonParts = jsonParts & "{""itemname"":""" & itemName & """,""itemcode"":""" & itemCode & """," & _
                    """spec"":""" & spec & """,""unit"":""" & unit & """,""lastprice"":" & Replace(FormatNumber(lastPrice, 4, -1, 0, 0), ",", "") & "," & _
                    """lastsupplier"":""" & lastSupplier & """,""lastsupplierid"":" & lastSupplierID & "," & _
                    """purchasecount"":" & CLng(purchaseCount) & ",""lastorderdate"":""" & lastOrderDate & """}"
        
        rs.MoveNext
    Loop
    rs.Close
End If
Set rs = Nothing

If Err.Number <> 0 Then
    jsonParts = ""
    Err.Clear
End If
On Error GoTo 0

Response.Write "[" & jsonParts & "]"

Call CloseConnection()
%>
