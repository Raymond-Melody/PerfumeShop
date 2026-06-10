<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<% Option Explicit %>
<!--#include file="../../includes/connection.asp"-->
<%
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

Call OpenConnection()

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Dim itemCode, orderType, jsonResult
itemCode = Trim(Request.QueryString("itemcode"))
orderType = Trim(Request.QueryString("ordertype"))

Dim resultSupplierID, resultSupplierName, resultUnitPrice
resultSupplierID = 0
resultSupplierName = ""
resultUnitPrice = 0

If itemCode = "" Then
    Response.Write "{""supplier_id"":0,""supplier_name"":"""",""unit_price"":0}"
    Call CloseConnection()
    Response.End
End If

On Error Resume Next

' Step 1: 从 SupplierPrices 查询最新供应商报价
Dim rs, sql
sql = "SELECT TOP 1 sp.SupplierID, s.SupplierName, sp.UnitPrice " & _
      "FROM SupplierPrices sp " & _
      "LEFT JOIN Suppliers s ON sp.SupplierID = s.SupplierID " & _
      "WHERE sp.ItemCode = '" & SafeSQL(itemCode) & "' AND sp.IsActive = 1 " & _
      "ORDER BY sp.CreatedAt DESC"

Set rs = conn.Execute(sql)
If Not rs Is Nothing And Not rs.EOF Then
    resultSupplierID = SafeNum(rs("SupplierID"))
    resultSupplierName = rs("SupplierName") & ""
    resultUnitPrice = SafeNum(rs("UnitPrice"))
    rs.Close
End If
Set rs = Nothing

' Step 2: 如果在 SupplierPrices 未找到，尝试从库存表的 LastPurchaseDate 关联的采购订单查询
If resultSupplierID = 0 Then
    Dim tableName, idField, nameField
    Select Case orderType
        Case "RawMaterial": tableName = "RawMaterialInventory": idField = "MaterialID"
        Case "Packaging":   tableName = "PackagingInventory": idField = "PackagingID"
        Case "Bottle":      tableName = "BottleStyles":        idField = "BottleID"
        Case "Printing":    tableName = "PrintingInventory":   idField = "PrintingID"
        Case "SprayHead":   tableName = "SprayHeadInventory":  idField = "SprayHeadID"
        Case Else:          tableName = "": idField = ""
    End Select
    
    If tableName <> "" Then
        ' 尝试通过采购订单历史获取供应商
        Dim itemName
        If orderType = "Bottle" Then
            ' Bottle itemCode is now CAST(BottleID AS NVARCHAR(20)), directly use BottleID
            If IsNumeric(itemCode) Then
                sql = "SELECT BottleName FROM BottleStyles WHERE BottleID=" & CLng(itemCode)
                Set rs = conn.Execute(sql)
                If Not rs Is Nothing And Not rs.EOF Then
                    itemName = rs(0) & ""
                    rs.Close
                End If
                Set rs = Nothing
            End If
        Else
            sql = "SELECT ItemName FROM " & tableName & " WHERE ItemCode='" & SafeSQL(itemCode) & "'"
            Set rs = conn.Execute(sql)
            If Not rs Is Nothing And Not rs.EOF Then
                itemName = rs(0) & ""
                rs.Close
            End If
            Set rs = Nothing
        End If
        
        If itemName <> "" Then
            sql = "SELECT TOP 1 po.SupplierID, s.SupplierName, " & _
                  "(SELECT TOP 1 pod.UnitPrice FROM PurchaseOrderDetails pod " & _
                  "  LEFT JOIN PurchaseOrders po2 ON pod.PurchaseID = po2.PurchaseID " & _
                  "  WHERE po2.PurchaseID = po.PurchaseID AND pod.ItemName = '" & SafeSQL(itemName) & "' " & _
                  "  ORDER BY po2.OrderDate DESC) AS UnitPrice " & _
                  "FROM PurchaseOrders po " & _
                  "LEFT JOIN Suppliers s ON po.SupplierID = s.SupplierID " & _
                  "WHERE po.OrderType = '" & SafeSQL(orderType) & "' " & _
                  "ORDER BY po.OrderDate DESC"
            Set rs = conn.Execute(sql)
            If Not rs Is Nothing And Not rs.EOF Then
                resultSupplierID = SafeNum(rs("SupplierID"))
                resultSupplierName = rs("SupplierName") & ""
                resultUnitPrice = SafeNum(rs("UnitPrice"))
                rs.Close
            End If
            Set rs = Nothing
        End If
    End If
End If

If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

' 构建JSON响应（手动构建避免JSON库依赖）
jsonResult = "{"
jsonResult = jsonResult & """supplier_id"":" & CStr(resultSupplierID) & ","
jsonResult = jsonResult & """supplier_name"":""" & Replace(Replace(resultSupplierName, "\", "\\"), """", "\""") & ""","
jsonResult = jsonResult & """unit_price"":" & Replace(FormatNumber(resultUnitPrice, 4, -1, 0, 0), ",", "") & ""
jsonResult = jsonResult & "}"

Response.Write jsonResult

Call CloseConnection()
%>