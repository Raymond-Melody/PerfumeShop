<%@Language="VBScript" CodePage="65001"%>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
Response.AddHeader "Cache-Control", "no-cache, no-store"

' Disable session for async performance
Session.CodePage = 65001

%>
<!--#include file="../../includes/config.asp"-->
<%
Call OpenConnection()

Dim orderType, search, limit
orderType = Trim(Request.QueryString("ordertype"))
search = Trim(Request.QueryString("search"))
limit = 50
If Request.QueryString("limit") <> "" Then limit = CLng(Request.QueryString("limit"))

If orderType = "" Then orderType = "RawMaterial"

Dim sql, rs, items(), itemCount
itemCount = 0

On Error Resume Next
Select Case orderType
    Case "RawMaterial"
        sql = "SELECT TOP " & limit & " MaterialID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, '' AS Unit, UnitPrice AS WCost FROM RawMaterialInventory"
        If search <> "" Then
            sql = sql & " WHERE (ItemName LIKE '%" & Replace(search, "'", "''") & "%' OR ItemCode LIKE '%" & Replace(search, "'", "''") & "%')"
        End If
        sql = sql & " ORDER BY ItemName"
    Case "Packaging"
        sql = "SELECT TOP " & limit & " PackagingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, '' AS Unit, UnitPrice AS WCost FROM PackagingInventory"
        If search <> "" Then
            sql = sql & " WHERE (ItemName LIKE '%" & Replace(search, "'", "''") & "%' OR ItemCode LIKE '%" & Replace(search, "'", "''") & "%')"
        End If
        sql = sql & " ORDER BY ItemName"
    Case "Bottle"
        sql = "SELECT TOP " & limit & " BottleID AS ItemID, BottleName AS ItemName, CAST(BottleID AS NVARCHAR(20)) AS ItemCode, StockQty, SafetyStock, 'pcs' AS Unit, UnitPrice AS WCost FROM BottleStyles"
        If search <> "" Then
            sql = sql & " WHERE (BottleName LIKE '%" & Replace(search, "'", "''") & "%' OR CAST(BottleID AS NVARCHAR(20)) LIKE '%" & Replace(search, "'", "''") & "%')"
        End If
        sql = sql & " ORDER BY BottleName"
    Case "Printing"
        sql = "SELECT TOP " & limit & " PrintingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost FROM PrintingInventory"
        If search <> "" Then
            sql = sql & " WHERE (ItemName LIKE '%" & Replace(search, "'", "''") & "%' OR ItemCode LIKE '%" & Replace(search, "'", "''") & "%')"
        End If
        sql = sql & " ORDER BY ItemName"
    Case "SprayHead"
        sql = "SELECT TOP " & limit & " SprayHeadID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost FROM SprayHeadInventory"
        If search <> "" Then
            sql = sql & " WHERE (ItemName LIKE '%" & Replace(search, "'", "''") & "%' OR ItemCode LIKE '%" & Replace(search, "'", "''") & "%')"
        End If
        sql = sql & " ORDER BY ItemName"
End Select

Set rs = conn.Execute(sql)
If Not rs Is Nothing Then
    Do While Not rs.EOF
        ReDim Preserve items(itemCount)
        Dim itemData : itemData = "{""itemid"":""" & rs("ItemID") & """,""itemname"":""" & Replace(rs("ItemName") & "", """", "\""") & """,""itemcode"":""" & Replace(rs("ItemCode") & "", """", "\""") & """,""stock"":" & SafeNum(rs("StockQty")) & ",""safety"":" & SafeNum(rs("SafetyStock")) & ",""unit"":""" & Replace(rs("Unit") & "", """", "\""") & """,""wcost"":" & SafeNum(rs("WCost")) & "}"
        items(itemCount) = itemData
        itemCount = itemCount + 1
        rs.MoveNext
    Loop
    rs.Close
End If
Set rs = Nothing
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

Call CloseConnection()

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Response.Write "[" & Join(items, ",") & "]"
%>
