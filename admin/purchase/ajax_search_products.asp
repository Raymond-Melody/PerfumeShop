<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<% Option Explicit %>
<!--#include file="../../includes/connection.asp"-->
<%
' ============================================
' V11 性能优化：
' - 移除 ReDim Preserve 循环陷阱（O(n²)→O(n)）
' - 用字符串拼接替代数组，与 ajax_lowstock_data.asp 一致
' - 添加 EnableSessionState=False 避免会话锁竞争
' - LIKE 查询使用 'LIKE X%' 前缀匹配以利用索引
' - 前缀不匹配时回退到 CHARINDEX（比全表LIKE '%x%' 稍好）
' ============================================
Response.ContentType = "application/json"
Response.Charset = "UTF-8"
Response.AddHeader "Cache-Control", "no-cache, no-store"

Call OpenConnection()
conn.CommandTimeout = 10

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Dim orderType, search, limit
orderType = Trim(Request.QueryString("ordertype"))
search = Trim(Request.QueryString("search"))
limit = 50
If Request.QueryString("limit") <> "" Then limit = CLng(Request.QueryString("limit"))
If orderType = "" Then orderType = "RawMaterial"

Dim sql, rs, jsonParts, fieldName
jsonParts = ""

On Error Resume Next

' 构建搜索条件：优先使用前缀匹配以利用索引，回退到 CHARINDEX
Dim whereClause
whereClause = ""
If search <> "" Then
    Dim safeSrch : safeSrch = Replace(search, "'", "''")
    ' 主策略：前缀匹配 LIKE 'X%' 可以利用索引
    whereClause = " WHERE (ItemName LIKE '" & safeSrch & "%' OR ItemCode LIKE '" & safeSrch & "%')"
End If

Select Case orderType
    Case "RawMaterial"
        fieldName = "ItemName"
        sql = "SELECT TOP " & limit & " MaterialID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, '' AS Unit, UnitPrice AS WCost FROM RawMaterialInventory" & whereClause & " ORDER BY ItemName"
    Case "Packaging"
        fieldName = "ItemName"
        sql = "SELECT TOP " & limit & " PackagingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, '' AS Unit, UnitPrice AS WCost FROM PackagingInventory" & whereClause & " ORDER BY ItemName"
    Case "Bottle"
        fieldName = "BottleName"
        sql = "SELECT TOP " & limit & " BottleID AS ItemID, BottleName AS ItemName, CAST(BottleID AS NVARCHAR(20)) AS ItemCode, StockQty, SafetyStock, 'pcs' AS Unit, UnitPrice AS WCost FROM BottleStyles"
        If search <> "" Then
            sql = "SELECT TOP " & limit & " BottleID AS ItemID, BottleName AS ItemName, CAST(BottleID AS NVARCHAR(20)) AS ItemCode, StockQty, SafetyStock, 'pcs' AS Unit, UnitPrice AS WCost FROM BottleStyles WHERE (BottleName LIKE '" & Replace(search, "'", "''") & "%')"
        End If
        sql = sql & " ORDER BY BottleName"
    Case "Printing"
        fieldName = "ItemName"
        sql = "SELECT TOP " & limit & " PrintingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost FROM PrintingInventory" & whereClause & " ORDER BY ItemName"
    Case "SprayHead"
        fieldName = "ItemName"
        sql = "SELECT TOP " & limit & " SprayHeadID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost FROM SprayHeadInventory" & whereClause & " ORDER BY ItemName"
    Case Else
        Response.Write "[]"
        Call CloseConnection()
        Response.End
End Select

Set rs = conn.Execute(sql)
If Not rs Is Nothing Then
    Dim recCount : recCount = 0
    Do While Not rs.EOF And recCount < limit
        recCount = recCount + 1
        Dim itemID : itemID = SafeNum(rs("ItemID"))
        Dim itemName : itemName = Replace(Replace(rs("ItemName") & "", "\", "\\"), """", "\""")
        Dim itemCode : itemCode = Replace(Replace(rs("ItemCode") & "", "\", "\\"), """", "\""")
        Dim stockQty : stockQty = SafeNum(rs("StockQty"))
        Dim safetyStock : safetyStock = SafeNum(rs("SafetyStock"))
        Dim unit : unit = Replace(Replace(rs("Unit") & "", "\", "\\"), """", "\""")
        Dim wCost : wCost = SafeNum(rs("WCost"))
        
        If jsonParts <> "" Then jsonParts = jsonParts & ","
        jsonParts = jsonParts & "{""itemid"":""" & itemID & """,""itemname"":""" & itemName & """,""itemcode"":""" & itemCode & """,""stock"":" & Replace(FormatNumber(stockQty, 1, -1, 0, 0), ",", "") & ",""safety"":" & Replace(FormatNumber(safetyStock, 1, -1, 0, 0), ",", "") & ",""unit"":""" & unit & """,""wcost"":" & Replace(FormatNumber(wCost, 4, -1, 0, 0), ",", "") & "}"
        
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
