<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<% Option Explicit %>
<!--#include file="../../includes/connection.asp"-->
<%
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

Dim tabParam : tabParam = Trim(Request.QueryString("tab"))
If tabParam = "" Then tabParam = Trim(Request.Form("tab"))
If tabParam = "" Then tabParam = "RawMaterial"

Dim sql, rs, jsonParts
jsonParts = ""

On Error Resume Next

Select Case tabParam
    Case "RawMaterial"
        sql = "SELECT TOP 50 MaterialID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost, " & _
              "ISNULL(AvgDailyUsage,0) AS AvgUsage, ISNULL(LeadTimeDays,7) AS LeadDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
              "FROM RawMaterialInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
    Case "Packaging"
        sql = "SELECT TOP 50 PackagingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, '' AS Unit, UnitPrice AS WCost, " & _
              "ISNULL(AvgDailyUsage,0) AS AvgUsage, ISNULL(LeadTimeDays,7) AS LeadDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
              "FROM PackagingInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
    Case "Bottle"
        sql = "SELECT TOP 50 BottleID AS ItemID, BottleName AS ItemName, CAST(BottleID AS NVARCHAR(20)) AS ItemCode, StockQty, SafetyStock, 'pcs' AS Unit, UnitPrice AS WCost, " & _
              "ISNULL(AvgDailyUsage,0) AS AvgUsage, ISNULL(LeadTimeDays,7) AS LeadDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
              "FROM BottleStyles WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
    Case "Printing"
        sql = "SELECT TOP 50 PrintingID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost, " & _
              "ISNULL(AvgDailyUsage,0) AS AvgUsage, ISNULL(LeadTimeDays,7) AS LeadDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
              "FROM PrintingInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
    Case "SprayHead"
        sql = "SELECT TOP 50 SprayHeadID AS ItemID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice AS WCost, " & _
              "ISNULL(AvgDailyUsage,0) AS AvgUsage, ISNULL(LeadTimeDays,7) AS LeadDays, ISNULL(ReorderPoint,0) AS ReorderPoint " & _
              "FROM SprayHeadInventory WHERE StockQty <= SafetyStock * 1.5 AND SafetyStock > 0 ORDER BY (StockQty - SafetyStock) ASC"
    Case Else
        Response.Write "[]"
        Call CloseConnection()
        Response.End
End Select

Set rs = conn.Execute(sql)
If Not rs Is Nothing Then
    Dim recCount : recCount = 0
    Do While Not rs.EOF And recCount < 50
        recCount = recCount + 1
        Dim itemID : itemID = SafeNum(rs("ItemID"))
        Dim itemName : itemName = Replace(Replace(rs("ItemName") & "", "\", "\\"), """", "\""")
        Dim itemCode : itemCode = Replace(Replace(rs("ItemCode") & "", "\", "\\"), """", "\""")
        Dim stockQty : stockQty = SafeNum(rs("StockQty"))
        Dim safetyStock : safetyStock = SafeNum(rs("SafetyStock"))
        Dim unit : unit = Replace(Replace(rs("Unit") & "", "\", "\\"), """", "\""")
        Dim wCost : wCost = SafeNum(rs("WCost"))
        Dim avgUsage : avgUsage = SafeNum(rs("AvgUsage"))
        Dim leadDays : leadDays = SafeNum(rs("LeadDays"))
        If leadDays <= 0 Then leadDays = 7
        
        ' V11: 升级补货算法 - 基于日均消耗和交货周期
        Dim suggestQty
        Dim leadDemand : leadDemand = avgUsage * leadDays * 1.2
        Dim demandBased : demandBased = leadDemand + safetyStock - stockQty
        Dim safetyBased : safetyBased = safetyStock * 1.5 - stockQty
        suggestQty = demandBased
        If safetyBased > suggestQty Then suggestQty = safetyBased
        suggestQty = Round(suggestQty, 0)
        If suggestQty < 1 Then suggestQty = 1
        
        If jsonParts <> "" Then jsonParts = jsonParts & ","
        jsonParts = jsonParts & "{""itemid"":" & itemID & ",""itemname"":""" & itemName & """,""itemcode"":""" & itemCode & """," & _
                    """stock"":" & Replace(FormatNumber(stockQty, 1, -1, 0, 0), ",", "") & ",""safety"":" & Replace(FormatNumber(safetyStock, 1, -1, 0, 0), ",", "") & "," & _
                    """unit"":""" & unit & """,""wcost"":" & Replace(FormatNumber(wCost, 4, -1, 0, 0), ",", "") & "," & _
                    """avgusage"":" & Replace(FormatNumber(avgUsage, 6, -1, 0, 0), ",", "") & ",""leaddays"":" & CLng(leadDays) & ",""suggest"":" & suggestQty & "}"
        
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
