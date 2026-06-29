<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim json, rs
json = "{"

' 1. BaseNotes with prices
Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM BaseNotes WHERE IsActive <> 0")
json = json & """activeBaseNotes"":" & rs("Cnt") & ","
rs.Close

Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM BaseNotes WHERE IsActive <> 0 AND UnitPrice > 0")
json = json & """baseNotesWithPrice"":" & rs("Cnt") & ","
rs.Close

Set rs = conn.Execute("SELECT TOP 5 BaseNoteID, BaseNoteName, UnitPrice FROM BaseNotes WHERE IsActive <> 0 ORDER BY UnitPrice DESC")
json = json & """baseNoteSamples"":["
Dim first : first = True
Do While Not rs.EOF
    If Not first Then json = json & ","
    json = json & "{""id"":" & rs("BaseNoteID") & ",""name"":""" & Replace(rs("BaseNoteName") & "", """", "\""") & """,""price"":" & CDbl(rs("UnitPrice")) & "}"
    first = False
    rs.MoveNext
Loop
json = json & "],"
rs.Close

' 2. Products with cost
Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM Products")
json = json & """totalProducts"":" & rs("Cnt") & ","
rs.Close

Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM Products WHERE UnitCost > 0")
json = json & """productsWithCost"":" & rs("Cnt") & ","
rs.Close

Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM Products WHERE BasePrice > 0")
json = json & """productsWithPrice"":" & rs("Cnt") & ","
rs.Close

' 3. Raw material inventory with cost
Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM RawMaterialInventory WHERE StockQty > 0")
json = json & """materialsWithStock"":" & rs("Cnt") & ","
rs.Close

Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM RawMaterialInventory WHERE StockQty > 0 AND UnitPrice > 0")
json = json & """materialsWithPrice"":" & rs("Cnt") & ","
rs.Close

' 4. Products cost sample
Set rs = conn.Execute("SELECT TOP 5 ProductID, ProductName, BasePrice, ISNULL(UnitCost,0) AS UnitCost, ISNULL(BOMCost,0) AS BOMCost, ProductType FROM Products ORDER BY ProductID")
json = json & """productSamples"":["
first = True
Do While Not rs.EOF
    If Not first Then json = json & ","
    json = json & "{""id"":" & rs("ProductID") & ",""name"":""" & Replace(rs("ProductName") & "", """", "\""") & """,""basePrice"":" & CDbl(rs("BasePrice")) & ",""unitCost"":" & CDbl(rs("UnitCost")) & ",""bomCost"":" & CDbl(rs("BOMCost")) & ",""type"":""" & rs("ProductType") & """}"
    first = False
    rs.MoveNext
Loop
json = json & "],"
rs.Close

' 5. Supplier prices
Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM SupplierPrices WHERE IsActive=1")
json = json & """activeSupplierPrices"":" & rs("Cnt") & ","
rs.Close

' 6. 财务总览
Set rs = conn.Execute("SELECT ISNULL(SUM(TotalAmount),0) AS Revenue, ISNULL(SUM(ISNULL(TotalCost,0)),0) AS TotalCost FROM Orders WHERE Status NOT IN ('Cancelled','Pending')")
json = json & """financeSummary"":{"
If Not rs.EOF Then
    json = json & """revenue"":" & CDbl(rs("Revenue")) & ",""totalCost"":" & CDbl(rs("TotalCost"))
End If
json = json & "},"
rs.Close

' 7. 有成本的订单数
Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM Orders WHERE TotalCost > 0")
json = json & """ordersWithCost"":" & rs("Cnt")
rs.Close

json = json & "}"
Response.Write json
%>
