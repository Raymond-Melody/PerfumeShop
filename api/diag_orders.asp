<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim json, rs, sql

json = "{"

' 1. 订单统计
sql = "SELECT Status, COUNT(*) AS Cnt FROM Orders GROUP BY Status ORDER BY Status"
Set rs = conn.Execute(sql)
json = json & """orderStatuses"":{"
Dim first : first = True
Do While Not rs.EOF
    If Not first Then json = json & ","
    json = json & """" & rs("Status") & """:" & rs("Cnt")
    first = False
    rs.MoveNext
Loop
rs.Close
json = json & "},"

' 2. ProductionOrders 总数
sql = "SELECT COUNT(*) AS Cnt FROM ProductionOrders"
Set rs = conn.Execute(sql)
json = json & """productionOrdersTotal"":" & rs("Cnt") & ","
rs.Close

' 3. ProductionOrders 按状态
sql = "SELECT Status, COUNT(*) AS Cnt FROM ProductionOrders GROUP BY Status ORDER BY Status"
Set rs = conn.Execute(sql)
json = json & """productionOrderStatuses"":{"
first = True
Do While Not rs.EOF
    If Not first Then json = json & ","
    json = json & """" & rs("Status") & """:" & rs("Cnt")
    first = False
    rs.MoveNext
Loop
rs.Close
json = json & "},"

' 4. 检查第一个已支付订单
sql = "SELECT TOP 1 OrderID, OrderNo, Status FROM Orders WHERE Status='Paid' ORDER BY OrderID"
Set rs = conn.Execute(sql)
json = json & """firstPaidOrder"":"
If Not rs.EOF Then
    json = json & "{" & """OrderID"":" & rs("OrderID") & ",""OrderNo"":""" & rs("OrderNo") & """,""Status"":""" & rs("Status") & """}"
Else
    json = json & "null"
End If
json = json & ","
rs.Close

' 5. 该订单是否有 OrderDetails
sql = "SELECT COUNT(*) AS Cnt FROM OrderDetails WHERE OrderID = (SELECT TOP 1 OrderID FROM Orders WHERE Status='Paid' ORDER BY OrderID)"
Set rs = conn.Execute(sql)
json = json & """firstPaidOrderDetails"":" & rs("Cnt") & ","
rs.Close

' 6. 该订单是否有 ProductionOrders
sql = "SELECT COUNT(*) AS Cnt FROM ProductionOrders WHERE OrderID = (SELECT TOP 1 OrderID FROM Orders WHERE Status='Paid' ORDER BY OrderID)"
Set rs = conn.Execute(sql)
json = json & """firstPaidOrderProdOrders"":" & rs("Cnt") & ","
rs.Close

' 7. OrderDetails 总数
sql = "SELECT COUNT(*) AS Cnt FROM OrderDetails"
Set rs = conn.Execute(sql)
json = json & """orderDetailsTotal"":" & rs("Cnt") & ","
rs.Close

' 8. OrderItems 总数
sql = "SELECT COUNT(*) AS Cnt FROM OrderItems"
Set rs = conn.Execute(sql)
If Err.Number = 0 Then
    json = json & """orderItemsTotal"":" & rs("Cnt")
    rs.Close
Else
    Err.Clear
    json = json & """orderItemsTotal"":" & """table_not_found"""
End If

json = json & "}"

Response.Write json
%>
