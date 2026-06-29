<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
On Error Resume Next
Response.Write "1-OpenConn... "
Call OpenConnection()
Response.Write "OK<br>"

Response.Write "2-CreateDict... "
Dim dictResult : Set dictResult = Server.CreateObject("Scripting.Dictionary")
Response.Write "OK<br>"

Response.Write "3-Strategy1-SQL..."
Dim userId : userId = 25
Dim topN : topN = 8
sql = "SELECT TOP " & topN & " p2.ProductID, MAX(p2.ProductName) AS ProductName, MAX(p2.BasePrice) AS BasePrice, " & _
      "MAX(p2.ImageURL) AS ImageURL, MAX(p2.ProductType) AS ProductType " & _
      "FROM OrderDetails od1 " & _
      "INNER JOIN Products p1 ON od1.ProductID = p1.ProductID " & _
      "INNER JOIN Products p2 ON p1.ProductType = p2.ProductType AND p2.IsActive = 1 " & _
      "INNER JOIN Orders o ON od1.OrderID = o.OrderID " & _
      "WHERE o.UserID = " & userId & " AND p2.ProductID NOT IN " & _
      "(SELECT ProductID FROM OrderDetails od2 INNER JOIN Orders o2 ON od2.OrderID = o2.OrderID WHERE o2.UserID = " & userId & ") " & _
      "AND p2.ProductID NOT IN (SELECT ProductID FROM UserFavorites WHERE UserID = " & userId & ") " & _
      "GROUP BY p2.ProductID ORDER BY NEWID()"
Response.Write "OK<br>"

Response.Write "4-ExecStrategy1..."
Dim rs : Set rs = conn.Execute(sql)
If Err.Number <> 0 Then
    Response.Write "ERROR(" & Err.Number & "): " & Err.Description & "<br>"
    Err.Clear
    Set rs = Nothing
ElseIf rs Is Nothing Then
    Response.Write "NULL<br>"
Else
    Response.Write "Got RS, EOF=" & rs.EOF & "<br>"
    If Not rs.EOF Then
        Response.Write "  Row1: ProductID=" & rs("ProductID") & " Name=" & rs("ProductName") & "<br>"
    End If
    rs.Close : Set rs = Nothing
End If

Response.Write "5-Strategy2-SQL..."
sql = "SELECT DISTINCT TOP " & topN & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType " & _
      "FROM UserFavorites uf INNER JOIN Products p ON uf.ProductID = p.ProductID " & _
      "WHERE uf.UserID = " & userId & " AND p.IsActive = 1"
Set rs = conn.Execute(sql)
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & "<br>"
    Err.Clear
    Set rs = Nothing
ElseIf rs Is Nothing Then
    Response.Write "NULL<br>"
Else
    Response.Write "OK, EOF=" & rs.EOF & "<br>"
    rs.Close : Set rs = Nothing
End If

Response.Write "6-Strategy3-SQL..."
sql = "SELECT TOP " & topN & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType, " & _
      "ISNULL((SELECT COUNT(*) FROM OrderDetails od INNER JOIN Orders o ON od.OrderID = o.OrderID " & _
      "WHERE od.ProductID = p.ProductID AND o.Status='Paid'), 0) AS SaleCount " & _
      "FROM Products p WHERE p.IsActive = 1 AND p.ProductID NOT IN " & _
      "(SELECT od.ProductID FROM OrderDetails od INNER JOIN Orders o ON od.OrderID = o.OrderID WHERE o.UserID = " & userId & ") " & _
      "ORDER BY SaleCount DESC"
Set rs = conn.Execute(sql)
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & "<br>"
    Err.Clear
    Set rs = Nothing
ElseIf rs Is Nothing Then
    Response.Write "NULL<br>"
Else
    Response.Write "OK, EOF=" & rs.EOF & "<br>"
    rs.Close : Set rs = Nothing
End If

Response.Write "7-DictReturn..."
Set RE_GetUserRecommendations = dictResult
Response.Write "OK<br>"

Response.Write "ALL DONE<br>"
Call CloseConnection()
%>
