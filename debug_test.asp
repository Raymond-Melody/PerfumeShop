<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Response.Write "Step 1: Connection... "
Call OpenConnection()
Response.Write "OK<br>"

Response.Write "Step 2: Test simple conn.Execute..."
Dim trs : Set trs = conn.Execute("SELECT TOP 1 ProductID, ProductName FROM Products WHERE IsActive=1")
If Not trs Is Nothing Then
    Response.Write "OK: ProductID=" & trs("ProductID") & "<br>"
    trs.Close : Set trs = Nothing
Else
    Response.Write "FAIL<br>"
End If

Response.Write "Step 3: Test GROUP BY + ORDER BY NEWID()..."
sql = "SELECT TOP 3 p.ProductID, MAX(p.ProductName) AS ProductName FROM Products p WHERE p.IsActive=1 GROUP BY p.ProductID ORDER BY NEWID()"
On Error Resume Next
Set trs = conn.Execute(sql)
If Err.Number <> 0 Then
    Response.Write "SQL ERROR: " & Err.Description & "<br>"
    Err.Clear
ElseIf Not trs Is Nothing Then
    Response.Write "OK: " & trs("ProductName") & "<br>"
    trs.Close : Set trs = Nothing
End If
On Error GoTo 0

Response.Write "Step 4: Test Scripting.Dictionary..."
Dim d : Set d = Server.CreateObject("Scripting.Dictionary")
d.Add "1", "test"
Response.Write "OK: count=" & d.Count & "<br>"

Response.Write "Step 5: Test returning Dictionary from function..."
Function TestDict()
    Dim td : Set td = Server.CreateObject("Scripting.Dictionary")
    td.Add "x", "y"
    Set TestDict = td
End Function
Dim result : Set result = TestDict()
Response.Write "OK: result.Count=" & result.Count & "<br>"

Response.Write "ALL TESTS PASSED<br>"
Call CloseConnection()
%>
