<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.Write "START<br>"
Response.Flush
%>
<!--#include file="includes/config.asp"-->
<%
Response.Write "Config loaded<br>"
Response.Flush
%>
<!--#include file="includes/connection.asp"-->
<%
Response.Write "Connection loaded<br>"
Response.Flush
Call OpenConnection()
Response.Write "Connection opened<br>"
Response.Flush

Dim rs : Set rs = ExecuteQuery("SELECT TOP 1 ProductID, ProductName FROM Products WHERE IsActive <> 0")
If Not rs Is Nothing Then
    If Not rs.EOF Then
        Response.Write "Product: " & rs("ProductName") & "<br>"
    End If
    rs.Close : Set rs = Nothing
End If
Response.Write "Query done<br>"
Response.Flush

Call CloseConnection()
Response.Write "DONE<br>"
%>