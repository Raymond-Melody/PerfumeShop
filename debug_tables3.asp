<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()
Response.Write "Checking tables:<br>"

Dim tablesToCheck, tbl, sql, rs
tablesToCheck = Array("ProductReviews", "ReviewLikes", "OrderItems", "OrderDetails", "Orders", "Products", "Users", "UserFavorites")

For Each tbl In tablesToCheck
    On Error Resume Next
    sql = "SELECT TOP 1 1 FROM [" & tbl & "]"
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        Response.Write tbl & ": EXISTS<br>"
        rs.Close
    Else
        Response.Write tbl & ": <b style='color:red'>NOT FOUND</b> (" & Err.Description & ")<br>"
        Err.Clear
    End If
    On Error GoTo 0
    Set rs = Nothing
Next
%>
