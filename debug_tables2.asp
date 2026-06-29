<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Response.Write "A"
Call OpenConnection()
Response.Write "B<br>"

Dim myRS

Response.Write "Check OrderItems...<br>"
On Error Resume Next
Set myRS = conn.Execute("SELECT COUNT(*) AS cnt FROM sys.tables WHERE name='OrderItems'")
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "cnt=" & myRS("cnt") & "<br>"
    myRS.Close
End If
Set myRS = Nothing
On Error GoTo 0

Response.Write "Check Products...<br>"
On Error Resume Next
Set myRS = conn.Execute("SELECT COUNT(*) AS cnt FROM sys.tables WHERE name='Products'")
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "cnt=" & myRS("cnt") & "<br>"
    myRS.Close
End If
Set myRS = Nothing
On Error GoTo 0

Response.Write "DONE<br>"
Call CloseConnection()
%>
