<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Response.Write "A"
Call OpenConnection()
Response.Write "B"

Response.Write "<br>Checking tables...<br>"

Dim tablesStr : tablesStr = "OrderItems|OrderDetails|Orders|Products|ProductReviews|UserFavorites|Users"
Dim arr : arr = Split(tablesStr, "|")
Dim i, t, rs, cnt
For i = 0 To UBound(arr)
    t = arr(i)
    On Error Resume Next
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM sys.tables WHERE name='" & t & "'")
    If Err.Number <> 0 Then
        Response.Write t & ": ERROR " & Err.Description & "<br>"
        Err.Clear
    Else
        cnt = rs("cnt")
        Response.Write t & ": " & cnt & "<br>"
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
Next

Response.Write "<br>Testing OrderItems JOIN...<br>"
On Error Resume Next
Set rs = conn.Execute("SELECT COUNT(*) FROM Orders o JOIN OrderItems oi ON o.OrderID=oi.OrderID")
If Err.Number <> 0 Then
    Response.Write "OrderItems JOIN ERROR: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "OrderItems JOIN OK<br>"
    rs.Close
End If
Set rs = Nothing
On Error GoTo 0

Response.Write "DONE<br>"
Call CloseConnection()
%>
