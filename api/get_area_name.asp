<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Charset = "UTF-8"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
On Error Resume Next

Call OpenConnection()

Dim areaId
areaId = Request.QueryString("id")

If areaId <> "" And IsNumeric(areaId) Then
    Dim sql, rs
    sql = "SELECT AreaName FROM Areas WHERE AreaID = " & CInt(areaId)
    Set rs = ExecuteQuery(sql)

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            Response.Write rs("AreaName")
        Else
            Response.Write ""
        End If
        rs.Close
        Set rs = Nothing
    End If
End If

Call CloseConnection()

If Err.Number <> 0 Then
    Response.Write ""
End If

On Error GoTo 0
%>