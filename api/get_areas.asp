<%@ Language="VBScript" CodePage="65001" %>
<%
Response.ContentType = "application/json"
Response.Charset = "UTF-8"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
On Error Resume Next

Call OpenConnection()

Dim parentId, parentName, level
parentId = Request.QueryString("parent_id")
parentName = Request.QueryString("parent_name")
level = Request.QueryString("level")

If parentId = "" Then parentId = 0
If Not IsNumeric(parentId) Then parentId = 0
parentId = CInt(parentId)
If level = "" Then level = 1

Dim sql, rs

' 如果提供了父级名称，先查找父级ID
If parentName <> "" Then
    Dim parentSql, parentRs
    parentSql = "SELECT AreaID FROM Areas WHERE AreaName = '" & SafeSQL(parentName) & "'"
    Set parentRs = ExecuteQuery(parentSql)
    If Not parentRs Is Nothing Then
        If Not parentRs.EOF Then
            parentId = parentRs("AreaID")
        End If
        parentRs.Close
        Set parentRs = Nothing
    End If
End If

sql = "SELECT AreaID, AreaName FROM Areas WHERE ParentID = " & CInt(parentId) & " ORDER BY AreaID"
Set rs = ExecuteQuery(sql)

Dim areasArray(), i
i = 0

If Not rs Is Nothing Then
    If Not rs.EOF Then
        Do While Not rs.EOF
            ReDim Preserve areasArray(i)
            Dim areaName
            areaName = rs("AreaName")
            If IsNull(areaName) Then areaName = ""
            areaName = Replace(areaName, "\", "\\")
            areaName = Replace(areaName, """", "\""")
            areasArray(i) = "{""AreaID"":" & rs("AreaID") & ",""AreaName"":""" & areaName & """}"
            i = i + 1
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End If

Dim jsonResponse
If i > 0 Then
    jsonResponse = "[" & Join(areasArray, ",") & "]"
Else
    jsonResponse = "[]"
End If

Call CloseConnection()

Response.Write jsonResponse

If Err.Number <> 0 Then
    Response.Clear
    Dim errMsg
    errMsg = Err.Description
    If IsNull(errMsg) Then errMsg = ""
    errMsg = Replace(errMsg, "\", "\\")
    errMsg = Replace(errMsg, """", "\""")
    Response.Write "{""error"":""" & errMsg & """,""areas"":[]}"
End If

On Error GoTo 0
%>