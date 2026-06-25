<%@ Language="VBScript" CodePage="65001" %>
<%
Response.ContentType = "application/json"
Response.Charset = "UTF-8"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
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

' 如果提供了父级名称，先查找父级ID
If parentName <> "" Then
    Dim parentParams(0)
    parentParams(0) = Array("@AreaName", DAL_adVarChar, 100, parentName)
    Dim parentRow : Set parentRow = DAL_GetRow("SELECT AreaID FROM Areas WHERE AreaName=@AreaName", parentParams)
    If Not parentRow Is Nothing Then
        parentId = parentRow("AreaID")
    End If
    Set parentRow = Nothing
End If

Dim rs
Set rs = DAL_GetList("SELECT AreaID, AreaName FROM Areas WHERE ParentID=@ParentID ORDER BY AreaID", _
    Array(Array("@ParentID", DAL_adInteger, 0, parentId)))

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