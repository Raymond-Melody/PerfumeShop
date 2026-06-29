<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim fixed : fixed = 0

On Error Resume Next

' 将所有非标准英文状态统一修复
conn.Execute "UPDATE ProductionOrders SET Status='Cancelled' WHERE Status NOT IN ('Pending','InProgress','Completed','Cancelled','QC_Review','Shipped')"
If Err.Number = 0 Then
    ' 获取受影响行数
    Dim rsCnt
    Set rsCnt = conn.Execute("SELECT COUNT(*) AS Cnt FROM ProductionOrders WHERE Status='Cancelled'")
    If Not rsCnt Is Nothing Then
        If Not rsCnt.EOF Then fixed = CLng(rsCnt("Cnt"))
        rsCnt.Close
    End If
    Set rsCnt = Nothing
End If
Err.Clear

' 同样修复 ProductionLogs
conn.Execute "UPDATE ProductionLogs SET Status='Cancelled' WHERE Status NOT IN ('Pending','InProgress','Completed','Cancelled','QC_Review','Shipped')"
Err.Clear

On Error GoTo 0

' 最终统计
Dim json : json = "{""success"":true,""fixed"":" & fixed & ",""finalStatuses"":{"
Dim rs, first : first = True
Set rs = conn.Execute("SELECT Status, COUNT(*) AS Cnt FROM ProductionOrders GROUP BY Status ORDER BY Status")
Do While Not rs.EOF
    If Not first Then json = json & ","
    json = json & """" & rs("Status") & """:" & rs("Cnt")
    first = False
    rs.MoveNext
Loop
rs.Close
json = json & "}}"

Response.Write json
%>
