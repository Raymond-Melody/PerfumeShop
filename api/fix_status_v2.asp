<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim json, rs
json = "{"

' 显示前20条ProductionOrders的Status原始值
Set rs = conn.Execute("SELECT TOP 20 ProductionID, WorkOrderNo, Status, CAST(Status AS varbinary(max)) AS StatusBytes FROM ProductionOrders ORDER BY ProductionID")
json = json & """rawStatuses"":["
Dim first : first = True
Do While Not rs.EOF
    If Not first Then json = json & ","
    Dim rawBytes : rawBytes = ""
    If Not IsNull(rs("StatusBytes")) Then
        Dim b
        For b = 1 To LenB(rs("StatusBytes"))
            rawBytes = rawBytes & Hex(AscB(MidB(rs("StatusBytes"), b, 1)))
        Next
    End If
    json = json & "{""id"":" & rs("ProductionID") & ",""wo"":""" & rs("WorkOrderNo") & """,""status"":""" & Replace(rs("Status") & "", """", "\""") & """,""bytes"":""" & rawBytes & """}"
    first = False
    rs.MoveNext
Loop
json = json & "],"

rs.Close

' 尝试用 COLLATE 查找所有不同的Status值
json = json & """distinctStatuses"":["
Set rs = conn.Execute("SELECT DISTINCT Status FROM ProductionOrders ORDER BY Status")
first = True
Do While Not rs.EOF
    If Not first Then json = json & ","
    json = json & """" & Replace(rs("Status") & "", """", "\""") & """"
    first = False
    rs.MoveNext
Loop
json = json & "],"
rs.Close

' 尝试直接修复 - 使用 LIKE 模糊匹配
Dim fixCount : fixCount = 0
conn.Execute "UPDATE ProductionOrders SET Status='Pending' WHERE Status LIKE N'%待%' AND Status <> 'Pending'"
If Err.Number = 0 Then fixCount = fixCount + 1 : Err.Clear
conn.Execute "UPDATE ProductionOrders SET Status='InProgress' WHERE Status LIKE N'%生%' AND Status <> 'InProgress'"
If Err.Number = 0 Then fixCount = fixCount + 1 : Err.Clear
conn.Execute "UPDATE ProductionOrders SET Status='Completed' WHERE Status LIKE N'%完%' AND Status <> 'Completed'"
If Err.Number = 0 Then fixCount = fixCount + 1 : Err.Clear
conn.Execute "UPDATE ProductionOrders SET Status='Cancelled' WHERE Status LIKE N'%消%' AND Status <> 'Cancelled'"
If Err.Number = 0 Then fixCount = fixCount + 1 : Err.Clear
conn.Execute "UPDATE ProductionOrders SET Status='QC_Review' WHERE Status LIKE N'%检%' AND Status <> 'QC_Review'"
If Err.Number = 0 Then fixCount = fixCount + 1 : Err.Clear

json = json & """likeFixAttempts"":" & fixCount & ","

' 最终统计
json = json & """finalStatuses"":{"
Set rs = conn.Execute("SELECT Status, COUNT(*) AS Cnt FROM ProductionOrders GROUP BY Status ORDER BY Status")
first = True
Do While Not rs.EOF
    If Not first Then json = json & ","
    json = json & """" & Replace(rs("Status") & "", """", "\""") & """:" & rs("Cnt")
    first = False
    rs.MoveNext
Loop
rs.Close
json = json & "}"

json = json & "}"
Response.Write json
%>
