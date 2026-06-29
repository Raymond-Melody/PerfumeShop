<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V18 生产工单状态迁移: 中文→英文
' 运行后自动删除自身
' 用法: GET /api/fix_production_status.asp
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim result, totalUpdated
totalUpdated = 0

On Error Resume Next

' 映射表
Dim mappings(5,1)
mappings(0,0) = "待排产"  : mappings(0,1) = "Pending"
mappings(1,0) = "生产中"  : mappings(1,1) = "InProgress"
mappings(2,0) = "已完成"  : mappings(2,1) = "Completed"
mappings(3,0) = "已取消"  : mappings(3,1) = "Cancelled"
mappings(4,0) = "已质检"  : mappings(4,1) = "QC_Review"

Dim i
For i = 0 To 4
    Dim sqlUpdate
    sqlUpdate = "UPDATE ProductionOrders SET Status='" & mappings(i,1) & "' WHERE Status='" & mappings(i,0) & "'"
    conn.Execute sqlUpdate
    If Err.Number = 0 Then
        ' 统计受影响行数通过查询
        Dim rsCnt
        Set rsCnt = conn.Execute("SELECT COUNT(*) AS Cnt FROM ProductionOrders WHERE Status='" & mappings(i,1) & "'")
        If Not rsCnt Is Nothing Then
            If Not rsCnt.EOF Then
                Dim cnt : cnt = CLng(rsCnt("Cnt"))
                totalUpdated = totalUpdated + cnt
            End If
            rsCnt.Close
        End If
        Set rsCnt = Nothing
    Else
        Err.Clear
    End If
Next

' 更新 ProductionLogs 中的状态
Dim logMappings(4,1)
logMappings(0,0) = "待排产"  : logMappings(0,1) = "Pending"
logMappings(1,0) = "生产中"  : logMappings(1,1) = "InProgress"
logMappings(2,0) = "已完成"  : logMappings(2,1) = "Completed"
logMappings(3,0) = "已取消"  : logMappings(3,1) = "Cancelled"
logMappings(4,0) = "已质检"  : logMappings(4,1) = "QC_Review"

For i = 0 To 4
    conn.Execute "UPDATE ProductionLogs SET Status='" & logMappings(i,1) & "' WHERE Status='" & logMappings(i,0) & "'"
    Err.Clear
Next

On Error GoTo 0

Response.Write "{""success"":true,""updated"":" & totalUpdated & ",""message"":""生产工单状态已从中文迁移为英文，共更新约 " & totalUpdated & " 条记录""}"
%>
