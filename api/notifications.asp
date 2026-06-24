<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V15.0 实时通知 SSE 端点 (Server-Sent Events)
' 端点: /api/notifications.asp
' 用法: var es = new EventSource('/api/notifications.asp');
'       es.addEventListener('order_status', function(e) { ... });
' 注意: 需在 global.asa 或 IIS 中配置禁用响应缓冲
' ============================================
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/notification.asp"-->
<%
' 登录验证
Dim userId
If Session("UserID") <> "" And Not IsEmpty(Session("UserID")) Then
    userId = CLng(Session("UserID"))
ElseIf Session("AdminID") <> "" And Not IsEmpty(Session("AdminID")) Then
    userId = -CLng(Session("AdminID")) ' 负数表示管理员
Else
    Response.Status = "401 Unauthorized"
    Response.End
End If

' 获取模式：sse=长连接, poll=短轮询
Dim mode
mode = Request.QueryString("mode")
If mode = "" Then mode = "sse"

' 设置响应头
Response.ContentType = "text/event-stream"
Response.Charset = "UTF-8"
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "X-Accel-Buffering", "no"  ' 禁用nginx缓冲
Response.AddHeader "Connection", "keep-alive"

Dim lastCheckId
lastCheckId = Request.QueryString("lastId")
If lastCheckId = "" Then lastCheckId = 0

If mode = "poll" Then
    ' 短轮询模式（IE/不支持SSE的浏览器回退）
    Call NF_Init()
    Dim notifications, nf, json
    
    Set notifications = NF_GetForUser(userId, lastCheckId)
    
    Response.ContentType = "application/json"
    
    ' 构建JSON
    json = "{""notifications"":["
    If notifications.Count > 0 Then
        Dim keys, i, key
        keys = notifications.Keys()
        For i = 0 To notifications.Count - 1
            If i > 0 Then json = json & ","
            key = keys(i)
            Set nf = notifications.Item(key)
            json = json & "{""id"":" & nf("id") & _
                   ",""type"":""" & nf("type") & _
                   """,""message"":""" & Replace(nf("message"), """", "\""") & _
                   """,""link"":""" & nf("link") & _
                   """,""data"":" & nf("data") & "}"
        Next
    End If
    json = json & "],""lastCheckId"":" & lastCheckId & "}"
    
    Response.Write json
Else
    ' SSE长连接模式
    ' 发送初始连接确认
    Response.Write "event: connected" & vbCrLf
    Response.Write "data: {""status"":""connected"",""userId"":" & userId & "}" & vbCrLf
    Response.Write vbCrLf
    
    ' 发送心跳和轮询新通知
    Dim pollCount, maxPolls, hasNew
    pollCount = 0
    maxPolls = 120  ' 最长约10分钟（每5秒检查一次）
    
    Do While pollCount < maxPolls
        ' 检查是否有新通知
        Set notifications = NF_GetForUser(userId, lastCheckId)
        hasNew = False
        
        If notifications.Count > 0 Then
            Dim nKeys, j, nKey
            nKeys = notifications.Keys()
            For j = 0 To notifications.Count - 1
                nKey = nKeys(j)
                Set nf = notifications.Item(nKey)
                
                ' 发送SSE事件
                Response.Write NF_ToSSE(nf)
                hasNew = True
                
                If CLng(nf("id")) > CLng(lastCheckId) Then
                    lastCheckId = CLng(nf("id"))
                End If
            Next
            
            If hasNew Then
                Response.Flush() ' 立即推送
            End If
        End If
        
        Set notifications = Nothing
        
        ' 发送心跳（每15秒）
        If pollCount Mod 3 = 0 Then
            Response.Write ": heartbeat " & Now() & vbCrLf
            Response.Write vbCrLf
            Response.Flush()
        End If
        
        ' 检查客户端是否断开
        If Not Response.IsClientConnected Then
            Exit Do
        End If
        
        ' 等待5秒
        Dim waitStart : waitStart = Timer()
        Do While Timer() - waitStart < 5
            ' 空循环等待（ASP没有Sleep函数）
            Dim dummy : dummy = 0
            Do While dummy < 5000
                dummy = dummy + 1
            Loop
            If Not Response.IsClientConnected Then Exit Do
        Loop
        
        pollCount = pollCount + 1
    Loop
End If
%>