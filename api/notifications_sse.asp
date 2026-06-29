<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V16.0 SSE实时通知端点 (Server-Sent Events)
' 推送: 订单状态变更、库存预警、促销活动
' 用法: new EventSource("/api/notifications_sse.asp")
' ============================================
Response.ContentType = "text/event-stream"
Response.Charset = "UTF-8"
Response.CacheControl = "no-cache"
Response.AddHeader "Connection", "keep-alive"
Response.AddHeader "X-Accel-Buffering", "no"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
' ---- V18: Push订阅管理（POST请求）----
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Call OpenConnection()
	Dim pushAction, pushEndpoint, pushP256dh, pushAuth, pushUserId
	pushAction = Request.Form("push_action")
	pushEndpoint = Request.Form("endpoint")
	pushP256dh = Request.Form("p256dh")
	pushAuth = Request.Form("auth")
	pushUserId = Session("UserID")

	If pushUserId = "" Or IsNull(pushUserId) Then
		Response.ContentType = "application/json"
		Response.Write """error"":""请先登录"""
		Call CloseConnection()
		Response.End
	End If

	If pushAction = "subscribe" And pushEndpoint <> "" Then
		' 删除旧订阅（同一用户+端点）
		DAL_Execute "DELETE FROM PushSubscriptions WHERE UserID = @UID AND Endpoint = @EP", _
			Array(Array("@UID", DAL_adInteger, 0, CLng(pushUserId)), Array("@EP", DAL_adVarChar, 500, pushEndpoint))
		' 插入新订阅
		DAL_Execute "INSERT INTO PushSubscriptions (UserID, Endpoint, P256dh, AuthKey) VALUES (@UID, @EP, @P256, @Auth)", _
			Array(Array("@UID", DAL_adInteger, 0, CLng(pushUserId)), _
			      Array("@EP", DAL_adVarChar, 500, pushEndpoint), _
			      Array("@P256", DAL_adVarChar, 200, pushP256dh), _
			      Array("@Auth", DAL_adVarChar, 100, pushAuth))
		Response.ContentType = "application/json"
		Response.Write """success"":true"
	ElseIf pushAction = "unsubscribe" And pushEndpoint <> "" Then
		DAL_Execute "DELETE FROM PushSubscriptions WHERE UserID = @UID AND Endpoint = @EP", _
			Array(Array("@UID", DAL_adInteger, 0, CLng(pushUserId)), Array("@EP", DAL_adVarChar, 500, pushEndpoint))
		Response.ContentType = "application/json"
		Response.Write """success"":true"
	Else
		Response.ContentType = "application/json"
		Response.Write """error"":""无效的请求"""
	End If

	Call CloseConnection()
	Response.End
End If

' ---- 以下为SSE流式推送（GET请求）----
Server.ScriptTimeout = 600  ' SSE长连接需要更长超时
Call OpenConnection()

Dim userId, adminId, lastEventId
userId = Session("UserID")
adminId = Session("AdminID")

' 未登录用户不推送
If userId = "" And adminId = "" Then
    Response.Write "event: error" & vbCrLf
    Response.Write "data: " & API_JsonEncode("请先登录") & vbCrLf & vbCrLf
    Response.Flush
    Response.End
End If

' 获取最后事件ID（用于断线重连）
lastEventId = Request.ServerVariables("HTTP_LAST_EVENT_ID")
If lastEventId = "" Then lastEventId = "0"

' 心跳和事件循环
Dim loopCount, hasEvents, rs, eventId
loopCount = 0
eventId = CLng(lastEventId)

Do While loopCount < 30  ' 最多保持30个心跳周期（约5分钟）
    hasEvents = False
    
    ' 1. 检查订单状态变更通知
    If userId <> "" Then
        Set rs = conn.Execute("SELECT TOP 3 o.OrderID, o.OrderNo, o.Status, o.UpdatedAt " & _
            "FROM Orders o WHERE o.UserID = " & CLng(userId) & " AND o.UpdatedAt > DATEADD(MINUTE, -5, GETDATE()) " & _
            "ORDER BY o.UpdatedAt DESC")
        If Not rs Is Nothing Then
            Do While Not rs.EOF
                eventId = eventId + 1
                Response.Write "id: " & eventId & vbCrLf
                Response.Write "event: order_update" & vbCrLf
                Response.Write "data: {""orderId"":" & rs("OrderID") & ",""orderNo"":""" & rs("OrderNo") & """,""status"":""" & rs("Status") & """}" & vbCrLf & vbCrLf
                hasEvents = True
                rs.MoveNext
            Loop
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    ' 2. 库存预警（管理员）
    If adminId <> "" Then
        Set rs = conn.Execute("SELECT TOP 3 ni.NoteID, fn.NoteName, ni.StockQuantity, ni.AlertThreshold " & _
            "FROM NoteInventory ni INNER JOIN FragranceNotes fn ON ni.NoteID = fn.NoteID " & _
            "WHERE ni.StockQuantity <= ni.AlertThreshold AND ni.AlertThreshold > 0")
        If Not rs Is Nothing Then
            Do While Not rs.EOF
                eventId = eventId + 1
                Response.Write "id: " & eventId & vbCrLf
                Response.Write "event: inventory_alert" & vbCrLf
                Response.Write "data: {""noteId"":" & rs("NoteID") & ",""noteName"":""" & rs("NoteName") & """,""stock"":" & rs("StockQuantity") & ",""threshold"":" & rs("AlertThreshold") & "}" & vbCrLf & vbCrLf
                hasEvents = True
                rs.MoveNext
            Loop
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    ' 3. 心跳（每10秒）
    If Not hasEvents Or loopCount Mod 6 = 0 Then
        Response.Write ": heartbeat " & Now() & vbCrLf & vbCrLf
    End If
    
    Response.Flush
    
    ' 检查客户端是否断开
    If Not Response.IsClientConnected Then Exit Do
    
    ' 等待10秒再检查
    Dim waitUntil
    waitUntil = DateAdd("s", 10, Now())
    Do While Now() < waitUntil
        ' 短暂休眠
        Dim sleepEnd: sleepEnd = DateAdd("s", 1, Now())
        Do While Now() < sleepEnd
            ' busy-wait (ASP限制)
        Loop
        If Not Response.IsClientConnected Then Exit Do
    Loop
    
    loopCount = loopCount + 1
Loop

Call CloseConnection()
%>
