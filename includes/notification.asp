<%
' ============================================
' V15.0 实时通知系统 (Notification System)
' 依赖: 无（独立模块，使用Application存储）
' 用法: <!--#include file="notification.asp"-->
' 调用: NF_Push "order_status", userId, "您的订单已发货", "orders.asp?id=123"
'        NF_Push "inventory_alert", 0, "SKU1001库存低于安全阈值", "admin/inventory/"
'        NF_GetForUser userId, lastCheckId  ' 获取用户未读通知
' ============================================

' 通知配置
Const NF_MAX_QUEUE = 500          ' 每个通道最大消息数
Const NF_MAX_TOTAL = 2000         ' 全局最大消息数
Const NF_RETENTION_SECONDS = 3600 ' 消息保留时间（1小时）

' 通知类型
Const NF_TYPE_ORDER = "order_status"
Const NF_TYPE_INVENTORY = "inventory_alert"
Const NF_TYPE_NEW_ORDER = "new_order"
Const NF_TYPE_ANNOUNCEMENT = "announcement"
Const NF_TYPE_PROMOTION = "promotion"
Const NF_TYPE_SYSTEM = "system"

' ============================================
' 初始化通知队列
' ============================================
Sub NF_Init()
    If Not IsObject(Application("NF_Queue")) Then
        Application.Lock
        If Not IsObject(Application("NF_Queue")) Then
            Set Application("NF_Queue") = Server.CreateObject("Scripting.Dictionary")
            Application("NF_NextID") = 1
            Application("NF_LastCleanup") = CDbl(Now())
        End If
        Application.UnLock
    End If
End Sub

' ============================================
' 生成唯一通知ID
' ============================================
Function NF_NextID()
    Dim nextId
    Application.Lock
    nextId = CLng(Application("NF_NextID"))
    Application("NF_NextID") = nextId + 1
    Application.UnLock
    NF_NextID = nextId
End Function

' ============================================
' 推送通知
' 参数:
'   nfType   - 通知类型（订单/库存/公告等）
'   targetId - 目标用户ID（0=广播给所有管理员）
'   message  - 通知消息文本
'   link     - 点击跳转链接（可选）
'   data     - 附加数据JSON（可选）
' ============================================
Sub NF_Push(nfType, targetId, message, link, data)
    Dim queue, nf, nfId, nowTs
    
    Call NF_Init()
    
    nfId = NF_NextID()
    nowTs = CDbl(Now())
    
    ' 构建通知对象
    Set nf = Server.CreateObject("Scripting.Dictionary")
    nf.Add "id", nfId
    nf.Add "type", nfType
    nf.Add "targetId", CLng(targetId)
    nf.Add "message", message
    If IsNull(link) Or IsEmpty(link) Then link = ""
    nf.Add "link", link
    If IsNull(data) Or IsEmpty(data) Then data = ""
    nf.Add "data", CStr(data)
    nf.Add "createdAt", nowTs
    nf.Add "isRead", False
    
    ' 存入Application队列
    Application.Lock
    
    Set queue = Application("NF_Queue")
    queue.Add CStr(nfId), nf
    
    ' 超过最大数量时清理旧消息
    If queue.Count > NF_MAX_TOTAL Then
        NF_CleanupInternal(queue)
    End If
    
    Application.UnLock
End Sub

' ============================================
' 便捷推送函数
' ============================================
Sub NF_PushOrderStatus(userId, message, orderId)
    Call NF_Push(NF_TYPE_ORDER, CLng(userId), message, "user/order_detail.asp?id=" & orderId, "{""orderId"":" & orderId & "}")
End Sub

Sub NF_PushInventoryAlert(message, link)
    Call NF_Push(NF_TYPE_INVENTORY, 0, message, link, "")
End Sub

Sub NF_PushNewOrder(adminId, message, orderId)
    Call NF_Push(NF_TYPE_NEW_ORDER, -1, message, "admin/operation/order_detail.asp?id=" & orderId, "{""orderId"":" & orderId & "}")
End Sub

Sub NF_PushAnnouncement(message, link)
    Call NF_Push(NF_TYPE_ANNOUNCEMENT, 0, message, link, "")
End Sub

' ============================================
' 获取用户的通知（自lastCheckId之后的新通知）
' ============================================
Function NF_GetForUser(userId, lastCheckId)
    Dim queue, keys, key, nf, result, i, count
    
    Call NF_Init()
    
    Set result = Server.CreateObject("Scripting.Dictionary")
    Set queue = Application("NF_Queue")
    
    If IsNull(lastCheckId) Or lastCheckId = "" Then lastCheckId = 0
    lastCheckId = CLng(lastCheckId)
    
    count = 0
    keys = queue.Keys()
    
    For Each key In keys
        Set nf = queue.Item(key)
        If IsObject(nf) Then
            ' 检查目标用户（targetId=0或-1为广播，匹配当前用户）
            Dim target
            target = CLng(nf("targetId"))
            If target = 0 Or target = -1 Or target = CLng(userId) Then
                If CLng(nf("id")) > lastCheckId Then
                    ' 添加到结果
                    Dim item
                    Set item = Server.CreateObject("Scripting.Dictionary")
                    item.Add "id", CLng(nf("id"))
                    item.Add "type", nf("type")
                    item.Add "message", nf("message")
                    item.Add "link", nf("link")
                    item.Add "data", nf("data")
                    item.Add "createdAt", nf("createdAt")
                    
                    result.Add CStr(count), item
                    count = count + 1
                    
                    If count >= 50 Then Exit For ' 最多返回50条
                End If
            End If
        End If
    Next
    
    Set NF_GetForUser = result
End Function

' ============================================
' 获取未读通知数量
' ============================================
Function NF_GetUnreadCount(userId)
    Dim queue, keys, key, nf, count
    
    Call NF_Init()
    Set queue = Application("NF_Queue")
    
    count = 0
    keys = queue.Keys()
    
    For Each key In keys
        Set nf = queue.Item(key)
        If IsObject(nf) Then
            Dim target
            target = CLng(nf("targetId"))
            If (target = 0 Or target = -1 Or target = CLng(userId)) And Not nf("isRead") Then
                count = count + 1
            End If
        End If
    Next
    
    NF_GetUnreadCount = count
End Function

' ============================================
' 标记通知为已读
' ============================================
Sub NF_MarkRead(notificationId)
    Dim queue, key
    Call NF_Init()
    Set queue = Application("NF_Queue")
    
    key = CStr(notificationId)
    If queue.Exists(key) Then
        queue.Item(key)("isRead") = True
    End If
End Sub

' ============================================
' 获取最近的通知（管理端用，不限用户）
' ============================================
Function NF_GetRecent(count)
    Dim queue, keys, key, nf, result, i, total, startIdx
    
    If IsNull(count) Or count < 1 Then count = 20
    
    Call NF_Init()
    Set queue = Application("NF_Queue")
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    keys = queue.Keys()
    total = queue.Count
    
    If total = 0 Then
        Set NF_GetRecent = result
        Exit Function
    End If
    
    ' 取最后count条
    startIdx = total - count
    If startIdx < 0 Then startIdx = 0
    
    For i = startIdx To total - 1
        key = keys(i)
        Set nf = queue.Item(key)
        If IsObject(nf) Then
            Dim item
            Set item = Server.CreateObject("Scripting.Dictionary")
            item.Add "id", CLng(nf("id"))
            item.Add "type", nf("type")
            item.Add "targetId", CLng(nf("targetId"))
            item.Add "message", nf("message")
            item.Add "link", nf("link")
            item.Add "isRead", nf("isRead")
            item.Add "createdAt", nf("createdAt")
            result.Add CStr(result.Count), item
        End If
    Next
    
    Set NF_GetRecent = result
End Function

' ============================================
' 内部清理函数（需持有Application锁）
' ============================================
Sub NF_CleanupInternal(queue)
    Dim keys, key, nf, nowTs, i
    nowTs = CDbl(Now())
    keys = queue.Keys()
    
    ' 删除过期消息
    For i = queue.Count - 1 To 0 Step -1
        key = keys(i)
        Set nf = queue.Item(key)
        If IsObject(nf) Then
            Dim age : age = (nowTs - CDbl(nf("createdAt"))) * 86400
            If age > NF_RETENTION_SECONDS Then
                queue.Remove key
            End If
        End If
    Next
End Sub

' ============================================
' 转换为SSE格式字符串
' ============================================
Function NF_ToSSE(notification)
    Dim sse
    sse = "id: " & notification("id") & vbCrLf
    sse = sse & "event: " & notification("type") & vbCrLf
    sse = sse & "data: {""id"":" & notification("id") & _
          ",""type"":""" & notification("type") & _
          """,""message"":""" & Replace(notification("message"), """", "\""") & _
          """,""link"":""" & notification("link") & """}" & vbCrLf
    sse = sse & vbCrLf
    NF_ToSSE = sse
End Function

' ============================================
' 初始化
' ============================================
Call NF_Init()
%>