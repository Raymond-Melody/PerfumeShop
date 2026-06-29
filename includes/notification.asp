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
' 内部辅助：序列化通知为管道分隔字符串
' 格式: id|type|targetId|message|link|data|createdAt|isRead
' ============================================
Function NF_Serialize(nf)
    NF_Serialize = nf("id") & "|" & nf("type") & "|" & nf("targetId") & "|" & _
                   Replace(nf("message"), "|", "&#124;") & "|" & _
                   Replace(nf("link"), "|", "&#124;") & "|" & _
                   Replace(nf("data"), "|", "&#124;") & "|" & _
                   nf("createdAt") & "|" & nf("isRead")
End Function

' ============================================
' 内部辅助：从管道分隔字符串反序列化为Dictionary
' ============================================
Function NF_Deserialize(str)
    Dim parts, nf
    parts = Split(str, "|")
    If UBound(parts) < 7 Then Set NF_Deserialize = Nothing : Exit Function
    Set nf = Server.CreateObject("Scripting.Dictionary")
    nf.Add "id", CLng(parts(0))
    nf.Add "type", parts(1)
    nf.Add "targetId", CLng(parts(2))
    nf.Add "message", Replace(parts(3), "&#124;", "|")
    nf.Add "link", Replace(parts(4), "&#124;", "|")
    nf.Add "data", Replace(parts(5), "&#124;", "|")
    nf.Add "createdAt", CDbl(parts(6))
    If UBound(parts) >= 7 Then
        nf.Add "isRead", CBool(parts(7))
    Else
        nf.Add "isRead", False
    End If
    Set NF_Deserialize = nf
End Function

' ============================================
' 内部辅助：注册/注销通知ID到键列表
' ============================================
Sub NF_RegisterKey(nfId)
    Dim keyList
    keyList = Application("NF_KeyList")
    If IsEmpty(keyList) Or keyList = "" Then
        Application("NF_KeyList") = CStr(nfId)
    Else
        Application("NF_KeyList") = keyList & "|" & CStr(nfId)
    End If
End Sub

Function NF_GetKeyArray()
    Dim keyList, keys
    keyList = Application("NF_KeyList")
    If IsEmpty(keyList) Or keyList = "" Then
        NF_GetKeyArray = Array()
    Else
        NF_GetKeyArray = Split(keyList, "|")
    End If
End Function

' ============================================
' 内部辅助：安全检查数组是否包含有效键
' VBScript的And/Or运算符不会短路，必须嵌套If
' ============================================
Function NF_HasKeys(keys)
    If Not IsArray(keys) Then
        NF_HasKeys = False
        Exit Function
    End If
    If UBound(keys) < 0 Then
        NF_HasKeys = False
        Exit Function
    End If
    If keys(0) = "" Then
        NF_HasKeys = False
        Exit Function
    End If
    NF_HasKeys = True
End Function

' ============================================
' 初始化通知队列（使用Application字符串变量）
' ============================================
Sub NF_Init()
    If IsEmpty(Application("NF_NextID")) Then
        Application.Lock
        If IsEmpty(Application("NF_NextID")) Then
            Application("NF_NextID") = 1
            Application("NF_LastCleanup") = CDbl(Now())
            Application("NF_KeyList") = ""
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
    Dim nf, nfId, nowTs, serialized
    
    Call NF_Init()
    
    nfId = NF_NextID()
    nowTs = CDbl(Now())
    
    ' 构建通知对象（内存中Dictionary，不存入Application）
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
    
    ' 序列化为字符串存入Application
    serialized = NF_Serialize(nf)
    
    Application.Lock
    
    Application("NF_" & nfId) = serialized
    Call NF_RegisterKey(nfId)
    
    ' 超过最大数量时清理旧消息
    If NF_GetKeyCount() > NF_MAX_TOTAL Then
        NF_CleanupInternal()
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
' 获取键列表数量
' ============================================
Function NF_GetKeyCount()
    Dim keys
    keys = NF_GetKeyArray()
    If NF_HasKeys(keys) Then
        NF_GetKeyCount = UBound(keys) + 1
    Else
        NF_GetKeyCount = 0
    End If
End Function

' ============================================
' 获取用户的通知（自lastCheckId之后的新通知）
' ============================================
Function NF_GetForUser(userId, lastCheckId)
    Dim keys, key, nf, result, i, count, serialized
    
    Call NF_Init()
    
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    If IsNull(lastCheckId) Or lastCheckId = "" Then lastCheckId = 0
    lastCheckId = CLng(lastCheckId)
    
    count = 0
    keys = NF_GetKeyArray()
    
    If NF_HasKeys(keys) Then
        For i = 0 To UBound(keys)
            key = keys(i)
            If key <> "" Then
                serialized = Application("NF_" & key)
                If Not IsEmpty(serialized) And serialized <> "" Then
                    Set nf = NF_Deserialize(serialized)
                    If Not nf Is Nothing Then
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
                End If
            End If
        Next
    End If
    
    Set NF_GetForUser = result
End Function

' ============================================
' 获取未读通知数量
' ============================================
Function NF_GetUnreadCount(userId)
    Dim keys, key, nf, count, i, serialized
    
    Call NF_Init()
    
    count = 0
    keys = NF_GetKeyArray()
    
    If NF_HasKeys(keys) Then
        For i = 0 To UBound(keys)
            key = keys(i)
            If key <> "" Then
                serialized = Application("NF_" & key)
                If Not IsEmpty(serialized) And serialized <> "" Then
                    Set nf = NF_Deserialize(serialized)
                    If Not nf Is Nothing Then
                        Dim unreadTarget
                        unreadTarget = CLng(nf("targetId"))
                        If unreadTarget = 0 Or unreadTarget = -1 Or unreadTarget = CLng(userId) Then
                            If Not nf("isRead") Then
                                count = count + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    NF_GetUnreadCount = count
End Function

' ============================================
' 标记通知为已读
' ============================================
Sub NF_MarkRead(notificationId)
    Dim nfKey, serialized, nf
    nfKey = "NF_" & CStr(notificationId)
    serialized = Application(nfKey)
    If Not IsEmpty(serialized) And serialized <> "" Then
        Set nf = NF_Deserialize(serialized)
        If Not nf Is Nothing Then
            nf("isRead") = True
            Application.Lock
            Application(nfKey) = NF_Serialize(nf)
            Application.UnLock
        End If
    End If
End Sub

' ============================================
' 获取最近的通知（管理端用，不限用户）
' ============================================
Function NF_GetRecent(count)
    Dim keys, key, nf, result, i, total, startIdx, serialized
    
    If IsNull(count) Or count < 1 Then count = 20
    
    Call NF_Init()
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    keys = NF_GetKeyArray()
    If Not NF_HasKeys(keys) Then
        Set NF_GetRecent = result
        Exit Function
    End If
    
    total = UBound(keys) + 1
    
    ' 取最后count条
    startIdx = total - count
    If startIdx < 0 Then startIdx = 0
    
    For i = startIdx To total - 1
        key = keys(i)
        If key <> "" Then
            serialized = Application("NF_" & key)
            If Not IsEmpty(serialized) And serialized <> "" Then
                Set nf = NF_Deserialize(serialized)
                If Not nf Is Nothing Then
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
            End If
        End If
    Next
    
    Set NF_GetRecent = result
End Function

' ============================================
' 内部清理函数（需持有Application锁）
' ============================================
Sub NF_CleanupInternal()
    Dim keys, key, nf, nowTs, i, serialized
    nowTs = CDbl(Now())
    keys = NF_GetKeyArray()
    
    If Not NF_HasKeys(keys) Then Exit Sub
    
    ' 删除过期消息
    Dim newKeyList : newKeyList = ""
    For i = 0 To UBound(keys)
        key = keys(i)
        If key <> "" Then
            serialized = Application("NF_" & key)
            If Not IsEmpty(serialized) And serialized <> "" Then
                Set nf = NF_Deserialize(serialized)
                If Not nf Is Nothing Then
                    Dim age : age = (nowTs - CDbl(nf("createdAt"))) * 86400
                    If age > NF_RETENTION_SECONDS Then
                        Application.Contents.Remove "NF_" & key
                    Else
                        If newKeyList = "" Then
                            newKeyList = key
                        Else
                            newKeyList = newKeyList & "|" & key
                        End If
                    End If
                End If
            Else
                ' 已损坏的数据，删除
                Application.Contents.Remove "NF_" & key
            End If
        End If
    Next
    Application("NF_KeyList") = newKeyList
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