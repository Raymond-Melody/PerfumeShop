<%
' ============================================
' V16.0 通知事件触发器 (Notification Triggers)
' 在關鍵業務流程中調用，觸發 SSE 推送
' ============================================
Sub TriggerNotification(eventType, eventData)
    ' 將事件寫入通知隊列（數據庫表），SSE端點定期輪詢
    If Not FEATURE_SSE_NOTIFICATIONS Then Exit Sub
    
    On Error Resume Next
    Dim sql, userId
    userId = Session("UserID")
    If userId = "" Then userId = "NULL"
    
    sql = "INSERT INTO NotificationQueue (UserID, EventType, EventData, IsRead, CreatedAt) " & _
          "VALUES (" & userId & ", '" & SafeSQL(eventType) & "', '" & SafeSQL(eventData) & "', 0, GETDATE())"
    conn.Execute sql
    Err.Clear
    On Error GoTo 0
End Sub

' 確保通知隊列表存在
Sub EnsureNotificationTable()
    On Error Resume Next
    conn.Execute "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='NotificationQueue') " & _
        "CREATE TABLE NotificationQueue (" & _
        "QueueID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "UserID INT NULL, " & _
        "EventType NVARCHAR(50), " & _
        "EventData NVARCHAR(MAX), " & _
        "IsRead BIT DEFAULT 0, " & _
        "CreatedAt DATETIME DEFAULT GETDATE())"
    Err.Clear
    On Error GoTo 0
End Sub
%>
