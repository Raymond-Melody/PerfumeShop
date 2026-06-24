<%
' ============================================
' V16.0 管理员操作审计日志 (Audit Log)
' 记录所有管理后台操作，满足合规要求
' ============================================

' 审计操作类型
Const AUDIT_ACTION_LOGIN = "login"
Const AUDIT_ACTION_LOGOUT = "logout"
Const AUDIT_ACTION_CREATE = "create"
Const AUDIT_ACTION_UPDATE = "update"
Const AUDIT_ACTION_DELETE = "delete"
Const AUDIT_ACTION_EXPORT = "export"
Const AUDIT_ACTION_BATCH = "batch_operation"
Const AUDIT_ACTION_VIEW = "view_sensitive"

' 审计目标类型
Const AUDIT_TARGET_ORDER = "order"
Const AUDIT_TARGET_PRODUCT = "product"
Const AUDIT_TARGET_USER = "user"
Const AUDIT_TARGET_COUPON = "coupon"
Const AUDIT_TARGET_SETTINGS = "settings"
Const AUDIT_TARGET_FINANCE = "finance"
Const AUDIT_TARGET_INVENTORY = "inventory"
Const AUDIT_TARGET_SYSTEM = "system"

' 确保审计日志表存在
Sub EnsureAuditLogTable()
    On Error Resume Next
    conn.Execute "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='AdminAuditLog') " & _
        "CREATE TABLE AdminAuditLog (" & _
        "LogID INT IDENTITY(1,1) PRIMARY KEY, " & _
        "AdminID INT NOT NULL, " & _
        "AdminName NVARCHAR(100), " & _
        "ActionType NVARCHAR(50) NOT NULL, " & _
        "TargetType NVARCHAR(50), " & _
        "TargetID INT, " & _
        "TargetName NVARCHAR(200), " & _
        "Details NVARCHAR(MAX), " & _
        "IPAddress NVARCHAR(50), " & _
        "UserAgent NVARCHAR(500), " & _
        "CreatedAt DATETIME DEFAULT GETDATE())"
    Err.Clear
    On Error GoTo 0
End Sub

' 写入审计日志
Sub AuditLog(actionType, targetType, targetID, targetName, details)
    On Error Resume Next
    
    Dim adminID, adminName, ipAddr, userAgent
    adminID = Session("AdminID")
    adminName = Session("AdminName")
    If adminName = "" Then adminName = "Admin#" & adminID
    
    ipAddr = Left(Request.ServerVariables("REMOTE_ADDR"), 50)
    userAgent = Left(Request.ServerVariables("HTTP_USER_AGENT"), 500)
    
    ' 构建SQL（使用参数化风格防止注入）
    Dim sql, targetIdValue
    If targetID <> "" Then
        targetIdValue = CLng(targetID)
    Else
        targetIdValue = "NULL"
    End If
    
    sql = "INSERT INTO AdminAuditLog (AdminID, AdminName, ActionType, TargetType, TargetID, TargetName, Details, IPAddress, UserAgent) " & _
          "VALUES (" & CLng(adminID) & ", " & _
          "'" & SafeSQL(adminName) & "', " & _
          "'" & SafeSQL(actionType) & "', " & _
          "'" & SafeSQL(targetType) & "', " & _
          targetIdValue & ", " & _
          "'" & SafeSQL(Left(targetName, 200)) & "', " & _
          "'" & SafeSQL(Left(details, 4000)) & "', " & _
          "'" & SafeSQL(ipAddr) & "', " & _
          "'" & SafeSQL(userAgent) & "')"
    
    conn.Execute sql
    
    Err.Clear
    On Error GoTo 0
End Sub

' 便捷函数：记录订单操作
Sub AuditOrder(actionType, orderID, orderNo, details)
    Call AuditLog(actionType, AUDIT_TARGET_ORDER, orderID, "Order#" & orderNo, details)
End Sub

' 便捷函数：记录产品操作
Sub AuditProduct(actionType, productID, productName, details)
    Call AuditLog(actionType, AUDIT_TARGET_PRODUCT, productID, productName, details)
End Sub

' 便捷函数：记录用户操作
Sub AuditUser(actionType, userID, userName, details)
    Call AuditLog(actionType, AUDIT_TARGET_USER, userID, userName, details)
End Sub

' 便捷函数：记录批量操作
Sub AuditBatch(actionName, totalCount, successCount, failCount, details)
    Dim batchDetail
    batchDetail = actionName & " | 总计:" & totalCount & " 成功:" & successCount & " 失败:" & failCount
    If details <> "" Then batchDetail = batchDetail & " | " & details
    Call AuditLog(AUDIT_ACTION_BATCH, AUDIT_TARGET_SYSTEM, 0, actionName, batchDetail)
End Sub

' 获取审计日志（分页）
Function GetAuditLogs(page, pageSize, actionFilter, dateFrom, dateTo)
    Dim sql, whereParts
    whereParts = "WHERE 1=1"
    
    If actionFilter <> "" Then
        whereParts = whereParts & " AND ActionType='" & SafeSQL(actionFilter) & "'"
    End If
    If dateFrom <> "" Then
        whereParts = whereParts & " AND CreatedAt>='" & SafeSQL(dateFrom) & "'"
    End If
    If dateTo <> "" Then
        whereParts = whereParts & " AND CreatedAt<='" & SafeSQL(dateTo) & " 23:59:59'"
    End If
    
    Dim offset
    offset = (page - 1) * pageSize
    
    sql = "SELECT LogID, AdminID, AdminName, ActionType, TargetType, TargetID, TargetName, " & _
          "Details, IPAddress, CreatedAt " & _
          "FROM AdminAuditLog " & whereParts & _
          " ORDER BY CreatedAt DESC " & _
          "OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    
    Set GetAuditLogs = conn.Execute(sql)
End Function

' 获取审计日志总数
Function GetAuditLogCount(actionFilter, dateFrom, dateTo)
    Dim sql, whereParts
    whereParts = "WHERE 1=1"
    
    If actionFilter <> "" Then
        whereParts = whereParts & " AND ActionType='" & SafeSQL(actionFilter) & "'"
    End If
    If dateFrom <> "" Then
        whereParts = whereParts & " AND CreatedAt>='" & SafeSQL(dateFrom) & "'"
    End If
    If dateTo <> "" Then
        whereParts = whereParts & " AND CreatedAt<='" & SafeSQL(dateTo) & " 23:59:59'"
    End If
    
    sql = "SELECT COUNT(*) FROM AdminAuditLog " & whereParts
    Dim rs, count
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        If Not rs.EOF Then count = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    GetAuditLogCount = count
End Function
%>
