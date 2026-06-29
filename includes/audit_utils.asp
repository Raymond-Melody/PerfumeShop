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

' V18: 隐私操作类型
Const AUDIT_ACTION_PRIVACY_EXPORT = "privacy_export"
Const AUDIT_ACTION_PRIVACY_DELETE = "privacy_delete"
Const AUDIT_ACTION_PRIVACY_CONSENT = "privacy_consent"
Const AUDIT_ACTION_PRIVACY_ACCESS = "privacy_access"
Const AUDIT_ACTION_PRIVACY_RECTIFY = "privacy_rectify"
Const AUDIT_TARGET_PRIVACY = "privacy"

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

' V17: 使用参数化查询防止SQL注入
Sub AuditLog(actionType, targetType, targetID, targetName, details)
    On Error Resume Next
    
    Dim adminID, adminName, ipAddr, userAgent, sql, params(), paramIdx, targetIdValue
    adminID = Session("AdminID")
    adminName = Session("AdminName")
    If adminName = "" Then adminName = "Admin#" & adminID
    
    ipAddr = Left(Request.ServerVariables("REMOTE_ADDR"), 50)
    userAgent = Left(Request.ServerVariables("HTTP_USER_AGENT"), 500)
    
    ' 使用参数化查询
    paramIdx = -1
    ReDim params(0)
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@AdminID", DAL_adInteger, 0, CLng(adminID))
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@AdminName", DAL_adVarChar, 100, Left(adminName, 100))
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@ActionType", DAL_adVarChar, 50, Left(actionType, 50))
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@TargetType", DAL_adVarChar, 50, Left(targetType, 50))
    
    If targetID <> "" AND IsNumeric(targetID) Then
        targetIdValue = CLng(targetID)
    Else
        targetIdValue = Null
    End If
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@TargetID", DAL_adInteger, 0, targetIdValue)
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@TargetName", DAL_adVarChar, 200, Left(targetName, 200))
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@Details", DAL_adVarChar, 4000, Left(details, 4000))
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@IPAddress", DAL_adVarChar, 50, ipAddr)
    
    paramIdx = paramIdx + 1
    ReDim Preserve params(paramIdx)
    params(paramIdx) = Array("@UserAgent", DAL_adVarChar, 500, userAgent)
    
    sql = "INSERT INTO AdminAuditLog (AdminID, AdminName, ActionType, TargetType, TargetID, TargetName, Details, IPAddress, UserAgent) " & _
          "VALUES (@AdminID, @AdminName, @ActionType, @TargetType, @TargetID, @TargetName, @Details, @IPAddress, @UserAgent)"
    
    DAL_Execute sql, params
    
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

' ============================================
' V18: LogPrivacyAction - GDPR 隐私操作审计
' 记录所有隐私相关操作（数据导出、账户注销、数据访问等）
' 参数:
'   actionType - 操作类型 (privacy_export/privacy_delete/privacy_consent/privacy_access/privacy_rectify)
'   userID     - 用户ID
'   userName   - 用户名
'   details    - 操作详情
' ============================================
Sub LogPrivacyAction(actionType, userID, userName, details)
    On Error Resume Next
    
    If Not FEATURE_GDPR_COMPLIANCE Then Exit Sub
    
    Dim ipAddr, sql, params(5)
    ipAddr = Left(Request.ServerVariables("REMOTE_ADDR"), 50)
    
    ' 使用 AppLogs 表记录隐私操作（与现有日志系统统一）
    params(0) = Array("@LogLevel", DAL_adVarChar, 10, "PRIVACY")
    params(1) = Array("@LogMessage", DAL_adVarChar, 2000, _
        actionType & " | UserID:" & userID & " | UserName:" & userName & " | " & details)
    params(2) = Array("@LogSource", DAL_adVarChar, 100, "privacy_action")
    params(3) = Array("@IPAddress", DAL_adVarChar, 50, ipAddr)
    params(4) = Array("@PageURL", DAL_adVarChar, 500, Left(Request.ServerVariables("SCRIPT_NAME"), 500))
    
    sql = "INSERT INTO AppLogs (LogLevel, LogMessage, LogSource, IPAddress, PageURL) " & _
          "VALUES (@LogLevel, @LogMessage, @LogSource, @IPAddress, @PageURL)"
    DAL_Execute sql, params
    
    ' 同时写入 AdminAuditLog（如果用户是管理员操作自己数据）
    If Session("AdminID") <> "" Then
        Call AuditLog(actionType, AUDIT_TARGET_PRIVACY, CLng(userID), userName, details)
    End If
    
    Err.Clear
    On Error GoTo 0
End Sub

' ============================================
' V18: LogCookieConsent - 记录 Cookie 同意
' ============================================
Sub LogCookieConsent(consentGiven, consentLevel)
    If Not FEATURE_GDPR_COMPLIANCE Then Exit Sub
    
    On Error Resume Next
    Dim userId, userName, details, ipAddr, sql, params(4)
    userId = "0"
    userName = "anonymous"
    If Session("UserID") <> "" Then
        userId = Session("UserID")
        userName = Session("Username")
    End If
    
    details = "CookieConsent: " & consentLevel & " | Accepted: " & LCase(CStr(consentGiven))
    ipAddr = Left(Request.ServerVariables("REMOTE_ADDR"), 50)
    
    params(0) = Array("@LogLevel", DAL_adVarChar, 10, "PRIVACY")
    params(1) = Array("@LogMessage", DAL_adVarChar, 2000, _
        "COOKIE_CONSENT | UserID:" & userId & " | " & details)
    params(2) = Array("@LogSource", DAL_adVarChar, 100, "cookie_consent")
    params(3) = Array("@IPAddress", DAL_adVarChar, 50, ipAddr)
    params(4) = Array("@PageURL", DAL_adVarChar, 500, Left(Request.ServerVariables("SCRIPT_NAME"), 500))
    
    sql = "INSERT INTO AppLogs (LogLevel, LogMessage, LogSource, IPAddress, PageURL) " & _
          "VALUES (@LogLevel, @LogMessage, @LogSource, @IPAddress, @PageURL)"
    DAL_Execute sql, params
    
    Err.Clear
    On Error GoTo 0
End Sub

' V17: 获取审计日志（分页）使用参数化查询
Function GetAuditLogs(page, pageSize, actionFilter, dateFrom, dateTo)
    Dim sql, params(), paramIdx
    
    sql = "SELECT LogID, AdminID, AdminName, ActionType, TargetType, TargetID, TargetName, " & _
          "Details, IPAddress, CreatedAt " & _
          "FROM AdminAuditLog WHERE 1=1"
    
    paramIdx = -1
    ReDim params(0)
    
    If actionFilter <> "" Then
        sql = sql & " AND ActionType=@ActionFilter"
        paramIdx = paramIdx + 1
        ReDim Preserve params(paramIdx)
        params(paramIdx) = Array("@ActionFilter", DAL_adVarChar, 50, actionFilter)
    End If
    If dateFrom <> "" Then
        sql = sql & " AND CreatedAt>=@DateFrom"
        paramIdx = paramIdx + 1
        ReDim Preserve params(paramIdx)
        params(paramIdx) = Array("@DateFrom", DAL_adVarChar, 20, dateFrom)
    End If
    If dateTo <> "" Then
        sql = sql & " AND CreatedAt<=@DateTo"
        paramIdx = paramIdx + 1
        ReDim Preserve params(paramIdx)
        params(paramIdx) = Array("@DateTo", DAL_adVarChar, 30, dateTo & " 23:59:59")
    End If
    
    Dim offset
    offset = (page - 1) * pageSize
    
    sql = sql & " ORDER BY CreatedAt DESC OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
    
    Set GetAuditLogs = DAL_GetList(sql, params)
End Function

' V17: 获取审计日志总数（参数化查询）
Function GetAuditLogCount(actionFilter, dateFrom, dateTo)
    Dim sql, params(), paramIdx, count
    
    sql = "SELECT COUNT(*) FROM AdminAuditLog WHERE 1=1"
    
    paramIdx = -1
    ReDim params(0)
    
    If actionFilter <> "" Then
        sql = sql & " AND ActionType=@ActionFilter"
        paramIdx = paramIdx + 1
        ReDim Preserve params(paramIdx)
        params(paramIdx) = Array("@ActionFilter", DAL_adVarChar, 50, actionFilter)
    End If
    If dateFrom <> "" Then
        sql = sql & " AND CreatedAt>=@DateFrom"
        paramIdx = paramIdx + 1
        ReDim Preserve params(paramIdx)
        params(paramIdx) = Array("@DateFrom", DAL_adVarChar, 20, dateFrom)
    End If
    If dateTo <> "" Then
        sql = sql & " AND CreatedAt<=@DateTo"
        paramIdx = paramIdx + 1
        ReDim Preserve params(paramIdx)
        params(paramIdx) = Array("@DateTo", DAL_adVarChar, 30, dateTo & " 23:59:59")
    End If
    
    count = CLng(DAL_GetScalar(sql, params, 0))
    GetAuditLogCount = count
End Function
%>
