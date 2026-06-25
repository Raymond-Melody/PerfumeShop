<%
' ============================================
' V17.0 DAL - 管理后台数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_admin.asp"-->
' 涵盖：管理员认证、后台仪表盘、审计日志、后台管理查询
' ============================================

' ============================================
' 管理员 - 根据用户名获取
' ============================================
Function DAL_Admin_GetByUsername(username)
    Dim sql, params(0)
    sql = "SELECT AdminID, Username, PasswordHash, IsActive, RoleID FROM AdminUsers WHERE Username=@Username"
    params(0) = Array("@Username", DAL_adVarChar, 50, username)
    Set DAL_Admin_GetByUsername = DAL_GetRow(sql, params)
End Function

' ============================================
' 管理员 - 根据ID获取
' ============================================
Function DAL_Admin_GetByID(adminId)
    Dim sql, params(0)
    sql = "SELECT AdminID, Username, RoleID, IsActive, Email, LastLoginAt FROM AdminUsers WHERE AdminID=@AdminID"
    params(0) = Array("@AdminID", DAL_adInteger, 0, CLng(adminId))
    Set DAL_Admin_GetByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 管理员 - 验证登录（参数化查询）
' ============================================
Function DAL_Admin_ValidateLogin(username, passwordHash)
    Dim sql, params(1)
    sql = "SELECT AdminID, Username, RoleID, IsActive FROM AdminUsers " & _
          "WHERE Username=@Username AND PasswordHash=@PasswordHash AND IsActive=1"
    params(0) = Array("@Username", DAL_adVarChar, 50, username)
    params(1) = Array("@PasswordHash", DAL_adVarChar, 255, passwordHash)
    Set DAL_Admin_ValidateLogin = DAL_GetRow(sql, params)
End Function

' ============================================
' 管理员 - 更新最后登录时间
' ============================================
Sub DAL_Admin_UpdateLastLogin(adminId)
    Dim sql, params(0)
    sql = "UPDATE AdminUsers SET LastLoginAt=GETDATE() WHERE AdminID=@AdminID"
    params(0) = Array("@AdminID", DAL_adInteger, 0, CLng(adminId))
    DAL_Execute sql, params
End Sub

' ============================================
' 管理员 - 获取权限列表
' ============================================
Function DAL_Admin_GetPermissions(roleId)
    Dim sql, params(0)
    sql = "SELECT PermissionKey, PermissionValue FROM RolePermissions WHERE RoleID=@RoleID"
    params(0) = Array("@RoleID", DAL_adInteger, 0, CLng(roleId))
    Set DAL_Admin_GetPermissions = DAL_GetList(sql, params)
End Function

' ============================================
' 管理员 - 列表（分页）
' ============================================
Function DAL_Admin_GetList(page, pageSize, ByRef pageInfo)
    Dim sql
    sql = "SELECT AdminID, Username, Email, RoleID, IsActive, LastLoginAt, CreatedAt " & _
          "FROM AdminUsers ORDER BY AdminID ASC"
    Set DAL_Admin_GetList = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
End Function

' ============================================
' 仪表盘 - 关键统计数据
' ============================================
Sub DAL_Admin_GetDashboardStats(ByRef userCount, ByRef orderCount, ByRef revenueToday, ByRef productCount)
    userCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Users WHERE IsActive <> 0", Null, 0))
    orderCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Orders WHERE CAST(OrderDate AS DATE)=CAST(GETDATE() AS DATE)", Null, 0))
    revenueToday = CDbl(DAL_GetScalar("SELECT ISNULL(SUM(CAST(TotalAmount AS FLOAT)), 0) FROM Orders WHERE CAST(OrderDate AS DATE)=CAST(GETDATE() AS DATE)", Null, 0))
    productCount = CLng(DAL_GetScalar("SELECT COUNT(*) FROM Products WHERE IsActive=1", Null, 0))
End Sub

' ============================================
' 审计日志 - 写入
' ============================================
Sub DAL_Admin_WriteAuditLog(adminId, adminName, actionType, targetType, targetId, targetName, details, ipAddress, userAgent)
    Dim sql, params(8)
    sql = "INSERT INTO AdminAuditLog (AdminID, AdminName, ActionType, TargetType, " & _
          "TargetID, TargetName, Details, IPAddress, UserAgent) " & _
          "VALUES (@AdminID, @AdminName, @ActionType, @TargetType, " & _
          "@TargetID, @TargetName, @Details, @IPAddress, @UserAgent)"
    params(0) = Array("@AdminID", DAL_adInteger, 0, CLng(adminId))
    params(1) = Array("@AdminName", DAL_adVarChar, 100, Left(adminName, 100))
    params(2) = Array("@ActionType", DAL_adVarChar, 50, Left(actionType, 50))
    params(3) = Array("@TargetType", DAL_adVarChar, 50, Left(targetType, 50))
    params(4) = Array("@TargetID", DAL_adInteger, 0, CLng(targetId))
    params(5) = Array("@TargetName", DAL_adVarChar, 200, Left(targetName, 200))
    params(6) = Array("@Details", DAL_adVarChar, 4000, Left(details, 4000))
    params(7) = Array("@IPAddress", DAL_adVarChar, 50, Left(ipAddress, 50))
    params(8) = Array("@UserAgent", DAL_adVarChar, 500, Left(userAgent, 500))
    DAL_Execute sql, params
End Sub

' ============================================
' 审计日志 - 获取（分页）
' ============================================
Function DAL_Admin_GetAuditLogs(actionFilter, dateFrom, dateTo, page, pageSize, ByRef pageInfo)
    Dim sql, params(), paramCount
    
    sql = "SELECT LogID, AdminID, AdminName, ActionType, TargetType, TargetID, " & _
          "TargetName, Details, IPAddress, CreatedAt FROM AdminAuditLog WHERE 1=1"
    paramCount = -1
    ReDim params(0)
    
    If actionFilter <> "" Then
        sql = sql & " AND ActionType=@ActionFilter"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@ActionFilter", DAL_adVarChar, 50, actionFilter)
    End If
    
    If dateFrom <> "" Then
        sql = sql & " AND CreatedAt>=@DateFrom"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@DateFrom", DAL_adVarChar, 20, dateFrom)
    End If
    
    If dateTo <> "" Then
        sql = sql & " AND CreatedAt<=@DateTo"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@DateTo", DAL_adVarChar, 30, dateTo & " 23:59:59")
    End If
    
    sql = sql & " ORDER BY CreatedAt DESC"
    
    If paramCount >= 0 Then
        Set DAL_Admin_GetAuditLogs = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
    Else
        Set DAL_Admin_GetAuditLogs = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
    End If
End Function

' ============================================
' 审计日志 - 获取总数
' ============================================
Function DAL_Admin_CountAuditLogs(actionFilter, dateFrom, dateTo)
    Dim sql, params(), paramCount
    
    sql = "SELECT COUNT(*) FROM AdminAuditLog WHERE 1=1"
    paramCount = -1
    ReDim params(0)
    
    If actionFilter <> "" Then
        sql = sql & " AND ActionType=@ActionFilter"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@ActionFilter", DAL_adVarChar, 50, actionFilter)
    End If
    
    If dateFrom <> "" Then
        sql = sql & " AND CreatedAt>=@DateFrom"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@DateFrom", DAL_adVarChar, 20, dateFrom)
    End If
    
    If dateTo <> "" Then
        sql = sql & " AND CreatedAt<=@DateTo"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@DateTo", DAL_adVarChar, 30, dateTo & " 23:59:59")
    End If
    
    DAL_Admin_CountAuditLogs = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 站点设置 - 获取
' ============================================
Function DAL_Admin_GetSetting(settingKey)
    Dim sql, params(0)
    sql = "SELECT SettingValue FROM SiteSettings WHERE SettingKey=@SettingKey"
    params(0) = Array("@SettingKey", DAL_adVarChar, 100, settingKey)
    DAL_Admin_GetSetting = DAL_GetScalar(sql, params, "")
End Function

' ============================================
' 站点设置 - 更新
' ============================================
Function DAL_Admin_UpdateSetting(settingKey, settingValue)
    Dim sql, params(1)
    sql = "IF EXISTS (SELECT 1 FROM SiteSettings WHERE SettingKey=@SettingKey) " & _
          "UPDATE SiteSettings SET SettingValue=@SettingValue, UpdatedAt=GETDATE() WHERE SettingKey=@SettingKey " & _
          "ELSE INSERT INTO SiteSettings (SettingKey, SettingValue, CreatedAt) VALUES (@SettingKey, @SettingValue, GETDATE())"
    params(0) = Array("@SettingKey", DAL_adVarChar, 100, Left(settingKey, 100))
    params(1) = Array("@SettingValue", DAL_adVarChar, 500, Left(settingValue, 500))
    DAL_Admin_UpdateSetting = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 后台 - 最近注册用户
' ============================================
Function DAL_Admin_GetRecentUsers(limit)
    Dim sql
    sql = "SELECT TOP " & CLng(limit) & " UserID, Username, Email, FullName, " & _
          "CreatedAt, IsActive FROM Users ORDER BY CreatedAt DESC"
    Set DAL_Admin_GetRecentUsers = DAL_GetList(sql, Null)
End Function

' ============================================
' 后台 - 最近订单（简要）
' ============================================
Function DAL_Admin_GetRecentOrders(limit)
    Dim sql
    sql = "SELECT TOP " & CLng(limit) & " OrderID, OrderNo, UserID, TotalAmount, " & _
          "Status, OrderDate FROM Orders ORDER BY OrderDate DESC"
    Set DAL_Admin_GetRecentOrders = DAL_GetList(sql, Null)
End Function

' ============================================
' 角色权限 - 获取角色列表
' ============================================
Function DAL_Admin_GetRoles()
    Dim sql
    sql = "SELECT RoleID, RoleName, Description FROM AdminRoles WHERE IsActive=1 ORDER BY RoleID"
    Set DAL_Admin_GetRoles = DAL_GetList(sql, Null)
End Function

' ============================================
' 系统诊断 - 数据库表存在性检查
' ============================================
Function DAL_Admin_CheckTableExists(tableName)
    Dim sql, params(0)
    sql = "SELECT COUNT(*) FROM sys.tables WHERE name=@TableName"
    params(0) = Array("@TableName", DAL_adVarChar, 100, tableName)
    DAL_Admin_CheckTableExists = (CLng(DAL_GetScalar(sql, params, 0)) > 0)
End Function

' ============================================
' 系统诊断 - 获取数据库统计信息
' ============================================
Sub DAL_Admin_GetDBStats(ByRef dbSizeMB, ByRef tableCount, ByRef activeConnections)
    dbSizeMB = CDbl(DAL_GetScalar( _
        "SELECT CAST(SUM(size * 8.0 / 1024) AS DECIMAL(18,2)) FROM sys.database_files WHERE type=0", Null, 0))
    tableCount = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM sys.tables", Null, 0))
    activeConnections = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM sys.dm_exec_sessions WHERE is_user_process=1 AND DB_NAME(database_id)=DB_NAME()", Null, 0))
End Sub
%>
