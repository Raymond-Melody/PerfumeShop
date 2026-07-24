<!--#include file="auth.asp"-->
<%
' ============================================
' V7 模块化权限认证组件
' 功能: 角色验证、权限检查、操作日志记录
' 
' 用法: 在页面开头包含此文件前，必须先包含:
'   <!--#include file="../includes/config.asp"-->
'   <!--#include file="../includes/connection.asp"-->
' ============================================

' 检查并加载角色信息
Sub CheckRoleAndLoad()
    ' 如果还没有加载角色信息
    If Session("AdminRoleID") = "" Then
        Call LoadAdminRoleInfo()
    End If
End Sub

' 加载管理员角色信息
Sub LoadAdminRoleInfo()
    On Error Resume Next
    
    Dim adminId
    adminId = Session("AdminID")
    
    If adminId = "" Then
        Exit Sub
    End If
    
    Call OpenConnection()
    
    Dim rsRole
    Set rsRole = ExecuteQuery(_
        "SELECT u.RoleID, r.RoleName, r.RoleCode, r.Permissions, " & _
        "u.FullName AS RealName, u.Department, u.IsLocked " & _
        "FROM AdminUsers u " & _
        "LEFT JOIN AdminRoles r ON u.RoleID = r.RoleID " & _
        "WHERE u.AdminID = " & CLng(adminId))
    
    If Not rsRole Is Nothing And Not rsRole.EOF Then
        Session("AdminRoleID") = rsRole("RoleID")
        Session("AdminRoleName") = rsRole("RoleName")
        Session("AdminRoleCode") = rsRole("RoleCode")
        Session("AdminPermissions") = rsRole("Permissions")
        Session("AdminRealName") = rsRole("RealName")
        Session("AdminName") = rsRole("RealName")  ' 供采购模块等使用
        Session("AdminDepartment") = rsRole("Department")
        Session("AdminIsLocked") = rsRole("IsLocked")
        
        ' 检查账户是否被锁定
        Dim isLockedVal
        isLockedVal = rsRole("IsLocked")
        If IsNull(isLockedVal) Then isLockedVal = False
        If CBool(isLockedVal) = True Then
            Session.Abandon()
            Response.Redirect "/admin/login.asp?error=account_locked"
            Response.End
        End If
        
        rsRole.Close
    End If
    Set rsRole = Nothing
    
    Call CloseConnection()
    On Error GoTo 0
End Sub

' 检查是否有权访问指定模块
Function HasModulePermission(moduleCode, requiredLevel)
    HasModulePermission = False
    
    ' 超级管理员拥有所有权限
    If UCase(Session("AdminRoleCode")) = "SUPER_ADMIN" Or Session("AdminPermissions") = "all" Then
        HasModulePermission = True
        Exit Function
    End If
    
    ' ============================================
    ' V21 P3: 权限表驱动——优先查 ModulePermissions.RequiredRole
    ' 若该模块配置了 RequiredRole，则以表为准（精确或前缀匹配）；
    ' 未配置则回退到下方基于角色代码前缀的判断（向后兼容）
    ' ============================================
    Dim mpFound, mpRole, mpRs, mpClose
    mpFound = False : mpRole = "" : mpClose = False
    On Error Resume Next
    If Not IsObject(conn) Then
        Call OpenConnection() : mpClose = True
    ElseIf conn Is Nothing Then
        Call OpenConnection() : mpClose = True
    ElseIf conn.State <> 1 Then
        Call OpenConnection() : mpClose = True
    End If
    Set mpRs = conn.Execute("SELECT TOP 1 RequiredRole FROM ModulePermissions WHERE ModuleCode='" & Replace(moduleCode, "'", "''") & "' AND ISNULL(IsActive,1)=1 AND RequiredRole IS NOT NULL AND LEN(RequiredRole)>0")
    If Not mpRs Is Nothing Then
        If Not mpRs.EOF Then
            mpFound = True
            mpRole = CStr(mpRs("RequiredRole") & "")
        End If
        mpRs.Close
    End If
    Set mpRs = Nothing
    If mpClose Then Call CloseConnection()
    Err.Clear
    On Error GoTo 0

    If mpFound Then
        Dim curRoleCode : curRoleCode = Session("AdminRoleCode")
        If UCase(mpRole) = "SUPER_ADMIN" Then
            HasModulePermission = False   ' 仅超管可访问；非超管在上方已放行，此处即拒绝
        ElseIf curRoleCode = mpRole Then
            HasModulePermission = True
        ElseIf Len(mpRole) > 0 And Left(curRoleCode, Len(mpRole)) = mpRole Then
            HasModulePermission = True    ' 前缀匹配：RequiredRole='PROD' 命中 PROD_XXX
        Else
            HasModulePermission = False
        End If
        Exit Function
    End If
    
    ' 获取当前用户权限
    Dim permissions
    permissions = Session("AdminPermissions")
    
    If permissions = "" Then
        Exit Function
    End If
    
    ' 简单的权限检查（基于角色代码前缀）
    Dim roleCode
    roleCode = Session("AdminRoleCode")
    
    Select Case moduleCode
        Case "operation", "operation_orders", "operation_customers", "operation_points", "operation_recipes", "operation_marketing", "operation_products", "operation_fragrance", "operation_reviews", "operation_aftersales"
            ' 运营模块（含原内容管理功能）
            If Left(roleCode, 2) = "OP" Or roleCode = "CONTENT_ADMIN" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If
            
        Case "semifinished", "semifinished_accord", "semifinished_basenote", "semifinished_inventory", "semifinished_outbound"
            ' 半成品生产中心（V8新增）
            If Left(roleCode, 4) = "PROD" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If

        Case "prodcenter", "prodcenter_orders", "prodcenter_schedule", "prodcenter_qc", "prodcenter_warehouse", "prodcenter_inventory"
            ' 产品生产管理中心（V8新增）
            If Left(roleCode, 4) = "PROD" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If

        Case "logistics", "logistics_shipments", "logistics_tracking", "logistics_delivery", "logistics_returns"
            ' 物流管理中心（V8新增）
            If Left(roleCode, 4) = "PROD" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If

        Case "production", "production_inventory", "production_orders", "production_recipes", "production_logistics", "production_suppliers"
            ' 生产模块（旧版保留）
            If Left(roleCode, 4) = "PROD" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If
            
        Case "finance", "finance_revenue", "finance_cost", "finance_payments", "finance_bills", "finance_reports"
            ' 财务模块
            If Left(roleCode, 3) = "FIN" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If
            
        Case "system", "system_roles", "system_admins", "system_permissions", "system_logs", "system_settings", "system_security", "system_backup"
            ' 系统模块 - 只有超级管理员
            If roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If
            
        Case "techcenter", "techcenter_recipes", "techcenter_base_notes", "techcenter_notes", "techcenter_products", "techcenter_specs"
            ' 产品技术中心模块
            If Left(roleCode, 4) = "TECH" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If
            
        Case "purchase", "purchase_orders", "purchase_suppliers", "purchase_prices", "purchase_analysis"
            ' 采购管理中心模块
            If Left(roleCode, 8) = "PURCHASE" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If

        Case "inventory", "inventory_dashboard", "inventory_alerts", "inventory_movements"
            ' 库存管理中心（V8新增）
            If Left(roleCode, 4) = "PROD" Or roleCode = "SUPER_ADMIN" Then
                HasModulePermission = True
            End If
    End Select
End Function

' 验证模块访问权限（如果无权限则跳转）
Sub VerifyModuleAccess(moduleCode, requiredLevel)
    Call CheckRoleAndLoad()
    
    If Not HasModulePermission(moduleCode, requiredLevel) Then
        ' 记录越权访问尝试
        Call LogAdminAction("越权访问尝试", moduleCode, "", "", "")
        
        ' 跳转到无权限页面
        Response.Redirect "/admin/unauthorized.asp?module=" & moduleCode
        Response.End
    End If
End Sub

' ============================================
' V21: 操作级权限校验（RolePermissions 驱动）
' action: view/create/edit/delete/export/approve
' 规则：SUPER_ADMIN 或 Permissions='all' 豁免；必先具备模块访问权；
'       若 RolePermissions 未配置该(角色,模块)行，回退为"具模块访问权即允许"（向后兼容）
' ============================================
Function HasActionPermission(moduleCode, action)
    HasActionPermission = False
    Call CheckRoleAndLoad()

    If UCase(Session("AdminRoleCode")) = "SUPER_ADMIN" Or Session("AdminPermissions") = "all" Then
        HasActionPermission = True
        Exit Function
    End If

    If Not HasModulePermission(moduleCode, 1) Then
        HasActionPermission = False
        Exit Function
    End If

    Dim roleId, colName, rs, hasRow, allowed, needClose
    roleId = Session("AdminRoleID")
    If roleId = "" Then Exit Function

    Select Case LCase(action)
        Case "view":    colName = "CanView"
        Case "create":  colName = "CanCreate"
        Case "edit":    colName = "CanEdit"
        Case "delete":  colName = "CanDelete"
        Case "export":  colName = "CanExport"
        Case "approve": colName = "CanApprove"
        Case Else:      colName = "CanView"
    End Select

    hasRow = False
    allowed = False
    needClose = False
    On Error Resume Next
    If Not IsObject(conn) Then
        Call OpenConnection() : needClose = True
    ElseIf conn Is Nothing Then
        Call OpenConnection() : needClose = True
    ElseIf conn.State <> 1 Then
        Call OpenConnection() : needClose = True
    End If
    Set rs = conn.Execute("SELECT " & colName & " AS P FROM RolePermissions WHERE RoleID=" & CLng(roleId) & " AND ModuleCode='" & Replace(moduleCode, "'", "''") & "'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            hasRow = True
            Dim pv : pv = rs("P")
            allowed = (CStr(pv & "") = "True" Or CStr(pv & "") = "1")
        End If
        rs.Close
    End If
    Set rs = Nothing
    If needClose Then Call CloseConnection()
    Err.Clear
    On Error GoTo 0

    If Not hasRow Then
        HasActionPermission = True   ' 未配置操作级权限时向后兼容
    Else
        HasActionPermission = allowed
    End If
End Function

' V21: 无操作权限则拦截（用于关键写操作前）
Sub RequirePermissionOrDie(moduleCode, action)
    If Not HasActionPermission(moduleCode, action) Then
        Call LogAdminAction("越权操作拦截", moduleCode, "", "", "action=" & action)
        Response.Redirect "/admin/unauthorized.asp?module=" & moduleCode
        Response.End
    End If
End Sub

' 记录管理员操作
Sub LogAdminAction(action, module, targetTable, targetId, description)
    On Error Resume Next
    
    Dim adminId, roleId
    adminId = Session("AdminID")
    roleId = Session("AdminRoleID")
    
    If adminId = "" Then
        Exit Sub
    End If
    
    ' 检查是否已有打开的数据库连接（避免覆盖调用方的conn）
    Dim needCloseConn
    needCloseConn = False
    If Not IsObject(conn) Then
        Call OpenConnection()
        needCloseConn = True
    ElseIf conn Is Nothing Then
        Call OpenConnection()
        needCloseConn = True
    ElseIf conn.State <> 1 Then
        Call OpenConnection()
        needCloseConn = True
    End If
    
    Dim sql
    sql = "INSERT INTO AdminLogs (AdminID, ActionType, ModuleCode, TableName, RecordID, Notes, CreatedAt, IPAddress) VALUES (" & _
        CLng(adminId) & ", " & _
        "'" & SafeSQL(action) & "', " & _
        "'" & SafeSQL(module) & "', " & _
        "'" & SafeSQL(targetTable) & "', " & _
        "'" & SafeSQL(targetId) & "', " & _
        "'" & SafeSQL(description) & "', " & _
        "GETDATE(), " & _
        "'" & SafeSQL(Request.ServerVariables("REMOTE_ADDR")) & "')"
    
    ExecuteNonQuery(sql)
    
    ' 仅在本函数自行打开连接时才关闭，避免破坏调用方的连接
    If needCloseConn Then Call CloseConnection()
    On Error GoTo 0
End Sub

' 获取当前管理员可访问的后台类型
Function GetAccessiblePortals()
    Dim portals
    portals = ""
    
    Call CheckRoleAndLoad()
    
    Dim roleCode
    roleCode = Session("AdminRoleCode")
    
    If roleCode = "SUPER_ADMIN" Then
        portals = "operation,semifinished,prodcenter,production,logistics,purchase,finance,techcenter,system,inventory,analytics"
    ElseIf Left(roleCode, 2) = "OP" Then
        portals = "operation"
    ElseIf Left(roleCode, 4) = "PROD" Then
        portals = "production,semifinished,prodcenter,logistics,inventory"
    ElseIf Left(roleCode, 3) = "FIN" Then
        portals = "finance"
    ElseIf Left(roleCode, 4) = "TECH" Then
        portals = "techcenter"
    ElseIf Left(roleCode, 8) = "PURCHASE" Then
        portals = "purchase"
    ElseIf roleCode = "CONTENT_ADMIN" Then
        portals = "operation"
    End If
    
    GetAccessiblePortals = portals
End Function

' 获取第一个可访问的后台
Function GetDefaultPortal()
    Dim portals
    portals = GetAccessiblePortals()
    
    If portals <> "" Then
        GetDefaultPortal = Split(portals, ",")(0)
    Else
        GetDefaultPortal = ""
    End If
End Function

' 注意: SafeSQL函数已在connection.asp中定义，此处不需要重复定义

' 自动加载角色信息
Call CheckRoleAndLoad()
%>
