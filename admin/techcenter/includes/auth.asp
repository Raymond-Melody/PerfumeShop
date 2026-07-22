<!--#include file="../../includes/role_auth.asp"-->
<%
' ============================================
' V18 产品技术管理中心认证组件
' 依赖: admin/includes/role_auth.asp (含主auth.asp)
' 提供: 模块访问控制、角色判定、CSRF保护、操作日志
' ============================================

' 1. 验证模块访问权限（含登录检查和角色验证）
Call VerifyModuleAccess("techcenter", 1)

' 2. 确保 Session 关键字段存在
If Session("AdminUsername") = "" Then
    Session("AdminUsername") = Session("AdminName")
End If
If Session("AdminUsername") = "" Then
    Session("AdminUsername") = "Unknown"
End If

' 3. CSRF Token 初始化（所有POST表单需携带）
If Session("CSRFToken") = "" Then
    Randomize
    Session("CSRFToken") = CStr(CLng(Rnd * 99999999) + 10000000) & CStr(Timer)
End If

' 4. CSP Nonce 初始化（用于内联脚本安全策略）
If Session("csp_nonce") = "" Then
    Randomize
    Session("csp_nonce") = CStr(CLng(Rnd * 99999999) + 10000000)
End If

' 5. 记录访问日志
Call LogAdminAction("访问产品技术管理中心", "techcenter", "", "", "")

' 6. 设置 isManager 变量（SUPER_ADMIN 和 TECH_MANAGER 为 True）
Dim isManager
isManager = False
If Session("AdminRoleCode") = "SUPER_ADMIN" Then
    isManager = True
ElseIf Session("AdminRoleCode") = "TECH_MANAGER" Then
    isManager = True
End If
%>
