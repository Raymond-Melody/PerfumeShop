<!--#include file="../../includes/role_auth.asp"-->
<%
' 产品技术管理中心认证
Call VerifyModuleAccess("techcenter", 1)
Call LogAdminAction("访问产品技术管理中心", "techcenter", "", "", "")

' 设置 isManager 变量（SUPER_ADMIN 和 TECH_MANAGER 为 True）
Dim isManager
isManager = False
If Session("AdminRoleCode") = "SUPER_ADMIN" Then
    isManager = True
ElseIf Session("AdminRoleCode") = "TECH_MANAGER" Then
    isManager = True
End If
%>
