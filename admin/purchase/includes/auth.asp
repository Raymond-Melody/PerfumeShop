<!--#include file="../../../includes/connection.asp"-->
<!--#include file="../../includes/role_auth.asp"-->
<%
' 采购管理中心认证
Call VerifyModuleAccess("purchase", 1)
Call LogAdminAction("访问采购后台", "purchase", "", "", "")

' 设置isManager变量（SUPER_ADMIN和PURCHASE_MANAGER为True）
Dim isManager
isManager = False
If Session("AdminRoleCode") = "SUPER_ADMIN" Then
    isManager = True
ElseIf Session("AdminRoleCode") = "PURCHASE_MANAGER" Then
    isManager = True
End If

' 全局IIf函数（用于在ASP表达式中做简化条件判断）
Function IIf(condition, trueVal, falseVal)
    If condition Then IIf = trueVal Else IIf = falseVal
End Function
%>
