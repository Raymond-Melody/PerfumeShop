<!--#include file="../../../../includes/config.asp"-->
<!--#include file="../../../../includes/connection.asp"-->
<!--#include file="../../../includes/role_auth.asp"-->
<%
' ========== 品牌定香采购模块权限验证 ==========
' 复用采购中心权限体系
If Session("AdminID") = "" Or IsEmpty(Session("AdminID")) Then
    Response.Redirect "/admin/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If

' 加载角色信息
Call CheckRoleAndLoad()

' 确保 CSRF Token 已生成
Call EnsureCSRFToken()

' 确保 AdminName 已设置（后备：使用 AdminUsername）
If Session("AdminName") = "" Then
    Session("AdminName") = Session("AdminUsername")
End If

' 设置统一权限标识：isManager（SUPER_ADMIN 或 PURCHASE_MANAGER 为 True）
Dim isManager
isManager = False
If Session("AdminRoleCode") = "SUPER_ADMIN" Then
    isManager = True
ElseIf Session("AdminRoleCode") = "PURCHASE_MANAGER" Then
    isManager = True
End If
%>
