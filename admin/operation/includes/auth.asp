<!--#include file="../../includes/role_auth.asp"-->
<%
' ============================================
' 运营管理后台认证
' ============================================

' 包含基础权限认证

' 验证运营后台访问权限
Call VerifyModuleAccess("operation", 1)

' 记录访问日志
Call LogAdminAction("访问运营后台", "operation", "", "", "")
%>
