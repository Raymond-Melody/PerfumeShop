<!--#include file="../../includes/role_auth.asp"-->
<%
' 系统管理后台认证
Call VerifyModuleAccess("system", 5)
Call LogAdminAction("访问系统后台", "system", "", "", "")
%>
