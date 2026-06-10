<!--#include file="../../includes/role_auth.asp"-->
<%
' 财务管理后台认证
Call VerifyModuleAccess("finance", 1)
Call LogAdminAction("访问财务后台", "finance", "", "", "")
%>
