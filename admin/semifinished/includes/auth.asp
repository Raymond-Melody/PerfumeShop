<!--#include file="../../includes/role_auth.asp"-->
<%
' 半成品生产中心认证
Call VerifyModuleAccess("semifinished", 1)
Call LogAdminAction("访问半成品生产中心", "semifinished", "", "", "")

' 检查当前用户是否拥有指定的半成品生产角色
Function CheckSemifinishedRole(allowedRoles)
    Dim currentRole, roles, i, allowed
    currentRole = Session("AdminRoleCode")
    
    If currentRole = "SUPER_ADMIN" Then
        CheckSemifinishedRole = True
        Exit Function
    End If
    If currentRole = "PROD_MANAGER" Then
        CheckSemifinishedRole = True
        Exit Function
    End If
    
    roles = Split(allowedRoles, ",")
    allowed = False
    For i = 0 To UBound(roles)
        If Trim(roles(i)) = currentRole Then
            allowed = True
            Exit For
        End If
    Next
    
    If Not allowed Then
        Response.Redirect "/admin/unauthorized.asp?module=semifinished"
        Response.End
    End If
    
    CheckSemifinishedRole = True
End Function

Function GetCurrentRoleCode()
    Dim rc
    rc = Session("AdminRoleCode")
    If IsNull(rc) Or rc = "" Then
        GetCurrentRoleCode = ""
    Else
        GetCurrentRoleCode = rc
    End If
End Function

Function IsManagerRole()
    Dim rc
    rc = Session("AdminRoleCode")
    IsManagerRole = (rc = "SUPER_ADMIN" Or rc = "PROD_MANAGER")
End Function
%>
