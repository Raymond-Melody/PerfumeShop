<!--#include file="../../includes/role_auth.asp"-->
<%
' 物流管理中心认证
Call VerifyModuleAccess("logistics", 1)
Call LogAdminAction("访问物流管理中心", "logistics", "", "", "")

Function CheckLogisticsRole(allowedRoles)
    Dim currentRole, roles, i, allowed
    currentRole = Session("AdminRoleCode")
    
    If currentRole = "SUPER_ADMIN" Then
        CheckLogisticsRole = True
        Exit Function
    End If
    If currentRole = "PROD_MANAGER" Then
        CheckLogisticsRole = True
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
        Response.Redirect "/admin/unauthorized.asp?module=logistics"
        Response.End
    End If
    
    CheckLogisticsRole = True
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
