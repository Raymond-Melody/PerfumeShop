<!--#include file="../../includes/role_auth.asp"-->
<%
' 库存管理中心认证
Call VerifyModuleAccess("inventory", 1)
Call LogAdminAction("访问库存管理中心", "inventory", "", "", "")

Function CheckInventoryRole(allowedRoles)
    Dim currentRole, roles, i, allowed
    currentRole = Session("AdminRoleCode")
    
    If currentRole = "SUPER_ADMIN" Then
        CheckInventoryRole = True
        Exit Function
    End If
    If currentRole = "PROD_MANAGER" Then
        CheckInventoryRole = True
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
        Response.Redirect "/admin/unauthorized.asp?module=inventory"
        Response.End
    End If
    
    CheckInventoryRole = True
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
