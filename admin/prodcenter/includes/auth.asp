<!--#include file="../../includes/role_auth.asp"-->
<%
' 产品生产管理中心认证
Call VerifyModuleAccess("prodcenter", 1)
Call LogAdminAction("访问产品生产管理中心", "prodcenter", "", "", "")

' 全局IIf函数（用于在ASP表达式中做简化条件判断）
Function IIf(condition, trueVal, falseVal)
    If condition Then IIf = trueVal Else IIf = falseVal
End Function

Function CheckProdcenterRole(allowedRoles)
    Dim currentRole, roles, i, allowed
    currentRole = Session("AdminRoleCode")
    
    If currentRole = "SUPER_ADMIN" Then
        CheckProdcenterRole = True
        Exit Function
    End If
    If currentRole = "PROD_MANAGER" Then
        CheckProdcenterRole = True
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
        Response.Redirect "/admin/unauthorized.asp?module=prodcenter"
        Response.End
    End If
    
    CheckProdcenterRole = True
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
