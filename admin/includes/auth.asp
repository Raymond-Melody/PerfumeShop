<%
' ============================================
' 管理员认证组件
' 
' 建议: 在页面开头包含此文件前，先包含:
'   <!--#include file="../includes/config.asp"-->
'   <!--#include file="../includes/connection.asp"-->
' 以启用Remember Me Cookie恢复功能
' ============================================

' 检查管理员是否已登录
If Session("AdminID") = "" Then
    ' 尝试从Remember Me Cookie恢复会话
    ' 只有当所需函数可用时才执行
    If Request.Cookies("AdminRememberMe") <> "" Then
        On Error Resume Next
        
        ' 检查ValidateSecureToken函数是否可用
        Dim testFunc
        testFunc = ValidateSecureToken("")
        
        If Err.Number = 0 Then
            Err.Clear
            
            Dim cookieToken, validatedAdminId
            cookieToken = Request.Cookies("AdminRememberMe")
            
            ' 使用安全令牌验证函数
            validatedAdminId = ValidateSecureToken(cookieToken)
            
            If validatedAdminId <> "" And IsNumeric(validatedAdminId) Then
                ' 从数据库验证管理员ID
                Call OpenConnection
                
                Dim rsAuth
                Set rsAuth = ExecuteQuery("SELECT AdminID, Username, IsActive FROM AdminUsers WHERE AdminID = " & CLng(validatedAdminId) & " AND IsActive = 1")
                
                If Not rsAuth Is Nothing And Not rsAuth.EOF Then
                    ' 会话恢复成功
                    Session("AdminID") = rsAuth("AdminID")
                    Session("AdminUsername") = rsAuth("Username")
                End If
                
                If Not rsAuth Is Nothing Then
                    rsAuth.Close
                    Set rsAuth = Nothing
                End If
                
                Call CloseConnection
            End If
        Else
            Err.Clear
        End If
        
        On Error GoTo 0
    End If
    
    ' 如果仍然没有登录，重定向到登录页面
    If Session("AdminID") = "" Then
        Response.Redirect "/admin/login.asp"
        Response.End
    End If
End If
%>
