<%@ Language="VBScript" CodePage="65001" %>
<%
' V10: 清除Remember Me Cookie时使用相同的安全标志
Response.Cookies("AdminRememberMe") = ""
Response.Cookies("AdminRememberMe").Expires = DateAdd("d", -1, Now())
Response.Cookies("AdminRememberMe").Path = "/"
' 注：HttpOnly需要IIS 7.0+支持，此处移除以保证兼容性
' Response.Cookies("AdminRememberMe").HttpOnly = True
If LCase(Request.ServerVariables("HTTPS")) = "on" Then
    Response.Cookies("AdminRememberMe").Secure = True
End If

Session.Abandon
Response.Redirect "login.asp"
%>