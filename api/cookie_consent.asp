<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V18.0 Cookie 同意记录 API
' 接收前端 Cookie 同意选择并记录到审计日志
' 调用方: header.asp 中的 acceptCookies() JS
' ============================================
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/api_response.asp"-->
<!--#include file="../includes/audit_utils.asp"-->
<%
' 仅接受 POST 请求
If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
    Call API_Error(API_ERR_PARAM_MISSING, "请使用 POST 请求")
    Response.End
End If

' CSRF 验证
If Not API_CheckCSRF() Then
    Call API_Error(API_ERR_CSRF_INVALID, "安全验证失败")
    Response.End
End If

Call OpenConnection()

Dim consentLevel
consentLevel = Trim(Request.Form("consent"))

' 验证 consent 值
If consentLevel = "" Then
    consentLevel = "essential"
End If

If consentLevel <> "all" And consentLevel <> "essential" Then
    consentLevel = "essential"
End If

' 记录 Cookie 同意到审计日志
Call LogCookieConsent(True, consentLevel)

Call API_Success(Null, "Cookie 偏好已记录")

Call CloseConnection()
%>
