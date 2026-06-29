<!--#include file="rate_limiter.asp"-->
<!--#include file="api_auth.asp"-->
<%
' ============================================
' V18.0 API 守卫模块（速率限制 + 认证）
' 用法: <!--#include file="../includes/api_guard.asp"-->
'       If Not API_Guard("api") Then Response.End
' ============================================

' ============================================
' API_Guard: 统一的 API 前置检查
' 参数:
'   bucketName - 限流桶名称 ("api", "checkout", "login", 或 "")
'   requireAuth - 是否要求认证 (默认 True)
' 返回: True（通过）/ False（已拒绝并发送响应）
' ============================================
Function API_Guard(bucketName, requireAuth)
    ' 参数默认值
    If IsEmpty(requireAuth) Or IsNull(requireAuth) Then requireAuth = True
    If bucketName = "" Then bucketName = "api"
    
    ' 1. 速率限制检查
    If FEATURE_RATE_LIMITER Then
        If Not RateLimitCheck(bucketName) Then
            Call RateLimitSend429(bucketName)
            API_Guard = False
            Exit Function
        End If
    End If
    
    ' 2. API 认证检查（Session 或 API Key）
    If FEATURE_API_AUTH And requireAuth Then
        If Not API_AuthCheck() Then
            Response.Status = "401 Unauthorized"
            Response.ContentType = "application/json"
            Dim guardAuthErr
            guardAuthErr = "{""code"":401,""message"":""API 认证失败，请提供有效的认证信息""}"
            On Error Resume Next
            If FEATURE_API_V1 Then
                API_Error 1001, "API 认证失败，请提供有效的认证信息"
                If Err.Number <> 0 Then
                    Err.Clear
                    Response.Write guardAuthErr
                End If
            Else
                Response.Write guardAuthErr
            End If
            On Error GoTo 0
            Response.End
            API_Guard = False
            Exit Function
        End If
    End If
    
    API_Guard = True
End Function
%>
