<%
' 邮件发送工具函数

' 发送密码重置邮件
Sub SendPasswordResetEmail(toEmail, fullName, resetToken)
    Dim subject, body, resetLink
    
    Dim emailProtocol
    If Request.ServerVariables("HTTPS") = "on" Then
        emailProtocol = "https://"
    Else
        emailProtocol = "http://"
    End If
    resetLink = emailProtocol & Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT") & "/admin/reset_password.asp?token=" & resetToken
    
    subject = "管理员密码重置 - " & SITE_NAME
    
    body = "<html><body>"
    body = body & "<h2>密码重置请求</h2>"
    body = body & "<p>亲爱的 " & fullName & "，</p>"
    body = body & "<p>您请求重置管理员账户的密码。请点击下面的链接设置新密码：</p>"
    body = body & "<p><a href='" & resetLink & "'>" & resetLink & "</a></p>"
    body = body & "<p>此链接将在1小时内过期。</p>"
    body = body & "<p>如果您没有请求密码重置，请忽略此邮件。</p>"
    body = body & "<hr>"
    body = body & "<p>香氛定制电商网站 管理系统</p>"
    body = body & "</body></html>"
    
    ' 使用CDONTS发送邮件 (适用于较老的Windows服务器)
    On Error Resume Next
    Dim objMail
    Set objMail = Server.CreateObject("CDONTS.NewMail")
    
    If Err.Number = 0 Then
        objMail.From = SITE_NOREPLY
        objMail.To = toEmail
        objMail.Subject = subject
        objMail.BodyFormat = 0  ' HTML格式
        objMail.MailFormat = 0  ' HTML格式
        objMail.Body = body
        objMail.Send
        Set objMail = Nothing
    Else
        ' 如果CDONTS不可用，记录错误（实际应用中应使用其他邮件服务）
        Session("EmailError") = "邮件发送失败: " & Err.Description
        Err.Clear
    End If
    On Error Goto 0
End Sub

' ============================================
' 用户密码重置邮件（前端用户专用）
' ============================================
Sub SendUserPasswordResetEmail(toEmail, fullName, resetToken)
    Dim subject, body, resetLink
    
    Dim emailProtocol
    If Request.ServerVariables("HTTPS") = "on" Then
        emailProtocol = "https://"
    Else
        emailProtocol = "http://"
    End If
    resetLink = emailProtocol & Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT") & "/user/reset_password.asp?token=" & resetToken
    
    subject = "密码重置 - " & SITE_NAME
    
    body = "<html><body>"
    body = body & "<h2>密码重置请求</h2>"
    body = body & "<p>亲爱的 " & fullName & "，</p>"
    body = body & "<p>您请求重置账户密码。请点击下面的链接设置新密码：</p>"
    body = body & "<p><a href='" & resetLink & "'>" & resetLink & "</a></p>"
    body = body & "<p>此链接将在1小时内过期。</p>"
    body = body & "<p>如果您没有请求密码重置，请忽略此邮件。</p>"
    body = body & "<hr>"
    body = body & "<p>" & SITE_NAME & "</p>"
    body = body & "</body></html>"
    
    On Error Resume Next
    Dim objMail
    Set objMail = Server.CreateObject("CDONTS.NewMail")
    
    If Err.Number = 0 Then
        objMail.From = SITE_NOREPLY
        objMail.To = toEmail
        objMail.Subject = subject
        objMail.BodyFormat = 0
        objMail.MailFormat = 0
        objMail.Body = body
        objMail.Send
        Set objMail = Nothing
    Else
        Session("EmailError") = "邮件发送失败: " & Err.Description
        Err.Clear
    End If
    On Error Goto 0
End Sub

' 发送一般邮件
Sub SendEmail(toEmail, subject, body, isHtml)
    On Error Resume Next
    Dim objMail
    Set objMail = Server.CreateObject("CDONTS.NewMail")
    
    If Err.Number = 0 Then
        objMail.From = SITE_NOREPLY
        objMail.To = toEmail
        objMail.Subject = subject
        If isHtml Then
            objMail.BodyFormat = 0  ' HTML格式
            objMail.MailFormat = 0  ' HTML格式
        End If
        objMail.Body = body
        objMail.Send
        Set objMail = Nothing
    Else
        Session("EmailError") = "邮件发送失败: " & Err.Description
        Err.Clear
    End If
    On Error Goto 0
End Sub
%>