<%
' ============================================
' V15.0 邮件服务增强 (Email Service)
' 依赖: email_utils.asp (基础发送), connection.asp (邮件队列)
' 用法: <!--#include file="email_service.asp"-->
' 调用: ES_SendOrderConfirmation userId, orderId
'        ES_SendShippingNotification userId, orderId, trackingNo
'        ES_SendWelcomeEmail userId
' ============================================

' 邮件模板目录
Const ES_TEMPLATE_DIR = "/email_templates/"

' ============================================
' 内部函数：加载邮件模板
' ============================================
Function ES_LoadTemplate(templateName, replacements)
    Dim fso, filePath, template, key
    On Error Resume Next
    
    filePath = Server.MapPath(ES_TEMPLATE_DIR & templateName & ".html")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Err.Number <> 0 Or Not fso.FileExists(filePath) Then
        Err.Clear
        Set fso = Nothing
        ES_LoadTemplate = ""
        Exit Function
    End If
    
    Dim file
    Set file = fso.OpenTextFile(filePath, 1)
    template = file.ReadAll()
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
    ' 替换占位符 {{key}}
    If IsObject(replacements) Then
        Dim replaceKeys
        replaceKeys = replacements.Keys()
        For Each key In replaceKeys
            template = Replace(template, "{{" & key & "}}", replacements.Item(key))
        Next
    End If
    
    ES_LoadTemplate = template
    Err.Clear
End Function

' ============================================
' 发送注册欢迎邮件
' ============================================
Sub ES_SendWelcomeEmail(userId, toEmail, fullName)
    Dim subject, body, replacements
    
    subject = "欢迎加入 " & SITE_NAME & "！"
    
    Set replacements = Server.CreateObject("Scripting.Dictionary")
    replacements.Add "FULL_NAME", fullName
    replacements.Add "SITE_NAME", SITE_NAME
    replacements.Add "SITE_URL", "https://" & Request.ServerVariables("SERVER_NAME")
    replacements.Add "YEAR", Year(Now())
    
    body = ES_LoadTemplate("welcome", replacements)
    If body = "" Then
        body = "<html><body><h2>欢迎 " & fullName & "！</h2>" & _
               "<p>感谢您注册 " & SITE_NAME & "。</p>" & _
               "<p>您可以开始探索我们独特的香水定制服务。</p></body></html>"
    End If
    
    Set replacements = Nothing
    Call SendEmail(toEmail, subject, body, True)
End Sub

' ============================================
' 发送订单确认邮件
' ============================================
Sub ES_SendOrderConfirmation(userId, orderId)
    Dim toEmail, fullName, orderNo, totalAmount, sql, rs
    
    ' 获取用户和订单信息
    On Error Resume Next
    sql = "SELECT u.Email, u.FullName, o.OrderNo, o.TotalAmount " & _
          "FROM Users u INNER JOIN Orders o ON u.UserID=o.UserID " & _
          "WHERE u.UserID=" & userId & " AND o.OrderID=" & orderId
    Set rs = conn.Execute(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            toEmail = rs("Email")
            fullName = rs("FullName")
            orderNo = rs("OrderNo")
            totalAmount = rs("TotalAmount")
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    If toEmail = "" Then Exit Sub
    
    Dim subject, body, replacements, siteUrl
    siteUrl = "https://" & Request.ServerVariables("SERVER_NAME")
    
    subject = "订单确认 #" & orderNo & " - " & SITE_NAME
    
    Set replacements = Server.CreateObject("Scripting.Dictionary")
    replacements.Add "FULL_NAME", fullName
    replacements.Add "ORDER_NO", orderNo
    replacements.Add "TOTAL_AMOUNT", FormatNumber(totalAmount, 2)
    replacements.Add "ORDER_URL", siteUrl & "/user/order_detail.asp?id=" & orderId
    replacements.Add "SITE_NAME", SITE_NAME
    replacements.Add "SITE_URL", siteUrl
    replacements.Add "YEAR", Year(Now())
    
    body = ES_LoadTemplate("order_confirmation", replacements)
    If body = "" Then
        body = "<html><body><h2>订单确认</h2>" & _
               "<p>亲爱的 " & fullName & "，</p>" & _
               "<p>订单号：<strong>" & orderNo & "</strong></p>" & _
               "<p>金额：¥" & FormatNumber(totalAmount, 2) & "</p>" & _
               "<p>我们将尽快处理您的订单。</p></body></html>"
    End If
    
    Set replacements = Nothing
    Call SendEmail(toEmail, subject, body, True)
End Sub

' ============================================
' 发送发货通知邮件
' ============================================
Sub ES_SendShippingNotification(userId, orderId, trackingNumber, shippingCompany)
    Dim toEmail, fullName, orderNo, sql, rs
    
    On Error Resume Next
    sql = "SELECT u.Email, u.FullName, o.OrderNo FROM Users u " & _
          "INNER JOIN Orders o ON u.UserID=o.UserID " & _
          "WHERE u.UserID=" & userId & " AND o.OrderID=" & orderId
    Set rs = conn.Execute(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            toEmail = rs("Email")
            fullName = rs("FullName")
            orderNo = rs("OrderNo")
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    If toEmail = "" Then Exit Sub
    
    If IsNull(shippingCompany) Or shippingCompany = "" Then shippingCompany = "快递公司"
    
    Dim subject, body, replacements, siteUrl
    siteUrl = "https://" & Request.ServerVariables("SERVER_NAME")
    
    subject = "订单已发货 #" & orderNo & " - " & SITE_NAME
    
    Set replacements = Server.CreateObject("Scripting.Dictionary")
    replacements.Add "FULL_NAME", fullName
    replacements.Add "ORDER_NO", orderNo
    replacements.Add "TRACKING_NO", trackingNumber
    replacements.Add "SHIPPING_COMPANY", shippingCompany
    replacements.Add "ORDER_URL", siteUrl & "/user/order_detail.asp?id=" & orderId
    replacements.Add "SITE_NAME", SITE_NAME
    replacements.Add "SITE_URL", siteUrl
    
    body = ES_LoadTemplate("shipping_notification", replacements)
    If body = "" Then
        body = "<html><body><h2>订单发货通知</h2>" & _
               "<p>亲爱的 " & fullName & "，</p>" & _
               "<p>订单号：<strong>" & orderNo & "</strong></p>" & _
               "<p>快递单号：<strong>" & trackingNumber & "</strong> (" & shippingCompany & ")</p>" & _
               "<p>请留意收货。</p></body></html>"
    End If
    
    Set replacements = Nothing
    Call SendEmail(toEmail, subject, body, True)
End Sub

' ============================================
' 发送密码重置邮件（增强版）
' ============================================
Sub ES_SendPasswordReset(toEmail, fullName, resetToken, isAdmin)
    Dim subject, body, replacements, resetPath, siteUrl
    siteUrl = "https://" & Request.ServerVariables("SERVER_NAME")
    
    If isAdmin Then
        resetPath = "/admin/forgot_password.asp?token=" & resetToken
    Else
        resetPath = "/user/login.asp?reset=" & resetToken
    End If
    
    subject = "密码重置请求 - " & SITE_NAME
    
    Set replacements = Server.CreateObject("Scripting.Dictionary")
    replacements.Add "FULL_NAME", fullName
    replacements.Add "RESET_LINK", siteUrl & resetPath
    replacements.Add "SITE_NAME", SITE_NAME
    replacements.Add "SITE_URL", siteUrl
    
    body = ES_LoadTemplate("password_reset", replacements)
    If body = "" Then
        body = "<html><body><h2>密码重置</h2>" & _
               "<p>亲爱的 " & fullName & "，</p>" & _
               "<p>请点击以下链接重置密码：</p>" & _
               "<p><a href='" & siteUrl & resetPath & "'>重置密码</a></p>" & _
               "<p>此链接1小时内有效。</p></body></html>"
    End If
    
    Set replacements = Nothing
    Call SendEmail(toEmail, subject, body, True)
End Sub

' ============================================
' 发送退款处理通知
' ============================================
Sub ES_SendRefundNotification(userId, orderId, refundAmount)
    Dim toEmail, fullName, orderNo, sql, rs
    
    On Error Resume Next
    sql = "SELECT u.Email, u.FullName, o.OrderNo FROM Users u " & _
          "INNER JOIN Orders o ON u.UserID=o.UserID " & _
          "WHERE u.UserID=" & userId & " AND o.OrderID=" & orderId
    Set rs = conn.Execute(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            toEmail = rs("Email")
            fullName = rs("FullName")
            orderNo = rs("OrderNo")
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    If toEmail = "" Then Exit Sub
    
    Dim subject, body, siteUrl
    siteUrl = "https://" & Request.ServerVariables("SERVER_NAME")
    subject = "退款处理通知 #" & orderNo & " - " & SITE_NAME
    
    body = "<html><body><h2>退款处理通知</h2>" & _
           "<p>亲爱的 " & fullName & "，</p>" & _
           "<p>订单号：<strong>" & orderNo & "</strong></p>" & _
           "<p>退款金额：<strong>¥" & FormatNumber(refundAmount, 2) & "</strong></p>" & _
           "<p>退款将按原支付方式返还，请留意查收。</p>" & _
           "<p><a href='" & siteUrl & "/user/order_detail.asp?id=" & orderId & "'>查看订单</a></p></body></html>"
    
    Call SendEmail(toEmail, subject, body, True)
End Sub
%>