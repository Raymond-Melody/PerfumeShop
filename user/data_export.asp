<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 必须登录
If Session("UserID") = "" Or IsEmpty(Session("UserID")) Then
    Response.Redirect "login.asp?return=" & Server.URLEncode("data_export.asp")
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<%
Call OpenConnection()

Dim userId, format, action, errorMsg, successMsg
userId = CLng(Session("UserID"))
format = Trim(Request.QueryString("format"))
action = Trim(Request.QueryString("action"))
errorMsg = ""
successMsg = ""

' 处理导出请求
If action = "export" Then
    If format = "" Then format = "json"
    
    ' 验证 CSRF
    Dim csrfToken
    csrfToken = Trim(Request.QueryString("csrf_token"))
    If csrfToken = "" Then csrfToken = Trim(Request.Form("csrf_token"))
    If csrfToken <> Session("CSRFToken") Then
        errorMsg = T("privacy_export_csrf_error", Empty)
        If errorMsg = "" Then errorMsg = "安全验证失败，请刷新页面重试"
    Else
        ' 记录隐私操作审计日志
        If FEATURE_GDPR_COMPLIANCE Then
            On Error Resume Next
            Dim auditParams(5)
            ReDim auditParams(0)
            auditParams(0) = Array("@LogLevel", DAL_adVarChar, 10, "PRIVACY")
            ReDim Preserve auditParams(1)
            auditParams(1) = Array("@LogMessage", DAL_adVarChar, 2000, "用户数据导出请求 - UserID:" & userId & " 格式:" & UCase(format))
            ReDim Preserve auditParams(2)
            auditParams(2) = Array("@LogSource", DAL_adVarChar, 100, "user/data_export.asp")
            ReDim Preserve auditParams(3)
            auditParams(3) = Array("@IPAddress", DAL_adVarChar, 50, Left(Request.ServerVariables("REMOTE_ADDR"), 50))
            ReDim Preserve auditParams(4)
            auditParams(4) = Array("@PageURL", DAL_adVarChar, 500, Left(Request.ServerVariables("SCRIPT_NAME"), 500))
            DAL_Execute "INSERT INTO AppLogs (LogLevel, LogMessage, LogSource, IPAddress, PageURL) VALUES (@LogLevel, @LogMessage, @LogSource, @IPAddress, @PageURL)", auditParams
            Err.Clear
            On Error GoTo 0
        End If
        
        Select Case LCase(format)
            Case "json"
                Call ExportDataJSON(userId)
            Case "csv"
                Call ExportDataCSV(userId)
            Case Else
                Call ExportDataJSON(userId)
        End Select
        Response.End
    End If
End If

Call EnsureCSRFToken()
%>
<!--#include file="../includes/header.asp"-->

<div class="export-page">
    <div class="container">
        <div class="export-header">
            <h1><% If FEATURE_I18N Then %><%= T("privacy_export_title", Empty) %><% Else %>数据导出<% End If %></h1>
            <p><% If FEATURE_I18N Then %><%= T("privacy_export_desc", Empty) %><% Else %>根据数据保护法规，您可以导出我们持有的您的个人数据副本。<% End If %></p>
        </div>

        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= HTMLEncode(errorMsg) %></div>
        <% End If %>

        <% If successMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= HTMLEncode(successMsg) %></div>
        <% End If %>

        <div class="export-options">
            <div class="export-card">
                <div class="export-icon"><i class="fas fa-file-code"></i></div>
                <h3><% If FEATURE_I18N Then %><%= T("privacy_export_json", Empty) %><% Else %>JSON 格式<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("privacy_export_json_desc", Empty) %><% Else %>结构化数据，适合程序处理和数据迁移<% End If %></p>
                <a href="?action=export&format=json&csrf_token=<%= Session("CSRFToken") %>" class="btn btn-primary">
                    <i class="fas fa-download"></i> <% If FEATURE_I18N Then %><%= T("privacy_export_btn_json", Empty) %><% Else %>导出 JSON<% End If %>
                </a>
            </div>

            <div class="export-card">
                <div class="export-icon"><i class="fas fa-file-csv"></i></div>
                <h3><% If FEATURE_I18N Then %><%= T("privacy_export_csv", Empty) %><% Else %>CSV 格式<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("privacy_export_csv_desc", Empty) %><% Else %>表格数据，可用 Excel 打开查看<% End If %></p>
                <a href="?action=export&format=csv&csrf_token=<%= Session("CSRFToken") %>" class="btn btn-secondary">
                    <i class="fas fa-download"></i> <% If FEATURE_I18N Then %><%= T("privacy_export_btn_csv", Empty) %><% Else %>导出 CSV<% End If %>
                </a>
            </div>
        </div>

        <div class="export-info">
            <h3><% If FEATURE_I18N Then %><%= T("privacy_export_what", Empty) %><% Else %>导出内容包含：<% End If %></h3>
            <ul>
                <li><i class="fas fa-user"></i> <% If FEATURE_I18N Then %><%= T("privacy_export_item_profile", Empty) %><% Else %>个人资料（用户名、邮箱、姓名、手机）<% End If %></li>
                <li><i class="fas fa-map-marker-alt"></i> <% If FEATURE_I18N Then %><%= T("privacy_export_item_address", Empty) %><% Else %>收货地址<% End If %></li>
                <li><i class="fas fa-shopping-bag"></i> <% If FEATURE_I18N Then %><%= T("privacy_export_item_orders", Empty) %><% Else %>订单记录<% End If %></li>
                <li><i class="fas fa-flask"></i> <% If FEATURE_I18N Then %><%= T("privacy_export_item_custom", Empty) %><% Else %>定制配方记录<% End If %></li>
                <li><i class="fas fa-heart"></i> <% If FEATURE_I18N Then %><%= T("privacy_export_item_favorites", Empty) %><% Else %>收藏列表<% End If %></li>
            </ul>
            <p class="export-note">
                <i class="fas fa-lock"></i> 
                <% If FEATURE_I18N Then %><%= T("privacy_export_note", Empty) %><% Else %>密码等敏感信息不会包含在导出数据中。数据导出操作会被记录在审计日志中。<% End If %>
            </p>
        </div>
    </div>
</div>

<style>
.export-page { padding: 40px 0; max-width: 800px; margin: 0 auto; }
.export-header { text-align: center; margin-bottom: 32px; }
.export-header h1 { font-size: 2rem; color: #2d3748; margin-bottom: 8px; }
.export-header p { color: #718096; }
.export-options { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 32px; }
.export-card { background: #fff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 32px 24px; text-align: center; transition: box-shadow 0.2s; }
.export-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
.export-icon { font-size: 3rem; color: #667eea; margin-bottom: 16px; }
.export-card h3 { font-size: 1.2rem; color: #2d3748; margin-bottom: 8px; }
.export-card p { color: #718096; font-size: 0.9rem; margin-bottom: 20px; }
.export-info { background: #f7fafc; border-radius: 12px; padding: 24px; }
.export-info h3 { font-size: 1.1rem; color: #2d3748; margin-bottom: 12px; }
.export-info ul { list-style: none; padding: 0; margin-bottom: 16px; }
.export-info li { color: #4a5568; padding: 6px 0; }
.export-info li i { color: #667eea; width: 20px; margin-right: 8px; }
.export-note { color: #a0aec0; font-size: 0.85rem; border-top: 1px solid #e2e8f0; padding-top: 12px; margin: 0; }
.alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 20px; }
.alert-error { background: #fff5f5; border: 1px solid #fed7d7; color: #c53030; }
.alert-success { background: #f0fff4; border: 1px solid #c6f6d5; color: #276749; }
@media (max-width: 640px) {
    .export-options { grid-template-columns: 1fr; }
}
</style>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()

' ============================================
' JSON 格式导出
' ============================================
Sub ExportDataJSON(userId)
    Response.ContentType = "application/json"
    Response.AddHeader "Content-Disposition", "attachment; filename=my_data_" & userId & ".json"
    
    Dim json, rs, parts, i
    
    ' 用户资料
    Dim userParams(0)
    userParams(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    Set rs = DAL_GetRow("SELECT Username, Email, FullName, Phone, CreatedAt FROM Users WHERE UserID=@UserID", userParams)
    
    json = "{"
    json = json & """export_date"":""" & Now() & ""","
    json = json & """user_id"":" & userId & ","
    
    If Not rs Is Nothing And Not rs.EOF Then
        json = json & """profile"":{"
        json = json & """username"":""" & JsonSafe(rs("Username")) & ""","
        json = json & """email"":""" & JsonSafe(rs("Email")) & ""","
        json = json & """full_name"":""" & JsonSafe(rs("FullName")) & ""","
        json = json & """phone"":""" & JsonSafe(rs("Phone")) & ""","
        json = json & """created_at"":""" & JsonSafe(rs("CreatedAt")) & """"
        json = json & "},"
        rs.Close
    End If
    Set rs = Nothing
    
    ' 收货地址
    Set rs = DAL_GetList("SELECT AddressID, RecipientName, Phone, Province, City, District, DetailAddress, IsDefault FROM UserAddresses WHERE UserID=@UserID ORDER BY IsDefault DESC", userParams)
    json = json & """addresses"":["
    parts = ""
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If parts <> "" Then parts = parts & ","
            parts = parts & "{"
            parts = parts & """id"":" & rs("AddressID") & ","
            parts = parts & """recipient"":""" & JsonSafe(rs("RecipientName")) & ""","
            parts = parts & """phone"":""" & JsonSafe(rs("Phone")) & ""","
            parts = parts & """province"":""" & JsonSafe(rs("Province")) & ""","
            parts = parts & """city"":""" & JsonSafe(rs("City")) & ""","
            parts = parts & """district"":""" & JsonSafe(rs("District")) & ""","
            parts = parts & """detail"":""" & JsonSafe(rs("DetailAddress")) & ""","
            parts = parts & """is_default"":" & LCase(CStr(rs("IsDefault") = True))
            parts = parts & "}"
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    json = json & parts & "],"
    
    ' 订单记录
    Set rs = DAL_GetList("SELECT OrderID, OrderNo, TotalAmount, Status, PaymentMethod, CreatedAt FROM Orders WHERE UserID=@UserID ORDER BY CreatedAt DESC", userParams)
    json = json & """orders"":["
    parts = ""
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If parts <> "" Then parts = parts & ","
            parts = parts & "{"
            parts = parts & """order_id"":" & rs("OrderID") & ","
            parts = parts & """order_no"":""" & JsonSafe(rs("OrderNo")) & ""","
            parts = parts & """amount"":" & CDbl(rs("TotalAmount")) & ","
            parts = parts & """status"":""" & JsonSafe(rs("Status")) & ""","
            parts = parts & """payment"":""" & JsonSafe(rs("PaymentMethod")) & ""","
            parts = parts & """created"":""" & JsonSafe(rs("CreatedAt")) & """"
            parts = parts & "}"
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    json = json & parts & "],"
    
    ' 收藏列表
    Set rs = DAL_GetList("SELECT p.ProductID, p.ProductName, p.BasePrice, f.CreatedTime FROM UserFavorites f INNER JOIN Products p ON f.ProductID=p.ProductID WHERE f.UserID=@UserID ORDER BY f.CreatedTime DESC", userParams)
    json = json & """favorites"":["
    parts = ""
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If parts <> "" Then parts = parts & ","
            parts = parts & "{"
            parts = parts & """product_id"":" & rs("ProductID") & ","
            parts = parts & """name"":""" & JsonSafe(rs("ProductName")) & ""","
            parts = parts & """price"":" & CDbl(rs("BasePrice")) & ","
            parts = parts & """added"":""" & JsonSafe(rs("CreatedTime")) & """"
            parts = parts & "}"
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    json = json & parts & "]"
    
    json = json & "}"
    Response.Write json
End Sub

' ============================================
' CSV 格式导出
' ============================================
Sub ExportDataCSV(userId)
    Dim userParams(0)
    userParams(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    
    Response.ContentType = "text/csv; charset=UTF-8"
    Response.AddHeader "Content-Disposition", "attachment; filename=my_data_" & userId & ".csv"
    Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF) ' BOM
    
    ' 用户资料
    Response.Write "=== 个人资料 ===" & vbCrLf
    Dim rs
    Set rs = DAL_GetRow("SELECT Username, Email, FullName, Phone, CreatedAt FROM Users WHERE UserID=@UserID", userParams)
    If Not rs Is Nothing And Not rs.EOF Then
        Response.Write "用户名,邮箱,姓名,手机,注册时间" & vbCrLf
        Response.Write """" & rs("Username") & """,""" & rs("Email") & """,""" & rs("FullName") & """,""" & rs("Phone") & """,""" & rs("CreatedAt") & """" & vbCrLf
        rs.Close
    End If
    Set rs = Nothing
    Response.Write vbCrLf
    
    ' 收货地址
    Response.Write "=== 收货地址 ===" & vbCrLf
    Response.Write "ID,收件人,电话,省,市,区,详细地址,默认" & vbCrLf
    Set rs = DAL_GetList("SELECT AddressID, RecipientName, Phone, Province, City, District, DetailAddress, IsDefault FROM UserAddresses WHERE UserID=@UserID ORDER BY IsDefault DESC", userParams)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            Response.Write rs("AddressID") & ",""" & rs("RecipientName") & """,""" & rs("Phone") & """,""" & rs("Province") & """,""" & rs("City") & """,""" & rs("District") & """,""" & rs("DetailAddress") & """," & rs("IsDefault") & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    Response.Write vbCrLf
    
    ' 订单记录
    Response.Write "=== 订单记录 ===" & vbCrLf
    Response.Write "订单ID,订单号,金额,状态,支付方式,创建时间" & vbCrLf
    Set rs = DAL_GetList("SELECT OrderID, OrderNo, TotalAmount, Status, PaymentMethod, CreatedAt FROM Orders WHERE UserID=@UserID ORDER BY CreatedAt DESC", userParams)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            Response.Write rs("OrderID") & ",""" & rs("OrderNo") & """," & CDbl(rs("TotalAmount")) & ",""" & rs("Status") & """,""" & rs("PaymentMethod") & """,""" & rs("CreatedAt") & """" & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
End Sub

' ============================================
' JSON 安全转义
' ============================================
Function JsonSafe(val)
    If IsNull(val) Then
        JsonSafe = ""
        Exit Function
    End If
    Dim s : s = CStr(val)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonSafe = s
End Function
%>
