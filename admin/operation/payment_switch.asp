<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' 确保所有支付开关配置项存在（在打开任何Recordset之前执行，避免Access MARS问题）
Sub EnsurePaymentSettings()
    Dim settings, s, cnt
    settings = Array("EnableAlipay", "EnableWechatPay", "EnableBankTransfer", "EnablePayPal", "EnableCOD", "EnableStripe", "EnableUnionPay")
    For Each s In settings
        cnt = GetScalar("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey='" & s & "'")
        If CLng("0" & cnt) = 0 Then
            ExecuteNonQuery "INSERT INTO SiteSettings (SettingKey, SettingValue, Description) VALUES ('" & s & "', '1', '" & s & "')"
        End If
    Next
End Sub

Call EnsurePaymentSettings()

' 处理表单提交 - 仅限开关控制
Dim action
action = Request.Form("action")

If action = "update_switch" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        Response.Redirect "payment_switch.asp?error=安全验证失败"
        Response.End
    End If
    
    ' 验证权限 - 只有运营经理及以上可以修改
    If Session("AdminRoleCode") = "OP_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
        Dim enableAlipay, enableWechat, enableBank, enablePayPal, enableCOD, enableStripe, enableUnionPay
        enableAlipay = IIf(Request.Form("enableAlipay") = "1", "1", "0")
        enableWechat = IIf(Request.Form("enableWechat") = "1", "1", "0")
        enableBank = IIf(Request.Form("enableBank") = "1", "1", "0")
        enablePayPal = IIf(Request.Form("enablePayPal") = "1", "1", "0")
        enableCOD = IIf(Request.Form("enableCOD") = "1", "1", "0")
        enableStripe = IIf(Request.Form("enableStripe") = "1", "1", "0")
        enableUnionPay = IIf(Request.Form("enableUnionPay") = "1", "1", "0")
        
        ' 更新支付开关状态
        ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & enableAlipay & "' WHERE SettingKey = 'EnableAlipay'"
        ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & enableWechat & "' WHERE SettingKey = 'EnableWechatPay'"
        ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & enableBank & "' WHERE SettingKey = 'EnableBankTransfer'"
        ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & enablePayPal & "' WHERE SettingKey = 'EnablePayPal'"
        ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & enableCOD & "' WHERE SettingKey = 'EnableCOD'"
        ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & enableStripe & "' WHERE SettingKey = 'EnableStripe'"
        ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & enableUnionPay & "' WHERE SettingKey = 'EnableUnionPay'"
        
        ' 记录日志
        Call LogAdminAction("修改支付开关", "operation", "SiteSettings", "", "支付开关状态变更")
        
        Response.Redirect "payment_switch.asp?msg=保存成功"
    Else
        Response.Redirect "payment_switch.asp?error=权限不足"
    End If
End If

' 获取当前支付开关状态
Dim alipayStatus, wechatStatus, bankStatus, paypalStatus, codStatus, stripeStatus, unionPayStatus
alipayStatus = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableAlipay'")
wechatStatus = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableWechatPay'")
bankStatus = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableBankTransfer'")
paypalStatus = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnablePayPal'")
codStatus = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableCOD'")
stripeStatus = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableStripe'")
unionPayStatus = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableUnionPay'")

' GetScalar返回"0"(记录不存在时)或实际SettingValue, 需要判断是否为"1"
' 注意：GetScalar对无结果返回"0", 对有结果返回实际值
If alipayStatus <> "1" And alipayStatus <> "0" Then alipayStatus = "1"
If wechatStatus <> "1" And wechatStatus <> "0" Then wechatStatus = "1"
If bankStatus <> "1" And bankStatus <> "0" Then bankStatus = "1"
If paypalStatus <> "1" And paypalStatus <> "0" Then paypalStatus = "1"
If codStatus <> "1" And codStatus <> "0" Then codStatus = "1"
If stripeStatus <> "1" And stripeStatus <> "0" Then stripeStatus = "1"
If unionPayStatus <> "1" And unionPayStatus <> "0" Then unionPayStatus = "1"

' 记录访问日志
Call LogAdminAction("查看支付开关", "operation", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>支付开关控制 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .warning-box { background: #fff3e0; border-left: 4px solid #ff9800; padding: 15px; margin-bottom: 25px; border-radius: 4px; }
        .warning-box i { color: #ff9800; margin-right: 8px; }
        .switch-container { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .payment-item { display: flex; align-items: center; justify-content: space-between; padding: 20px; border-bottom: 1px solid #f0f0f0; }
        .payment-item:last-child { border-bottom: none; }
        .payment-info { display: flex; align-items: center; gap: 20px; }
        .payment-icon { width: 50px; height: 50px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 24px; }
        .payment-icon.alipay { background: #1677ff; color: white; }
        .payment-icon.wechat { background: #07c160; color: white; }
        .payment-icon.bank { background: #ff6b6b; color: white; }
        .payment-icon.paypal { background: #003087; color: white; }
        .payment-icon.cod { background: #ff9800; color: white; }
        .payment-icon.stripe { background: #635bff; color: white; }
        .payment-icon.unionpay { background: #c00; color: white; }
        .payment-details h4 { margin: 0 0 5px 0; color: #333; }
        .payment-details p { margin: 0; color: #999; font-size: 13px; }
        
        .toggle-switch { position: relative; width: 60px; height: 30px; }
        .toggle-switch input { opacity: 0; width: 0; height: 0; }
        .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 30px; }
        .slider:before { position: absolute; content: ""; height: 22px; width: 22px; left: 4px; bottom: 4px; background-color: white; transition: .4s; border-radius: 50%; }
        input:checked + .slider { background-color: #4CAF50; }
        input:checked + .slider:before { transform: translateX(30px); }
        
        .readonly-notice { background: #e3f2fd; padding: 10px 15px; border-radius: 6px; color: #1976d2; font-size: 13px; margin-top: 10px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-toggle-on"></i> 支付开关控制</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>支付开关</span>
            </div>
        </div>
        
        <div class="warning-box">
            <i class="fas fa-exclamation-triangle"></i>
            <strong>权限说明：</strong>您当前仅有权控制支付功能的开关状态，无法访问具体支付参数配置。如需配置支付参数，请联系财务管理员。
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
        <% End If %>
        
        <% If Request.QueryString("error") <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(Request.QueryString("error")) %></div>
        <% End If %>
        
        <form method="post" action="payment_switch.asp" class="switch-container">
            <%= GetCSRFTokenField() %>
            <input type="hidden" name="action" value="update_switch">
            
            <div class="payment-item">
                <div class="payment-info">
                    <div class="payment-icon alipay"><i class="fab fa-alipay"></i></div>
                    <div class="payment-details">
                        <h4>支付宝支付</h4>
                        <p>支付宝扫码支付、支付宝快捷支付</p>
                    </div>
                </div>
                <label class="toggle-switch">
                    <input type="checkbox" name="enableAlipay" value="1" <%= IIf(alipayStatus="1", "checked", "") %> 
                    <%= IIf(Session("AdminRoleCode")="OP_STAFF", "disabled", "") %>>
                    <span class="slider"></span>
                </label>
            </div>
            
            <div class="payment-item">
                <div class="payment-info">
                    <div class="payment-icon wechat"><i class="fab fa-weixin"></i></div>
                    <div class="payment-details">
                        <h4>微信支付</h4>
                        <p>微信扫码支付、微信公众号支付</p>
                    </div>
                </div>
                <label class="toggle-switch">
                    <input type="checkbox" name="enableWechat" value="1" <%= IIf(wechatStatus="1", "checked", "") %>
                    <%= IIf(Session("AdminRoleCode")="OP_STAFF", "disabled", "") %>>
                    <span class="slider"></span>
                </label>
            </div>
            
            <div class="payment-item">
                <div class="payment-info">
                    <div class="payment-icon bank"><i class="fas fa-university"></i></div>
                    <div class="payment-details">
                        <h4>银行转账</h4>
                        <p>银行卡转账、网银支付</p>
                    </div>
                </div>
                <label class="toggle-switch">
                    <input type="checkbox" name="enableBank" value="1" <%= IIf(bankStatus="1", "checked", "") %>
                    <%= IIf(Session("AdminRoleCode")="OP_STAFF", "disabled", "") %>>
                    <span class="slider"></span>
                </label>
            </div>
            
            <div class="payment-item">
                <div class="payment-info">
                    <div class="payment-icon paypal"><i class="fab fa-paypal"></i></div>
                    <div class="payment-details">
                        <h4>PayPal</h4>
                        <p>国际支付、信用卡支付</p>
                    </div>
                </div>
                <label class="toggle-switch">
                    <input type="checkbox" name="enablePayPal" value="1" <%= IIf(paypalStatus="1", "checked", "") %>
                    <%= IIf(Session("AdminRoleCode")="OP_STAFF", "disabled", "") %>>
                    <span class="slider"></span>
                </label>
            </div>
            
            <div class="payment-item">
                <div class="payment-info">
                    <div class="payment-icon stripe"><i class="fab fa-stripe"></i></div>
                    <div class="payment-details">
                        <h4>Stripe</h4>
                        <p>国际信用卡支付、Apple Pay、Google Pay</p>
                    </div>
                </div>
                <label class="toggle-switch">
                    <input type="checkbox" name="enableStripe" value="1" <%= IIf(stripeStatus="1", "checked", "") %>
                    <%= IIf(Session("AdminRoleCode")="OP_STAFF", "disabled", "") %>>
                    <span class="slider"></span>
                </label>
            </div>
            
            <div class="payment-item">
                <div class="payment-info">
                    <div class="payment-icon unionpay"><i class="fas fa-credit-card"></i></div>
                    <div class="payment-details">
                        <h4>银联支付</h4>
                        <p>银联卡支付、云闪付</p>
                    </div>
                </div>
                <label class="toggle-switch">
                    <input type="checkbox" name="enableUnionPay" value="1" <%= IIf(unionPayStatus="1", "checked", "") %>
                    <%= IIf(Session("AdminRoleCode")="OP_STAFF", "disabled", "") %>>
                    <span class="slider"></span>
                </label>
            </div>
            
            <div class="payment-item">
                <div class="payment-info">
                    <div class="payment-icon cod"><i class="fas fa-truck"></i></div>
                    <div class="payment-details">
                        <h4>货到付款</h4>
                        <p>收货时现金支付</p>
                    </div>
                </div>
                <label class="toggle-switch">
                    <input type="checkbox" name="enableCOD" value="1" <%= IIf(codStatus="1", "checked", "") %>
                    <%= IIf(Session("AdminRoleCode")="OP_STAFF", "disabled", "") %>>
                    <span class="slider"></span>
                </label>
            </div>
            
            <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #f0f0f0;">
                <% If Session("AdminRoleCode") = "OP_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then %>
                <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存设置</button>
                <% Else %>
                <button type="button" class="btn-save" disabled><i class="fas fa-lock"></i> 无权修改</button>
                <div class="readonly-notice">
                    <i class="fas fa-info-circle"></i> 您当前为运营专员，仅可查看不可修改
                </div>
                <% End If %>
            </div>
        </form>
        
        <div style="text-align: center; margin-top: 20px;">
            <a href="../finance/payment_config.asp" style="color: #667eea; text-decoration: none;">
                <i class="fas fa-arrow-right"></i> 前往财务后台配置支付参数
            </a>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
