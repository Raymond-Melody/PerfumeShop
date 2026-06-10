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

' 安全保存配置项的函数（UPSERT模式）
Sub SaveConfig(key, value)
    Dim checkSQL, updateSQL, insertSQL
    checkSQL = "SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = '" & SafeSQL(key) & "'"
    If GetScalar(checkSQL) > 0 Then
        updateSQL = "UPDATE SiteSettings SET SettingValue = '" & SafeSQL(value) & "' WHERE SettingKey = '" & SafeSQL(key) & "'"
        ExecuteNonQuery updateSQL
    Else
        insertSQL = "INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('" & SafeSQL(key) & "', '" & SafeSQL(value) & "')"
        ExecuteNonQuery insertSQL
    End If
End Sub

' 处理表单提交
Dim action
action = Request.Form("action")

If action = "save_config" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        Response.Redirect "payment_config.asp?error=安全验证失败"
        Response.End
    End If
    
    ' 验证权限 - 财务经理及以上
    If Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
        ' 支付开关
        SaveConfig "EnableAlipay", IIf(Request.Form("enableAlipay")="1", "1", "0")
        SaveConfig "EnableWechatPay", IIf(Request.Form("enableWechat")="1", "1", "0")
        SaveConfig "EnableBankTransfer", IIf(Request.Form("enableBank")="1", "1", "0")
        SaveConfig "EnableStripe", IIf(Request.Form("enableStripe")="1", "1", "0")
        SaveConfig "EnableUnionPay", IIf(Request.Form("enableUnionPay")="1", "1", "0")
        SaveConfig "EnablePayPal", IIf(Request.Form("enablePayPal")="1", "1", "0")
        
        ' 支付宝配置
        SaveConfig "AlipayAppId", Request.Form("alipayAppId")
        SaveConfig "AlipayMerchantId", Request.Form("alipayMerchantId")
        SaveConfig "AlipayFeeRate", Request.Form("alipayFeeRate")
        
        ' 微信配置
        SaveConfig "WechatAppId", Request.Form("wechatAppId")
        SaveConfig "WechatMchId", Request.Form("wechatMchId")
        SaveConfig "WechatFeeRate", Request.Form("wechatFeeRate")
        
        ' 银行配置
        SaveConfig "BankAccountName", Request.Form("bankAccountName")
        SaveConfig "BankAccountNo", Request.Form("bankAccountNo")
        SaveConfig "BankName", Request.Form("bankName")
        
        ' Stripe配置
        SaveConfig "StripePublishableKey", Request.Form("stripePublishableKey")
        SaveConfig "StripeSecretKey", Request.Form("stripeSecretKey")
        SaveConfig "StripeWebhookSecret", Request.Form("stripeWebhookSecret")
        SaveConfig "StripeFeeRate", Request.Form("stripeFeeRate")
        SaveConfig "StripeFixedFee", Request.Form("stripeFixedFee")
        
        ' 银联配置
        SaveConfig "UnionPayMerchantId", Request.Form("unionPayMerchantId")
        SaveConfig "UnionPayCertPath", Request.Form("unionPayCertPath")
        SaveConfig "UnionPayFeeRate", Request.Form("unionPayFeeRate")
        
        ' PayPal配置
        SaveConfig "PayPalClientId", Request.Form("paypalClientId")
        SaveConfig "PayPalSecret", Request.Form("paypalSecret")
        SaveConfig "PayPalSandbox", IIf(Request.Form("paypalSandbox")="1", "1", "0")
        SaveConfig "PayPalFeeRate", Request.Form("paypalFeeRate")
        SaveConfig "PayPalFixedFee", Request.Form("paypalFixedFee")
        
        ' 全局配置
        SaveConfig "PaymentTestMode", IIf(Request.Form("paymentTestMode")="1", "1", "0")
        SaveConfig "DefaultPaymentMethod", Request.Form("defaultPaymentMethod")
        
        Call LogAdminAction("配置支付参数", "finance", "SiteSettings", "", "支付参数配置更新")
        Response.Redirect "payment_config.asp?msg=保存成功"
    Else
        Response.Redirect "payment_config.asp?error=权限不足"
    End If
End If

' 获取当前配置
Function GetConfig(key)
    Dim val
    val = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = '" & key & "'")
    If IsNull(val) Then val = ""
    GetConfig = val
End Function

' 获取配置值（带默认值）
Function GetConfigWithDefault(key, defaultValue)
    Dim val
    val = GetConfig(key)
    If val = "" Then val = defaultValue
    GetConfigWithDefault = val
End Function

Dim enableAlipay, enableWechat, enableBank, enableStripe, enableUnionPay, enablePayPal
enableAlipay = GetConfig("EnableAlipay")
enableWechat = GetConfig("EnableWechatPay")
enableBank = GetConfig("EnableBankTransfer")
enableStripe = GetConfig("EnableStripe")
enableUnionPay = GetConfig("EnableUnionPay")
enablePayPal = GetConfig("EnablePayPal")

Dim paymentTestMode, defaultPaymentMethod
paymentTestMode = GetConfig("PaymentTestMode")
defaultPaymentMethod = GetConfig("DefaultPaymentMethod")

Call LogAdminAction("查看支付配置", "finance", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>支付配置管理 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .config-container { display: grid; grid-template-columns: repeat(2, 1fr); gap: 25px; }
        .config-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.06); }
        .config-card h3 { font-size: 18px; color: #e0e0e0; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; padding-bottom: 15px; border-bottom: 2px solid rgba(255,255,255,0.06); cursor: pointer; }
        .config-card h3 .toggle-icon { margin-left: auto; transition: transform 0.3s; }
        .config-card.collapsed h3 .toggle-icon { transform: rotate(-90deg); }
        .config-card.collapsed .card-content { display: none; }
        
        .config-card.alipay h3 { color: #1677ff; border-color: #1677ff; }
        .config-card.wechat h3 { color: #07c160; border-color: #07c160; }
        .config-card.bank h3 { color: #ff6b6b; border-color: #ff6b6b; }
        .config-card.stripe h3 { color: #635bff; border-color: #635bff; }
        .config-card.unionpay h3 { color: #c00; border-color: #c00; }
        .config-card.paypal h3 { color: #003087; border-color: #003087; }
        .config-card.global h3 { color: #ffa726; border-color: #ffa726; }
        
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; color: #b0b0b0; font-weight: 500; font-size: 14px; }
        .form-group input, .form-group select { 
            width: 100%; padding: 12px 15px; border: 2px solid #3a3a4a; border-radius: 8px; 
            font-size: 14px; transition: border-color 0.3s; background: #1a1a2e; color: #e0e0e0;
        }
        .form-group input:focus, .form-group select:focus { border-color: #00bcd4; outline: none; }
        .form-group .help-text { font-size: 12px; color: #888; margin-top: 5px; }
        
        .switch-group { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; padding-bottom: 20px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .switch-label { display: flex; align-items: center; gap: 10px; }
        .switch-label i { font-size: 24px; }
        .toggle-switch { position: relative; width: 60px; height: 30px; }
        .toggle-switch input { opacity: 0; width: 0; height: 0; }
        .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #555; transition: .4s; border-radius: 30px; }
        .slider:before { position: absolute; content: ""; height: 22px; width: 22px; left: 4px; bottom: 4px; background-color: white; transition: .4s; border-radius: 50%; }
        input:checked + .slider { background-color: #4CAF50; }
        input:checked + .slider:before { transform: translateX(30px); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-online { background: #1b5e20; color: #81c784; }
        .status-offline { background: #5e1b1b; color: #e57373; }
        
        .btn-save { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; padding: 15px 50px; border: none; border-radius: 8px; font-size: 16px; cursor: pointer; }
        .btn-save:hover { opacity: 0.9; }
        
        .readonly-mask { position: relative; }
        .readonly-mask::after { content: "无权编辑"; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: rgba(26,26,46,0.9); display: flex; align-items: center; justify-content: center; font-size: 18px; color: #888; border-radius: 12px; }
        
        /* 费率对比表样式 */
        .fee-comparison-section { margin-top: 30px; }
        .fee-comparison-section h3 { color: #e0e0e0; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
        .fee-table { width: 100%; border-collapse: collapse; background: #2a2a3a; border-radius: 12px; overflow: hidden; }
        .fee-table th, .fee-table td { padding: 15px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .fee-table th { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; font-weight: 600; }
        .fee-table td { color: #e0e0e0; }
        .fee-table tr:hover { background: #323242; }
        .fee-table .method-name { display: flex; align-items: center; gap: 10px; }
        .fee-table .method-icon { width: 32px; height: 32px; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-size: 16px; }
        .fee-table .icon-alipay { background: #1677ff20; color: #1677ff; }
        .fee-table .icon-wechat { background: #07c16020; color: #07c160; }
        .fee-table .icon-bank { background: #ff6b6b20; color: #ff6b6b; }
        .fee-table .icon-stripe { background: #635bff20; color: #635bff; }
        .fee-table .icon-unionpay { background: #c0000020; color: #c00; }
        .fee-table .icon-paypal { background: #00308720; color: #0070ba; }
        
        .global-config-banner { 
            background: linear-gradient(135deg, #1a1a2e 0%, #2a2a3a 100%); 
            border: 1px solid #3a3a4a; border-radius: 12px; padding: 20px; margin-bottom: 25px;
            display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 20px;
        }
        .global-config-banner .mode-indicator { display: flex; align-items: center; gap: 15px; }
        .global-config-banner .mode-indicator i { font-size: 28px; }
        .global-config-banner .mode-test i { color: #ffa726; }
        .global-config-banner .mode-prod i { color: #4CAF50; }
        .global-config-banner .mode-text h4 { margin: 0; color: #e0e0e0; font-size: 16px; }
        .global-config-banner .mode-text p { margin: 5px 0 0; color: #888; font-size: 13px; }
        
        @media (max-width: 1200px) {
            .config-container { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-credit-card"></i> 支付配置管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>支付配置</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %></div>
        <% End If %>
        
        <% If Request.QueryString("error") <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(Request.QueryString("error")) %></div>
        <% End If %>
        
        <form method="post" action="payment_config.asp">
            <%= GetCSRFTokenField() %>
            <input type="hidden" name="action" value="save_config">
            
            <!-- 全局配置横幅 -->
            <div class="global-config-banner">
                <div class="mode-indicator <%= IIf(paymentTestMode="1", "mode-test", "mode-prod") %>">
                    <i class="fas <%= IIf(paymentTestMode="1", "fa-flask", "fa-globe") %>"></i>
                    <div class="mode-text">
                        <h4>当前模式：<%= IIf(paymentTestMode="1", "测试模式", "生产模式") %></h4>
                        <p><%= IIf(paymentTestMode="1", "所有支付将使用沙箱环境，不会产生真实交易", "所有支付将使用真实环境，请谨慎操作") %></p>
                    </div>
                </div>
                <div style="display: flex; align-items: center; gap: 20px;">
                    <div class="switch-group" style="margin: 0; padding: 0; border: none;">
                        <div class="switch-label">
                            <span style="color: #b0b0b0;">测试模式</span>
                        </div>
                        <label class="toggle-switch">
                            <input type="checkbox" name="paymentTestMode" value="1" <%= IIf(paymentTestMode="1", "checked", "") %>>
                            <span class="slider"></span>
                        </label>
                    </div>
                    <div class="form-group" style="margin: 0; min-width: 200px;">
                        <label style="margin-bottom: 5px;">默认支付方式</label>
                        <select name="defaultPaymentMethod">
                            <option value="alipay" <%= IIf(defaultPaymentMethod="alipay", "selected", "") %>>支付宝</option>
                            <option value="wechat" <%= IIf(defaultPaymentMethod="wechat", "selected", "") %>>微信支付</option>
                            <option value="bank" <%= IIf(defaultPaymentMethod="bank", "selected", "") %>>银行转账</option>
                            <option value="stripe" <%= IIf(defaultPaymentMethod="stripe", "selected", "") %>>Stripe</option>
                            <option value="unionpay" <%= IIf(defaultPaymentMethod="unionpay", "selected", "") %>>银联</option>
                            <option value="paypal" <%= IIf(defaultPaymentMethod="paypal", "selected", "") %>>PayPal</option>
                        </select>
                    </div>
                </div>
            </div>
            
            <div class="config-container">
                <!-- 支付宝配置 -->
                <div class="config-card alipay">
                    <h3 onclick="toggleCard(this)">
                        <i class="fab fa-alipay"></i> 支付宝支付
                        <span class="status-badge <%= IIf(enableAlipay="1", "status-online", "status-offline") %>"><%= IIf(enableAlipay="1", "已上线", "已下线") %></span>
                        <i class="fas fa-chevron-down toggle-icon"></i>
                    </h3>
                    <div class="card-content">
                        <div class="switch-group">
                            <div class="switch-label">
                                <i class="fab fa-alipay" style="color: #1677ff;"></i>
                                <span style="color: #e0e0e0;">支付宝支付开关</span>
                            </div>
                            <label class="toggle-switch">
                                <input type="checkbox" name="enableAlipay" value="1" <%= IIf(enableAlipay="1", "checked", "") %>>
                                <span class="slider"></span>
                            </label>
                        </div>
                        
                        <div class="form-group">
                            <label>APP ID</label>
                            <input type="text" name="alipayAppId" value="<%= Server.HTMLEncode(GetConfig("AlipayAppId")) %>" placeholder="请输入支付宝APP ID">
                        </div>
                        <div class="form-group">
                            <label>商户号</label>
                            <input type="text" name="alipayMerchantId" value="<%= Server.HTMLEncode(GetConfig("AlipayMerchantId")) %>" placeholder="请输入商户号">
                        </div>
                        <div class="form-group">
                            <label>手续费率 (%)</label>
                            <input type="number" name="alipayFeeRate" value="<%= GetConfigWithDefault("AlipayFeeRate", "0.6") %>" placeholder="0.6" step="0.01" min="0" max="100">
                            <div class="help-text">支付宝手续费率，默认一般为0.6%</div>
                        </div>
                    </div>
                </div>
                
                <!-- 微信配置 -->
                <div class="config-card wechat">
                    <h3 onclick="toggleCard(this)">
                        <i class="fab fa-weixin"></i> 微信支付
                        <span class="status-badge <%= IIf(enableWechat="1", "status-online", "status-offline") %>"><%= IIf(enableWechat="1", "已上线", "已下线") %></span>
                        <i class="fas fa-chevron-down toggle-icon"></i>
                    </h3>
                    <div class="card-content">
                        <div class="switch-group">
                            <div class="switch-label">
                                <i class="fab fa-weixin" style="color: #07c160;"></i>
                                <span style="color: #e0e0e0;">微信支付开关</span>
                            </div>
                            <label class="toggle-switch">
                                <input type="checkbox" name="enableWechat" value="1" <%= IIf(enableWechat="1", "checked", "") %>>
                                <span class="slider"></span>
                            </label>
                        </div>
                        
                        <div class="form-group">
                            <label>APP ID</label>
                            <input type="text" name="wechatAppId" value="<%= Server.HTMLEncode(GetConfig("WechatAppId")) %>" placeholder="请输入微信APP ID">
                        </div>
                        <div class="form-group">
                            <label>商户号 (MCH ID)</label>
                            <input type="text" name="wechatMchId" value="<%= Server.HTMLEncode(GetConfig("WechatMchId")) %>" placeholder="请输入微信商户号">
                        </div>
                        <div class="form-group">
                            <label>手续费率 (%)</label>
                            <input type="number" name="wechatFeeRate" value="<%= GetConfigWithDefault("WechatFeeRate", "0.6") %>" placeholder="0.6" step="0.01" min="0" max="100">
                            <div class="help-text">微信手续费率，默认一般为0.6%</div>
                        </div>
                    </div>
                </div>
                
                <!-- 银行配置 -->
                <div class="config-card bank">
                    <h3 onclick="toggleCard(this)">
                        <i class="fas fa-university"></i> 银行转账
                        <span class="status-badge <%= IIf(enableBank="1", "status-online", "status-offline") %>"><%= IIf(enableBank="1", "已上线", "已下线") %></span>
                        <i class="fas fa-chevron-down toggle-icon"></i>
                    </h3>
                    <div class="card-content">
                        <div class="switch-group">
                            <div class="switch-label">
                                <i class="fas fa-university" style="color: #ff6b6b;"></i>
                                <span style="color: #e0e0e0;">银行转账开关</span>
                            </div>
                            <label class="toggle-switch">
                                <input type="checkbox" name="enableBank" value="1" <%= IIf(enableBank="1", "checked", "") %>>
                                <span class="slider"></span>
                            </label>
                        </div>
                        
                        <div class="form-group">
                            <label>户名</label>
                            <input type="text" name="bankAccountName" value="<%= Server.HTMLEncode(GetConfig("BankAccountName")) %>" placeholder="请输入银行卡户名">
                        </div>
                        <div class="form-group">
                            <label>账号</label>
                            <input type="text" name="bankAccountNo" value="<%= Server.HTMLEncode(GetConfig("BankAccountNo")) %>" placeholder="请输入银行卡号">
                        </div>
                        <div class="form-group">
                            <label>开户银行</label>
                            <input type="text" name="bankName" value="<%= Server.HTMLEncode(GetConfig("BankName")) %>" placeholder="请输入开户银行名称">
                        </div>
                    </div>
                </div>
                
                <!-- Stripe配置 -->
                <div class="config-card stripe">
                    <h3 onclick="toggleCard(this)">
                        <i class="fab fa-stripe"></i> Stripe
                        <span class="status-badge <%= IIf(enableStripe="1", "status-online", "status-offline") %>"><%= IIf(enableStripe="1", "已上线", "已下线") %></span>
                        <i class="fas fa-chevron-down toggle-icon"></i>
                    </h3>
                    <div class="card-content">
                        <div class="switch-group">
                            <div class="switch-label">
                                <i class="fab fa-stripe" style="color: #635bff;"></i>
                                <span style="color: #e0e0e0;">Stripe支付开关</span>
                            </div>
                            <label class="toggle-switch">
                                <input type="checkbox" name="enableStripe" value="1" <%= IIf(enableStripe="1", "checked", "") %>>
                                <span class="slider"></span>
                            </label>
                        </div>
                        
                        <div class="form-group">
                            <label>Publishable Key (公钥)</label>
                            <input type="text" name="stripePublishableKey" value="<%= Server.HTMLEncode(GetConfig("StripePublishableKey")) %>" placeholder="pk_live_... 或 pk_test_...">
                        </div>
                        <div class="form-group">
                            <label>Secret Key (密钥)</label>
                            <input type="password" name="stripeSecretKey" value="<%= Server.HTMLEncode(GetConfig("StripeSecretKey")) %>" placeholder="sk_live_... 或 sk_test_...">
                            <div class="help-text">此密钥仅用于服务器端，不会暴露给客户端</div>
                        </div>
                        <div class="form-group">
                            <label>Webhook Secret</label>
                            <input type="password" name="stripeWebhookSecret" value="<%= Server.HTMLEncode(GetConfig("StripeWebhookSecret")) %>" placeholder="whsec_...">
                            <div class="help-text">用于验证Webhook请求的签名</div>
                        </div>
                        <div class="form-group">
                            <label>手续费率 (%)</label>
                            <input type="number" name="stripeFeeRate" value="<%= GetConfigWithDefault("StripeFeeRate", "2.9") %>" placeholder="2.9" step="0.01" min="0" max="100">
                            <div class="help-text">Stripe标准费率为2.9%</div>
                        </div>
                        <div class="form-group">
                            <label>固定费用 (USD)</label>
                            <input type="number" name="stripeFixedFee" value="<%= GetConfigWithDefault("StripeFixedFee", "0.30") %>" placeholder="0.30" step="0.01" min="0">
                            <div class="help-text">每笔交易固定费用，通常为$0.30</div>
                        </div>
                    </div>
                </div>
                
                <!-- 银联配置 -->
                <div class="config-card unionpay">
                    <h3 onclick="toggleCard(this)">
                        <i class="fas fa-credit-card"></i> 银联支付
                        <span class="status-badge <%= IIf(enableUnionPay="1", "status-online", "status-offline") %>"><%= IIf(enableUnionPay="1", "已上线", "已下线") %></span>
                        <i class="fas fa-chevron-down toggle-icon"></i>
                    </h3>
                    <div class="card-content">
                        <div class="switch-group">
                            <div class="switch-label">
                                <i class="fas fa-credit-card" style="color: #c00;"></i>
                                <span style="color: #e0e0e0;">银联支付开关</span>
                            </div>
                            <label class="toggle-switch">
                                <input type="checkbox" name="enableUnionPay" value="1" <%= IIf(enableUnionPay="1", "checked", "") %>>
                                <span class="slider"></span>
                            </label>
                        </div>
                        
                        <div class="form-group">
                            <label>商户号</label>
                            <input type="text" name="unionPayMerchantId" value="<%= Server.HTMLEncode(GetConfig("UnionPayMerchantId")) %>" placeholder="请输入银联商户号">
                        </div>
                        <div class="form-group">
                            <label>证书路径</label>
                            <input type="text" name="unionPayCertPath" value="<%= Server.HTMLEncode(GetConfig("UnionPayCertPath")) %>" placeholder="/certs/unionpay/...">
                            <div class="help-text">银联证书文件在服务器的绝对路径或相对路径</div>
                        </div>
                        <div class="form-group">
                            <label>手续费率 (%)</label>
                            <input type="number" name="unionPayFeeRate" value="<%= GetConfigWithDefault("UnionPayFeeRate", "0.6") %>" placeholder="0.6" step="0.01" min="0" max="100">
                            <div class="help-text">银联标准费率一般为0.6%</div>
                        </div>
                    </div>
                </div>
                
                <!-- PayPal配置 -->
                <div class="config-card paypal">
                    <h3 onclick="toggleCard(this)">
                        <i class="fab fa-paypal"></i> PayPal
                        <span class="status-badge <%= IIf(enablePayPal="1", "status-online", "status-offline") %>"><%= IIf(enablePayPal="1", "已上线", "已下线") %></span>
                        <i class="fas fa-chevron-down toggle-icon"></i>
                    </h3>
                    <div class="card-content">
                        <div class="switch-group">
                            <div class="switch-label">
                                <i class="fab fa-paypal" style="color: #003087;"></i>
                                <span style="color: #e0e0e0;">PayPal支付开关</span>
                            </div>
                            <label class="toggle-switch">
                                <input type="checkbox" name="enablePayPal" value="1" <%= IIf(enablePayPal="1", "checked", "") %>>
                                <span class="slider"></span>
                            </label>
                        </div>
                        
                        <div class="switch-group" style="border-bottom: none; padding-bottom: 0;">
                            <div class="switch-label">
                                <i class="fas fa-flask" style="color: #ffa726;"></i>
                                <span style="color: #e0e0e0;">沙箱模式</span>
                            </div>
                            <label class="toggle-switch">
                                <input type="checkbox" name="paypalSandbox" value="1" <%= IIf(GetConfig("PayPalSandbox")="1", "checked", "") %>>
                                <span class="slider"></span>
                            </label>
                        </div>
                        <div class="help-text" style="margin-bottom: 20px;">启用后使用PayPal沙箱环境进行测试</div>
                        
                        <div class="form-group">
                            <label>Client ID</label>
                            <input type="text" name="paypalClientId" value="<%= Server.HTMLEncode(GetConfig("PayPalClientId")) %>" placeholder="请输入PayPal Client ID">
                        </div>
                        <div class="form-group">
                            <label>Secret (密钥)</label>
                            <input type="password" name="paypalSecret" value="<%= Server.HTMLEncode(GetConfig("PayPalSecret")) %>" placeholder="请输入PayPal Secret">
                        </div>
                        <div class="form-group">
                            <label>手续费率 (%)</label>
                            <input type="number" name="paypalFeeRate" value="<%= GetConfigWithDefault("PayPalFeeRate", "4.4") %>" placeholder="4.4" step="0.01" min="0" max="100">
                            <div class="help-text">PayPal标准费率为4.4%</div>
                        </div>
                        <div class="form-group">
                            <label>固定费用 (USD)</label>
                            <input type="number" name="paypalFixedFee" value="<%= GetConfigWithDefault("PayPalFixedFee", "0.30") %>" placeholder="0.30" step="0.01" min="0">
                            <div class="help-text">每笔交易固定费用，通常为$0.30</div>
                        </div>
                    </div>
                </div>
                
                <!-- 支付统计 -->
                <div class="config-card">
                    <h3 onclick="toggleCard(this)">
                        <i class="fas fa-chart-bar"></i> 支付统计
                        <i class="fas fa-chevron-down toggle-icon"></i>
                    </h3>
                    <div class="card-content">
                        <div style="padding: 20px; text-align: center;">
                            <div style="font-size: 32px; font-weight: bold; color: #4CAF50; margin-bottom: 10px;">
                                <%= GetScalar("SELECT COUNT(*) FROM Orders WHERE Status = 'Paid'") %>
                            </div>
                            <div style="color: #888;">成功支付订单</div>
                            <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #3a3a4a;">
                                <div style="font-size: 24px; font-weight: bold; color: #e0e0e0;">
                                    ¥<%= FormatNumber(CDbl("0" & GetScalar("SELECT CAST(IIF(SUM(TotalAmount) IS NULL, 0, SUM(TotalAmount)) AS FLOAT) FROM Orders WHERE Status = 'Paid'")), 2) %>
                                </div>
                                <div style="color: #888; font-size: 12px;">累计支付金额</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- 费率对比一览表 -->
            <div class="fee-comparison-section">
                <h3><i class="fas fa-table"></i> 费率对比一览表</h3>
                <table class="fee-table">
                    <thead>
                        <tr>
                            <th>支付方式</th>
                            <th>状态</th>
                            <th>费率</th>
                            <th>固定费</th>
                            <th>预估100元订单手续费</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>
                                <div class="method-name">
                                    <div class="method-icon icon-alipay"><i class="fab fa-alipay"></i></div>
                                    <span>支付宝</span>
                                </div>
                            </td>
                            <td><span class="status-badge <%= IIf(enableAlipay="1", "status-online", "status-offline") %>"><%= IIf(enableAlipay="1", "已启用", "已禁用") %></span></td>
                            <td><%= GetConfigWithDefault("AlipayFeeRate", "0.6") %>%</td>
                            <td>¥0.00</td>
                            <td>¥<%= FormatNumber(100 * CDbl(GetConfigWithDefault("AlipayFeeRate", "0.6")) / 100, 2) %></td>
                        </tr>
                        <tr>
                            <td>
                                <div class="method-name">
                                    <div class="method-icon icon-wechat"><i class="fab fa-weixin"></i></div>
                                    <span>微信支付</span>
                                </div>
                            </td>
                            <td><span class="status-badge <%= IIf(enableWechat="1", "status-online", "status-offline") %>"><%= IIf(enableWechat="1", "已启用", "已禁用") %></span></td>
                            <td><%= GetConfigWithDefault("WechatFeeRate", "0.6") %>%</td>
                            <td>¥0.00</td>
                            <td>¥<%= FormatNumber(100 * CDbl(GetConfigWithDefault("WechatFeeRate", "0.6")) / 100, 2) %></td>
                        </tr>
                        <tr>
                            <td>
                                <div class="method-name">
                                    <div class="method-icon icon-bank"><i class="fas fa-university"></i></div>
                                    <span>银行转账</span>
                                </div>
                            </td>
                            <td><span class="status-badge <%= IIf(enableBank="1", "status-online", "status-offline") %>"><%= IIf(enableBank="1", "已启用", "已禁用") %></span></td>
                            <td>0%</td>
                            <td>¥0.00</td>
                            <td>¥0.00</td>
                        </tr>
                        <tr>
                            <td>
                                <div class="method-name">
                                    <div class="method-icon icon-stripe"><i class="fab fa-stripe"></i></div>
                                    <span>Stripe</span>
                                </div>
                            </td>
                            <td><span class="status-badge <%= IIf(enableStripe="1", "status-online", "status-offline") %>"><%= IIf(enableStripe="1", "已启用", "已禁用") %></span></td>
                            <td><%= GetConfigWithDefault("StripeFeeRate", "2.9") %>%</td>
                            <td>$<%= GetConfigWithDefault("StripeFixedFee", "0.30") %></td>
                            <td>$<%= FormatNumber(100 * CDbl(GetConfigWithDefault("StripeFeeRate", "2.9")) / 100 + CDbl(GetConfigWithDefault("StripeFixedFee", "0.30")), 2) %></td>
                        </tr>
                        <tr>
                            <td>
                                <div class="method-name">
                                    <div class="method-icon icon-unionpay"><i class="fas fa-credit-card"></i></div>
                                    <span>银联</span>
                                </div>
                            </td>
                            <td><span class="status-badge <%= IIf(enableUnionPay="1", "status-online", "status-offline") %>"><%= IIf(enableUnionPay="1", "已启用", "已禁用") %></span></td>
                            <td><%= GetConfigWithDefault("UnionPayFeeRate", "0.6") %>%</td>
                            <td>¥0.00</td>
                            <td>¥<%= FormatNumber(100 * CDbl(GetConfigWithDefault("UnionPayFeeRate", "0.6")) / 100, 2) %></td>
                        </tr>
                        <tr>
                            <td>
                                <div class="method-name">
                                    <div class="method-icon icon-paypal"><i class="fab fa-paypal"></i></div>
                                    <span>PayPal</span>
                                </div>
                            </td>
                            <td><span class="status-badge <%= IIf(enablePayPal="1", "status-online", "status-offline") %>"><%= IIf(enablePayPal="1", "已启用", "已禁用") %></span></td>
                            <td><%= GetConfigWithDefault("PayPalFeeRate", "4.4") %>%</td>
                            <td>$<%= GetConfigWithDefault("PayPalFixedFee", "0.30") %></td>
                            <td>$<%= FormatNumber(100 * CDbl(GetConfigWithDefault("PayPalFeeRate", "4.4")) / 100 + CDbl(GetConfigWithDefault("PayPalFixedFee", "0.30")), 2) %></td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <div style="text-align: center; margin-top: 30px;">
                <% If Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then %>
                <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存配置</button>
                <% Else %>
                <div class="alert alert-warning">
                    <i class="fas fa-lock"></i> 您当前为财务专员，仅可查看不可编辑支付参数
                </div>
                <% End If %>
            </div>
        </form>
    </div>
    
    <script>
        // 卡片折叠展开功能
        function toggleCard(header) {
            var card = header.parentElement;
            card.classList.toggle('collapsed');
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
