<%
' ============================================
' V18.0 支付配置中心 (Payment Config)
' 集中管理所有支付渠道的配置项
' 各渠道密钥/商户号在此配置，后续填入
' 真实凭证即可启用对应支付方式
' ============================================

' ---- 支付方式编码常量 (V18修复: 此前未定义导致显示"未知") ----
Const PAYMENT_METHOD_WECHAT = "1"     ' 微信支付
Const PAYMENT_METHOD_ALIPAY = "2"     ' 支付宝
Const PAYMENT_METHOD_PAYPAL = "3"     ' PayPal
Const PAYMENT_METHOD_COD = "4"        ' 货到付款

' ---- 支付方式开关 ----
Const PAYMENT_WECHAT_ENABLED = False      ' 微信支付（需填入商户凭证后启用）
Const PAYMENT_ALIPAY_ENABLED = False      ' 支付宝（需填入商户凭证后启用）
Const PAYMENT_BANKCARD_ENABLED = True     ' 银行卡/信用卡（默认启用）

' ---- 微信支付配置（JSAPI 手机网页支付）----
' 申请地址: https://pay.weixin.qq.com
Const WECHAT_APPID = "wx_YOUR_APPID_HERE"
Const WECHAT_MCHID = "YOUR_MERCHANT_ID"
Const WECHAT_API_KEY = "YOUR_API_KEY_V3"
Const WECHAT_API_CERT_PATH = "/certs/apiclient_cert.pem"
Const WECHAT_API_KEY_PATH = "/certs/apiclient_key.pem"
Const WECHAT_NOTIFY_URL = "https://yourdomain.com/payment_callback.asp?channel=wechat"
Const WECHAT_REFUND_URL = "https://yourdomain.com/admin/finance/refund.asp"

' ---- 支付宝配置（手机网页支付）----
' 申请地址: https://open.alipay.com
Const ALIPAY_APPID = "YOUR_ALIPAY_APPID"
Const ALIPAY_PRIVATE_KEY = "YOUR_PRIVATE_KEY"
Const ALIPAY_PUBLIC_KEY = "YOUR_ALIPAY_PUBLIC_KEY"
Const ALIPAY_NOTIFY_URL = "https://yourdomain.com/payment_callback.asp?channel=alipay"
Const ALIPAY_RETURN_URL = "https://yourdomain.com/order_success.asp"

' ---- 通用支付配置 ----
Const PAYMENT_CURRENCY = "CNY"             ' 货币代码
Const PAYMENT_TIMEOUT_MINUTES = 30         ' 支付超时（分钟）
Const PAYMENT_AUTO_CANCEL = True           ' 超时自动取消订单
Const PAYMENT_REMEMBER_METHOD = True       ' 是否允许记住支付方式

' ============================================
' 获取可用支付方式列表（用于结算页展示）
' 返回：Dictionary，Key=渠道代码, Value=显示名称
' ============================================
Function Payment_GetAvailableMethods()
    Dim methods
    Set methods = Server.CreateObject("Scripting.Dictionary")
    
    If PAYMENT_BANKCARD_ENABLED Then
        methods.Add "bankcard", "银行卡支付"
    End If
    
    If PAYMENT_WECHAT_ENABLED Then
        methods.Add "wechat", "微信支付"
    End If
    
    If PAYMENT_ALIPAY_ENABLED Then
        methods.Add "alipay", "支付宝"
    End If
    
    Set Payment_GetAvailableMethods = methods
End Function

' ============================================
' 获取支付方式图标
' 返回：FontAwesome 图标类名
' ============================================
Function Payment_GetIcon(channel)
    Select Case LCase(channel)
        Case "wechat":    Payment_GetIcon = "fa-weixin"
        Case "alipay":    Payment_GetIcon = "fa-alipay"
        Case "bankcard":  Payment_GetIcon = "fa-credit-card"
        Case Else:        Payment_GetIcon = "fa-money-bill-wave"
    End Select
End Function

' ============================================
' 获取支付方式颜色
' 返回：CSS 颜色代码
' ============================================
Function Payment_GetColor(channel)
    Select Case LCase(channel)
        Case "wechat":    Payment_GetColor = "#07C160"
        Case "alipay":    Payment_GetColor = "#1677FF"
        Case "bankcard":  Payment_GetColor = "#8B4513"
        Case Else:        Payment_GetColor = "#666"
    End Select
End Function
%>
