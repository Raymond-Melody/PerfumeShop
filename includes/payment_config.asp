<%
' ============================================
' 支付配置文件
' ============================================

' 微信支付配置（需要实际商户信息）
Const WECHAT_PAY_ENABLED = True      ' 是否启用微信支付
Const WECHAT_PAY_APPID = ""       ' 您的微信支付APPID
Const WECHAT_PAY_MCH_ID = ""      ' 您的微信支付商户号
Const WECHAT_PAY_KEY = ""         ' 您的微信支付API密钥
Const WECHAT_PAY_NOTIFY_URL = ""  ' 微信支付异步通知地址

' 支付宝配置（需要实际商户信息）
Const ALIPAY_ENABLED = True         ' 是否启用支付宝
Const ALIPAY_APP_ID = ""          ' 您的支付宝APP ID
Const ALIPAY_PRIVATE_KEY = ""     ' 您的支付宝私钥
Const ALIPAY_PUBLIC_KEY = ""      ' 支付宝公钥
Const ALIPAY_NOTIFY_URL = ""      ' 支付宝异步通知地址

' PayPal配置（需要实际商户信息）
Const PAYPAL_ENABLED = True         ' 是否启用PayPal
Const PAYPAL_CLIENT_ID = ""       ' 您的PayPal客户端ID
Const PAYPAL_SECRET = ""          ' 您的PayPal密钥
Const PAYPAL_MODE = "sandbox"     ' PayPal模式: sandbox 或 live
Const PAYPAL_NOTIFY_URL = ""      ' PayPal异步通知地址

' 货到付款配置
Const COD_ENABLED = True            ' 是否启用货到付款

' 支付状态常量
Const PAYMENT_STATUS_PENDING = 0    ' 待支付
Const PAYMENT_STATUS_PAID = 1       ' 已支付
Const PAYMENT_STATUS_FAILED = 2     ' 支付失败
Const PAYMENT_STATUS_REFUNDED = 3   ' 已退款

' 支付方式常量
Const PAYMENT_METHOD_WECHAT = 1     ' 微信支付
Const PAYMENT_METHOD_ALIPAY = 2     ' 支付宝
Const PAYMENT_METHOD_PAYPAL = 3     ' PayPal
Const PAYMENT_METHOD_COD = 4        ' 货到付款

' 支付相关URL
Const WECHAT_PAY_API_URL = "https://api.mch.weixin.qq.com/pay/unifiedorder"
Const ALIPAY_GATEWAY_URL = "https://openapi.alipay.com/gateway.do"
Const PAYPAL_API_URL_SANDBOX = "https://api.sandbox.paypal.com/v1/payments/payment"
Const PAYPAL_API_URL_LIVE = "https://api.paypal.com/v1/payments/payment"
%>