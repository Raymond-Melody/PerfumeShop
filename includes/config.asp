<%
' ============================================
' 网站配置文件
' ============================================

' 网站基本信息
Const SITE_NAME = "香氛定制"
Const SITE_URL = "http://localhost"
Const SITE_EMAIL = "contact@perfumeshop.com"
Const SITE_PHONE = "400-888-8888"

' 系统版本
Const SYS_VERSION = "V9.0"
Const SYS_VERSION_NAME = "香氛定制系统V9.0 - 性能优化与体验升级"

' 数据库连接配置已迁移至 connection.asp (SQL Server)
' Const DB_CONNECTION_STRING 已废弃，连接字符串在 connection.asp 中定义

' 分页设置
Const PAGE_SIZE = 12

' 运费设置
Const FREE_SHIPPING_AMOUNT = 299
Const SHIPPING_FEE = 15

' 图片路径
Const IMAGE_PATH = "/images/"
Const DEFAULT_PRODUCT_IMAGE = "/images/default-product.svg"
Const DEFAULT_AVATAR = "/images/default-avatar.svg"

' 图片上传路径
Const UPLOAD_PATH_PRODUCTS = "/images/products/"
Const UPLOAD_PATH_NOTES = "/images/notes/"
Const UPLOAD_PATH_BOTTLES = "/images/bottles/"
Const UPLOAD_PATH_AVATARS = "/images/avatars/"
Const UPLOAD_PATH_DEFAULT = "/images/uploads/"

' Session超时时间（分钟）
Session.Timeout = 60

' === V7 安全响应头 ===
Response.AddHeader "X-Content-Type-Options", "nosniff"
Response.AddHeader "X-Frame-Options", "SAMEORIGIN"
Response.AddHeader "X-XSS-Protection", "1; mode=block"
Response.AddHeader "Referrer-Policy", "strict-origin-when-cross-origin"
%>