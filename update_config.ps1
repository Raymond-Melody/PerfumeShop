$filePath = 'f:\网站制作\网站\网站二\includes\config.asp'
$content = [IO.File]::ReadAllText($filePath, [Text.Encoding]::UTF8)

$oldSection = @"
' ============================================
' V16 Feature Flags - 新功能默认关闭，验证后逐项开启
' ============================================
Const FEATURE_MSOLEDBSQL = True         ' P0: MSOLEDBSQL (需先启用SQL Server TCP/IP协议)
                                        ' 原因：YOURPERFUME实例仅监听Shared Memory，
                                        ' MSOLEDBSQL要求TCP。解决：SQL Server配置管理器
                                        ' → YOURPERFUME协议 → 启用TCP/IP → 重启服务
                                        ' 回退：若MSOLEDBSQL不可用，自动回退SQLOLEDB
                                        ' 脚本: database/enable_tcpip_sqlserver.ps1
Const FEATURE_DAL_ENABLED = True        ' P0: 启用统一数据访问层 (V16激活)
Const FEATURE_PASSWORD_V3 = True        ' P0: 启用SHA-512密码哈希V3 (V16激活)
Const FEATURE_STRUCTURED_LOGGING = True ' P1: 启用结构化日志 (无外部依赖)
Const FEATURE_API_V1 = True             ' P1: 启用API v1统一响应格式 (V16激活)
Const FEATURE_CACHE_MANAGER = True      ' P1: 启用缓存管理器 (无外部依赖)
Const FEATURE_SSE_NOTIFICATIONS = True  ' P2: 启用SSE实时通知 (V16激活)
Const FEATURE_EMAIL_NOTIFICATIONS = True ' P2: 启用邮件通知 (V16激活)
Const FEATURE_ANALYTICS_DASHBOARD = True ' P2: 启用数据分析仪表盘 (V16激活)
Const FEATURE_PWA_ENHANCED = True       ' P2: 启用PWA增强 (V16激活)
Const FEATURE_I18N = True              ' P2: 启用国际化 (V17激活)
Const FEATURE_API_AUTH = True           ' P0: 启用API认证 (V18新增)
Const FEATURE_RATE_LIMITER = True       ' P0: 启用速率限制 (V18新增)
Const FEATURE_GDPR_COMPLIANCE = True    ' P0: 启用GDPR隐私合规 (V18新增)
Const FEATURE_AI_RECOMMENDATIONS = True ' P1: 启用AI推荐引擎 (V18新增)
Const FEATURE_AI_FRAGRANCE_MATCH = True  ' P1: 启用智能香氛匹配 (V18新增)
Const FEATURE_AI_SEARCH = True          ' P1: 启用智能搜索升级 (V18新增)
Const FEATURE_AI_CHATBOT = True         ' P1: 启用智能客服机器人 (V18新增)
Const FEATURE_MEMBER_TIERS = True     ' P2: 启用会员等级体系 (V18新增)
Const FEATURE_POINTS_SYSTEM = True   ' P2: 启用积分与奖励系统 (V18新增)
Const FEATURE_COUPON_SYSTEM = True   ' P2: 启用优惠券与促销引擎 (V18新增)
Const FEATURE_FLASH_SALE = True     ' P2: 启用限时秒杀引擎 (V18新增)
Const FEATURE_GROUP_BUY = True      ' P2: 启用拼团活动引擎 (V18新增)
Const FEATURE_SUBSCRIPTION = True   ' P2: 启用订阅制香氛盒 (V18新增)
Const FEATURE_COMMUNITY = True      ' P2: 启用会员社区UGC (V18新增)
"@

$newSection = @"
' ============================================
' V16 Feature Flags - 新功能默认关闭，验证后逐项开启
' ============================================
Const FEATURE_MSOLEDBSQL = True         ' P0: MSOLEDBSQL (需先启用SQL Server TCP/IP协议)
                                        ' 原因：YOURPERFUME实例仅监听Shared Memory，
                                        ' MSOLEDBSQL要求TCP。解决：SQL Server配置管理器
                                        ' → YOURPERFUME协议 → 启用TCP/IP → 重启服务
                                        ' 回退：若MSOLEDBSQL不可用，自动回退SQLOLEDB
                                        ' 脚本: database/enable_tcpip_sqlserver.ps1
Const FEATURE_DAL_ENABLED = True        ' P0: 启用统一数据访问层 (V16激活)
Const FEATURE_PASSWORD_V3 = True        ' P0: 启用SHA-512密码哈希V3 (V16激活)
Const FEATURE_STRUCTURED_LOGGING = True ' P1: 启用结构化日志 (无外部依赖)
Const FEATURE_API_V1 = True             ' P1: 启用API v1统一响应格式 (V16激活)
Const FEATURE_CACHE_MANAGER = True      ' P1: 启用缓存管理器 (无外部依赖)
Const FEATURE_SSE_NOTIFICATIONS = True  ' P2: 启用SSE实时通知 (V16激活)
Const FEATURE_EMAIL_NOTIFICATIONS = True ' P2: 启用邮件通知 (V16激活)
Const FEATURE_ANALYTICS_DASHBOARD = True ' P2: 启用数据分析仪表盘 (V16激活)
Const FEATURE_PWA_ENHANCED = True       ' P2: 启用PWA增强 (V16激活)
Const FEATURE_I18N = True              ' P2: 启用国际化 (V17激活)

' ============================================
' V18 Feature Flags - 渐进式架构现代化 + 运营功能
' 说明: 所有V18 Flags默认开启，可按需关闭回退
' 依赖: V18数据库Schema（v18_*.sql）需先执行
' ============================================
Const FEATURE_API_AUTH = True           ' P0: 启用API认证 (V18新增)
                                        ' 依赖: includes/api_auth.asp + includes/api_guard.asp
                                        ' 功能: Session认证 + API Key/HMAC-SHA256签名验证
Const FEATURE_RATE_LIMITER = True       ' P0: 启用速率限制 (V18新增)
                                        ' 依赖: includes/rate_limiter.asp + includes/api_guard.asp
                                        ' 功能: 令牌桶算法 60req/60s，超限返回429
Const FEATURE_GDPR_COMPLIANCE = True    ' P0: 启用GDPR隐私合规 (V18新增)
                                        ' 依赖: api/cookie_consent.asp + 隐私政策
                                        ' 功能: Cookie同意弹窗 + 数据导出/删除
Const FEATURE_AI_RECOMMENDATIONS = True ' P1: 启用AI推荐引擎 (V18新增)
                                        ' 依赖: includes/recommendation_engine.asp
                                        ' 功能: 个性化推荐 + "猜你喜欢"模块
Const FEATURE_AI_FRAGRANCE_MATCH = True  ' P1: 启用智能香氛匹配 (V18新增)
                                        ' 依赖: api/fragrance_match.asp + fragrance_quiz.asp
                                        ' 功能: 香氛测试问答 → 智能匹配推荐
Const FEATURE_AI_SEARCH = True          ' P1: 启用智能搜索升级 (V18新增)
                                        ' 依赖: api/search_suggestions.asp
                                        ' 功能: 模糊搜索 + 自动补全 + 加权排序
Const FEATURE_AI_CHATBOT = True         ' P1: 启用智能客服机器人 (V18新增)
                                        ' 依赖: api/chatbot.asp
                                        ' 功能: 智能客服 + 订单查询 + 产品推荐
Const FEATURE_MEMBER_TIERS = True     ' P2: 启用会员等级体系 (V18新增)
                                        ' 依赖: database/v18_member_tiers.sql
                                        ' 功能: 青铜/白银/黄金/钻石等级权益
Const FEATURE_POINTS_SYSTEM = True   ' P2: 启用积分与奖励系统 (V18新增)
                                        ' 依赖: database/v18_points_system.sql
                                        ' 功能: 积分获取/消耗 + 奖励兑换
Const FEATURE_COUPON_SYSTEM = True   ' P2: 启用优惠券与促销引擎 (V18新增)
                                        ' 依赖: database/v18_coupon_system.sql + api/coupon_validate.asp
                                        ' 功能: 优惠券发放/核销 + 满减/折扣/免邮
Const FEATURE_FLASH_SALE = True     ' P2: 启用限时秒杀引擎 (V18新增)
                                        ' 依赖: database/v18_flash_group_activities.sql + flash_sale.asp
                                        ' 功能: 限时秒杀 + 倒计时 + 库存扣减
Const FEATURE_GROUP_BUY = True      ' P2: 启用拼团活动引擎 (V18新增)
                                        ' 依赖: database/v18_flash_group_activities.sql + group_buy.asp
                                        ' 功能: 拼团发起/参与 + 人数进度
Const FEATURE_SUBSCRIPTION = True   ' P2: 启用订阅制香氛盒 (V18新增)
                                        ' 依赖: database/v18_subscription.sql + subscribe.asp
                                        ' 功能: 月/季/年订阅计划 + 自动续费
Const FEATURE_COMMUNITY = True      ' P2: 启用会员社区UGC (V18新增)
                                        ' 依赖: database/v18_community_ugc.sql + community.asp
                                        ' 功能: 帖子/评论/点赞 + 香评分享
' ============================================
"@

if ($content.Contains($oldSection)) {
    $newContent = $content.Replace($oldSection, $newSection)
    [IO.File]::WriteAllText($filePath, $newContent, [Text.Encoding]::UTF8)
    Write-Host "config.asp V18 Feature Flags comments updated successfully"
} else {
    Write-Host "WARNING: Could not find exact section to replace"
    # Try line-by-line check
    $lines = $content -split "`r?`n"
    Write-Host "File has $($lines.Count) lines"
    $foundV18 = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "FEATURE_API_AUTH") {
            Write-Host "Found FEATURE_API_AUTH at line $($i+1)"
            $foundV18 = $true
            break
        }
    }
    if (-not $foundV18) {
        Write-Host "FEATURE_API_AUTH not found in file"
    }
}
