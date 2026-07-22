namespace PerfumeShop.Shared.Auth;

/// <summary>
/// AuthBridge 中间件配置选项
/// 对应 V18 includes/api_auth.asp 的双模式认证配置
/// </summary>
public class AuthBridgeOptions
{
    /// <summary>认证配置节名称</summary>
    public const string SectionName = "Auth";

    /// <summary>是否启用 Cookie 桥接认证（V18→V19 Session 互通）</summary>
    public bool EnableCookieBridge { get; set; } = true;

    /// <summary>是否启用 API Key 认证</summary>
    public bool EnableApiKey { get; set; } = true;

    /// <summary>V19 认证 Cookie 名称</summary>
    public string CookieName { get; set; } = "V19_AUTH_TOKEN";

    /// <summary>CSRF 防护开关</summary>
    public bool EnableCsrf { get; set; } = true;

    /// <summary>CSRF Token Header 名称</summary>
    public string CsrfHeaderName { get; set; } = "X-CSRF-Token";

    /// <summary>CSRF Token Form 字段名称</summary>
    public string CsrfFormFieldName { get; set; } = "__RequestVerificationToken";

    /// <summary>跳过 CSRF 校验的公开端点路径前缀</summary>
    public HashSet<string> CsrfExemptPaths { get; set; } = new(StringComparer.OrdinalIgnoreCase)
    {
        "/api/health",
        "/api/metrics",
        "/api/auth/login",
        "/api/auth/register",
        "/api/auth/forgot-password",
        "/swagger",
        "/api/notifications/stream"
    };
}
