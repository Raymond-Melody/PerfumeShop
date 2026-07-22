using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Security.Claims;

namespace PerfumeShop.Shared.Auth;

/// <summary>
/// 认证桥接中间件 — 核心 POC
/// 读取 V19_AUTH_TOKEN Cookie（HttpOnly, Secure），查 AuthTokens 表验证，
/// 验证通过后创建 ClaimsPrincipal 注入 HttpContext.User，
/// 同时签发 ASP.NET Core Cookie Authentication 票据以维持后续请求身份。
/// </summary>
public class AuthBridgeMiddleware
{
    private readonly RequestDelegate _next;
    private readonly ILogger<AuthBridgeMiddleware> _logger;

    public AuthBridgeMiddleware(RequestDelegate next, ILogger<AuthBridgeMiddleware> logger)
    {
        _next = next;
        _logger = logger;
    }

    public async Task InvokeAsync(HttpContext context, IOptions<AuthBridgeOptions> options, IAuthTokenStore tokenStore)
    {
        var opts = options.Value;

        if (!opts.EnableCookieBridge)
        {
            await _next(context);
            return;
        }

        // 如果已有认证用户（Cookie Authentication 已处理），跳过
        if (context.User?.Identity?.IsAuthenticated == true)
        {
            await _next(context);
            return;
        }

        // 读取 V19_AUTH_TOKEN Cookie
        var rawToken = context.Request.Cookies[opts.CookieName];
        if (string.IsNullOrWhiteSpace(rawToken))
        {
            await _next(context);
            return;
        }

        try
        {
            // 计算 Token 的 SHA-256 哈希用于查表
            var tokenHash = ComputeTokenHash(rawToken);

            // 通过 IAuthTokenStore 验证 Token
            var principal = await tokenStore.ValidateTokenAsync(tokenHash);
            if (principal != null)
            {
                // 注入 HttpContext.User
                context.User = principal;

                // 签发 ASP.NET Core Cookie 票据（后续请求无需再查 Token 表）
                await context.SignInAsync(
                    CookieAuthenticationDefaults.AuthenticationScheme,
                    principal,
                    new AuthenticationProperties
                    {
                        IsPersistent = true,
                        ExpiresUtc = DateTimeOffset.UtcNow.AddHours(8),
                        IssuedUtc = DateTimeOffset.UtcNow
                    });

                _logger.LogDebug("AuthBridge: Token 验证成功, UserId={UserId}",
                    principal.FindFirstValue(ClaimTypes.NameIdentifier));
            }
            else
            {
                _logger.LogDebug("AuthBridge: Token 验证失败或已过期");
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "AuthBridge: Token 验证异常");
        }

        await _next(context);
    }

    /// <summary>
    /// 计算 Token 的 SHA-256 哈希（与 V18 侧写入时保持一致）
    /// </summary>
    public static string ComputeTokenHash(string rawToken)
    {
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        var bytes = System.Text.Encoding.UTF8.GetBytes(rawToken);
        var hash = sha256.ComputeHash(bytes);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }
}

/// <summary>扩展方法：注册 AuthBridge 中间件</summary>
public static class AuthBridgeMiddlewareExtensions
{
    public static IApplicationBuilder UseAuthBridge(this IApplicationBuilder builder)
    {
        return builder.UseMiddleware<AuthBridgeMiddleware>();
    }
}
