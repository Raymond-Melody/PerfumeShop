using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PerfumeShop.Shared.Auth;
using System.Net;
using System.Security.Claims;

namespace PerfumeShop.IntegrationTests.Auth;

/// <summary>
/// AuthBridgeMiddleware 集成测试
/// 使用 TestServer + 内存 Mock IAuthTokenStore 验证双系统 Session 互通 POC
/// </summary>
public class AuthBridgeMiddlewareTests
{
    // ====== 测试数据 ======
    private const string ValidToken = "valid-token-abc123";
    private const string ExpiredToken = "expired-token-xyz789";
    private const string InactiveToken = "inactive-token-def456";
    private const string AdminToken = "admin-token-ghi012";
    private const string InvalidToken = "invalid-token-xxx999";
    private const int TestUserId = 42;
    private const string TestUsername = "testuser";
    private const string TestEmail = "test@example.com";

    /// <summary>
    /// 用例 1：有效 Token 识别用户身份
    /// </summary>
    [Fact]
    public async Task ValidToken_SetsAuthenticatedUser()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=" + ValidToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains(TestUserId.ToString(), body);
        Assert.Contains(TestUsername, body);
        Assert.Contains("True", body); // IsAuthenticated
    }

    /// <summary>
    /// 用例 2：过期 Token 不设置用户（返回匿名）
    /// </summary>
    [Fact]
    public async Task ExpiredToken_DoesNotSetUser()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=" + ExpiredToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("False", body); // IsAuthenticated = false
    }

    /// <summary>
    /// 用例 3：缺失 Cookie 允许匿名访问
    /// </summary>
    [Fact]
    public async Task MissingCookie_AllowsAnonymousAccess()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        // 不设置任何 Cookie
        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("False", body); // IsAuthenticated = false
        Assert.Contains("Anonymous", body);
    }

    /// <summary>
    /// 用例 4：IsActive=false 的 Token 被拒绝
    /// </summary>
    [Fact]
    public async Task InactiveToken_DoesNotSetUser()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=" + InactiveToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("False", body); // IsAuthenticated = false
    }

    /// <summary>
    /// 用例 5：Source 区分 UserLogin/AdminLogin（AdminLogin 带 Admin 角色）
    /// </summary>
    [Fact]
    public async Task AdminLoginToken_HasAdminRole()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=" + AdminToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("True", body); // IsAuthenticated
        Assert.Contains("Admin", body); // Role
    }

    /// <summary>
    /// 用例 6：无效 Token（不存在于数据库）返回匿名
    /// </summary>
    [Fact]
    public async Task InvalidToken_DoesNotSetUser()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=" + InvalidToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("False", body);
    }

    /// <summary>
    /// 用例 7：AuthBridge 关闭时跳过 Token 验证
    /// </summary>
    [Fact]
    public async Task AuthBridgeDisabled_SkipsValidation()
    {
        using var server = CreateTestServer(enableCookieBridge: false);
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=" + ValidToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        // 即使 Token 有效，AuthBridge 关闭后不设置用户
        Assert.Contains("False", body);
    }

    /// <summary>
    /// 用例 8：自定义 Cookie 名称生效
    /// </summary>
    [Fact]
    public async Task CustomCookieName_Works()
    {
        using var server = CreateTestServer(cookieName: "MY_CUSTOM_TOKEN");
        var client = server.CreateClient();

        // 使用自定义 Cookie 名
        client.DefaultRequestHeaders.Add("Cookie", "MY_CUSTOM_TOKEN=" + ValidToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("True", body); // IsAuthenticated
    }

    /// <summary>
    /// 用例 9：空 Token 字符串不触发验证
    /// </summary>
    [Fact]
    public async Task EmptyToken_DoesNotTriggerValidation()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=");

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("False", body);
    }

    /// <summary>
    /// 用例 10：UserLogin Source 返回 User 角色（非 Admin）
    /// </summary>
    [Fact]
    public async Task UserLoginToken_HasUserRole()
    {
        using var server = CreateTestServer();
        var client = server.CreateClient();

        client.DefaultRequestHeaders.Add("Cookie", "V19_AUTH_TOKEN=" + ValidToken);

        var response = await client.GetAsync("/api/test-auth");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("True", body);
        Assert.Contains("User", body); // Role = User (not Admin)
    }

    // ====== 辅助方法 ======

    private static TestServer CreateTestServer(
        bool enableCookieBridge = true,
        string cookieName = "V19_AUTH_TOKEN")
    {
        var builder = new WebHostBuilder()
            .ConfigureServices(services =>
            {
                services.AddRouting();
                services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
                    .AddCookie(options =>
                    {
                        options.Cookie.Name = "V19_AUTH";
                        options.Events.OnRedirectToLogin = ctx =>
                        {
                            ctx.Response.StatusCode = 401;
                            return Task.CompletedTask;
                        };
                    });

                services.Configure<AuthBridgeOptions>(opts =>
                {
                    opts.EnableCookieBridge = enableCookieBridge;
                    opts.CookieName = cookieName;
                });

                services.AddSingleton<IAuthTokenStore, MockAuthTokenStore>();
            })
            .Configure(app =>
            {
                app.UseAuthentication();
                app.UseAuthBridge();

                app.Run(async context =>
                {
                    if (context.Request.Path.StartsWithSegments("/api/test-auth"))
                    {
                        var user = context.User;
                        var isAuth = user?.Identity?.IsAuthenticated ?? false;
                        var userId = user?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "Anonymous";
                        var username = user?.FindFirstValue(ClaimTypes.Name) ?? "Anonymous";
                        var role = user?.FindFirstValue(ClaimTypes.Role) ?? "None";

                        context.Response.ContentType = "text/plain";
                        await context.Response.WriteAsync(
                            $"IsAuthenticated:{isAuth}|UserId:{userId}|Username:{username}|Role:{role}");
                    }
                    else
                    {
                        context.Response.StatusCode = 404;
                    }
                });
            });

        return new TestServer(builder);
    }

    /// <summary>
    /// Mock Token 存储，用于测试
    /// </summary>
    private class MockAuthTokenStore : IAuthTokenStore
    {
        public Task<ClaimsPrincipal?> ValidateTokenAsync(string tokenHash)
        {
            // 使用 AuthBridgeMiddleware.ComputeTokenHash 计算原始 token 的哈希
            var validHash = AuthBridgeMiddleware.ComputeTokenHash(ValidToken);
            var expiredHash = AuthBridgeMiddleware.ComputeTokenHash(ExpiredToken);
            var inactiveHash = AuthBridgeMiddleware.ComputeTokenHash(InactiveToken);
            var adminHash = AuthBridgeMiddleware.ComputeTokenHash(AdminToken);

            if (tokenHash == validHash)
            {
                return Task.FromResult<ClaimsPrincipal?>(CreatePrincipal(
                    TestUserId, TestUsername, TestEmail, "UserLogin", "User"));
            }

            if (tokenHash == adminHash)
            {
                return Task.FromResult<ClaimsPrincipal?>(CreatePrincipal(
                    TestUserId, "adminuser", "admin@example.com", "AdminLogin", "Admin"));
            }

            // Expired 和 Inactive token 返回 null（模拟过期/被禁用）
            return Task.FromResult<ClaimsPrincipal?>(null);
        }

        private static ClaimsPrincipal CreatePrincipal(
            int userId, string username, string email, string source, string role)
        {
            var claims = new List<Claim>
            {
                new(ClaimTypes.NameIdentifier, userId.ToString()),
                new(ClaimTypes.Name, username),
                new(ClaimTypes.Email, email),
                new("AuthTokenSource", source),
                new(ClaimTypes.Role, role)
            };
            var identity = new ClaimsIdentity(claims, "AuthBridge");
            return new ClaimsPrincipal(identity);
        }
    }
}
