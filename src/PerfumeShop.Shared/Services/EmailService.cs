using System.Data.SqlClient;
using MailKit.Net.Smtp;
using Microsoft.Extensions.DependencyInjection;
using MimeKit;

namespace PerfumeShop.Shared.Services;

// ============================================
// V19 EmailService — 对齐 V18 includes/email_service.asp + includes/email_utils.asp
// 功能：SMTP 发送（MailKit）、模板渲染（字符串替换）、可选队列化发送
// ============================================

/// <summary>邮件消息模型 — 对齐 V18 SendEmail(toEmail, subject, body, isHtml)</summary>
public class EmailMessage
{
    public string To { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Body { get; set; } = "";
    public bool IsHtml { get; set; } = true;
    public string? FromName { get; set; }
    public string? FromAddress { get; set; }
}

/// <summary>邮件模板模型 — 对齐 V18 ES_LoadTemplate(templateName, replacements)</summary>
public class EmailTemplateModel
{
    /// <summary>模板变量键值对 — 对齐 V18 replacements Dictionary</summary>
    public Dictionary<string, string> Variables { get; set; } = new();
}

/// <summary>邮件发送结果</summary>
public class EmailResult
{
    public bool Success { get; set; }
    public string? Error { get; set; }
    public int? QueueId { get; set; }
}

/// <summary>SMTP 发送抽象接口（用于测试 Mock）</summary>
public interface ISmtpClient : IDisposable
{
    Task ConnectAsync(string host, int port, bool useSsl, CancellationToken cancellationToken = default);
    Task AuthenticateAsync(string user, string password, CancellationToken cancellationToken = default);
    Task SendAsync(MimeMessage message, CancellationToken cancellationToken = default);
    Task DisconnectAsync(bool quit, CancellationToken cancellationToken = default);
}

/// <summary>默认 MailKit SMTP 实现</summary>
public class MailKitSmtpClient : ISmtpClient
{
    private readonly SmtpClient _client = new();

    public Task ConnectAsync(string host, int port, bool useSsl, CancellationToken cancellationToken = default)
        => _client.ConnectAsync(host, port, useSsl, cancellationToken);

    public Task AuthenticateAsync(string user, string password, CancellationToken cancellationToken = default)
        => _client.AuthenticateAsync(user, password, cancellationToken);

    public Task SendAsync(MimeMessage message, CancellationToken cancellationToken = default)
        => _client.SendAsync(message, cancellationToken);

    public Task DisconnectAsync(bool quit, CancellationToken cancellationToken = default)
        => _client.DisconnectAsync(quit, cancellationToken);

    public void Dispose() => _client.Dispose();
}

/// <summary>SMTP 配置</summary>
public class SmtpOptions
{
    public string Host { get; set; } = "localhost";
    public int Port { get; set; } = 25;
    public bool UseSsl { get; set; } = false;
    public string? Username { get; set; }
    public string? Password { get; set; }
    public string FromAddress { get; set; } = "noreply@yourperfume.com";
    public string FromName { get; set; } = "香氛定制";
}

/// <summary>邮件服务接口</summary>
public interface IEmailService
{
    /// <summary>发送邮件 — 对齐 V18: SendEmail(toEmail, subject, body, isHtml)</summary>
    Task<EmailResult> SendAsync(EmailMessage msg);

    /// <summary>发送模板邮件 — 对齐 V18: ES_SendWelcomeEmail, ES_SendOrderConfirmation, ES_SendPasswordReset</summary>
    Task<EmailResult> SendTemplateAsync(string templateName, EmailTemplateModel model, string to);

    /// <summary>队列化发送 — 对齐 V18: EmailQueue 表后台处理</summary>
    Task<EmailResult> QueueAsync(EmailMessage msg);
}

/// <summary>
/// 邮件服务实现
/// 对齐 V18 函数映射：
///   - SendEmail(toEmail, subject, body, isHtml)          → SendAsync
///   - ES_LoadTemplate(templateName, replacements)        → LoadTemplate (内部)
///   - ES_SendWelcomeEmail(userId, toEmail, fullName)     → SendTemplateAsync("welcome", ...)
///   - ES_SendOrderConfirmation(userId, orderId)          → SendTemplateAsync("order-confirmation", ...)
///   - ES_SendShippingNotification(userId, orderId, ...)  → SendTemplateAsync("shipping-notification", ...)
///   - ES_SendPasswordReset(toEmail, fullName, token, ...)→ SendTemplateAsync("password-reset", ...)
///   - ES_SendRefundNotification(userId, orderId, amount) → SendAsync（直接构造）
///   - SendPasswordResetEmail(toEmail, fullName, token)   → SendTemplateAsync("password-reset", ...)
///   - SendUserPasswordResetEmail(toEmail, fullName, ...) → SendTemplateAsync("password-reset", ...)
/// </summary>
public class EmailService : IEmailService
{
    private readonly ISmtpClient _smtpClient;
    private readonly SmtpOptions _options;
    private readonly string _templateDirectory;
    private readonly string? _connectionString;

    /// <param name="smtpClient">SMTP 客户端（可注入用于 Mock 测试）</param>
    /// <param name="options">SMTP 配置</param>
    /// <param name="templateDirectory">模板目录路径</param>
    /// <param name="connectionString">数据库连接字符串（队列化发送时需要）</param>
    public EmailService(ISmtpClient smtpClient, SmtpOptions options,
        string templateDirectory = "EmailTemplates", string? connectionString = null)
    {
        _smtpClient = smtpClient;
        _options = options;
        _templateDirectory = templateDirectory;
        _connectionString = connectionString;
    }

    // 对齐 V18: SendEmail(toEmail, subject, body, isHtml)
    public async Task<EmailResult> SendAsync(EmailMessage msg)
    {
        try
        {
            var mime = new MimeMessage();
            mime.From.Add(new MailboxAddress(
                msg.FromName ?? _options.FromName,
                msg.FromAddress ?? _options.FromAddress));
            mime.To.Add(MailboxAddress.Parse(msg.To));
            mime.Subject = msg.Subject;

            var builder = new BodyBuilder();
            if (msg.IsHtml)
                builder.HtmlBody = msg.Body;
            else
                builder.TextBody = msg.Body;
            mime.Body = builder.ToMessageBody();

            await _smtpClient.ConnectAsync(_options.Host, _options.Port, _options.UseSsl);

            if (!string.IsNullOrEmpty(_options.Username))
                await _smtpClient.AuthenticateAsync(_options.Username, _options.Password ?? "");

            await _smtpClient.SendAsync(mime);
            await _smtpClient.DisconnectAsync(true);

            return new EmailResult { Success = true };
        }
        catch (Exception ex)
        {
            return new EmailResult { Success = false, Error = ex.Message };
        }
    }

    // 对齐 V18: ES_LoadTemplate(templateName, replacements)
    internal string LoadTemplate(string templateName, Dictionary<string, string> variables)
    {
        // 尝试多个路径查找模板
        string[] searchPaths =
        {
            Path.Combine(_templateDirectory, $"{templateName}.html"),
            Path.Combine(AppContext.BaseDirectory, _templateDirectory, $"{templateName}.html"),
            Path.Combine(Directory.GetCurrentDirectory(), _templateDirectory, $"{templateName}.html"),
        };

        string? templateContent = null;
        foreach (var path in searchPaths)
        {
            if (File.Exists(path))
            {
                templateContent = File.ReadAllText(path);
                break;
            }
        }

        if (templateContent == null)
            return "";

        // 对齐 V18: Replace(template, "{{" & key & "}}", replacements.Item(key))
        foreach (var kv in variables)
        {
            templateContent = templateContent.Replace($"{{{{{kv.Key}}}}}", kv.Value);
        }

        return templateContent;
    }

    // 对齐 V18: ES_SendWelcomeEmail, ES_SendOrderConfirmation, ES_SendPasswordReset 等
    public async Task<EmailResult> SendTemplateAsync(string templateName, EmailTemplateModel model, string to)
    {
        var body = LoadTemplate(templateName, model.Variables);
        if (string.IsNullOrEmpty(body))
        {
            // 对齐 V18: 模板不存在时使用默认 HTML
            body = $"<html><body><p>模板 {templateName} 未找到。</p></body></html>";
        }

        var subject = model.Variables.TryGetValue("SUBJECT", out var subj) ? subj : $"通知 - {_options.FromName}";

        return await SendAsync(new EmailMessage
        {
            To = to,
            Subject = subject,
            Body = body,
            IsHtml = true
        });
    }

    // 对齐 V18: EmailQueue 表后台处理
    public async Task<EmailResult> QueueAsync(EmailMessage msg)
    {
        if (string.IsNullOrEmpty(_connectionString))
        {
            return new EmailResult { Success = false, Error = "队列化发送需要数据库连接字符串" };
        }

        try
        {
            await EnsureEmailQueueTableAsync();

            const string sql = @"
                INSERT INTO EmailQueue (ToAddress, Subject, Body, IsHtml, Status, CreatedAt)
                VALUES (@To, @Subject, @Body, @IsHtml, 'Pending', GETDATE());
                SELECT SCOPE_IDENTITY();";

            await using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@To", msg.To);
            cmd.Parameters.AddWithValue("@Subject", msg.Subject);
            cmd.Parameters.AddWithValue("@Body", msg.Body);
            cmd.Parameters.AddWithValue("@IsHtml", msg.IsHtml);

            var id = await cmd.ExecuteScalarAsync();
            return new EmailResult { Success = true, QueueId = Convert.ToInt32(id) };
        }
        catch (Exception ex)
        {
            return new EmailResult { Success = false, Error = ex.Message };
        }
    }

    private async Task EnsureEmailQueueTableAsync()
    {
        const string sql = @"
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='EmailQueue')
            BEGIN
                CREATE TABLE [EmailQueue] (
                    [QueueID] INT IDENTITY(1,1) NOT NULL,
                    [ToAddress] NVARCHAR(200) NOT NULL,
                    [Subject] NVARCHAR(500) NULL,
                    [Body] NVARCHAR(MAX) NULL,
                    [IsHtml] BIT NOT NULL DEFAULT 1,
                    [Status] NVARCHAR(20) NOT NULL DEFAULT 'Pending',
                    [Attempts] INT NOT NULL DEFAULT 0,
                    [LastError] NVARCHAR(MAX) NULL,
                    [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(),
                    [SentAt] DATETIME2(7) NULL,
                    CONSTRAINT [PK_EmailQueue] PRIMARY KEY CLUSTERED ([QueueID] ASC)
                );
                CREATE NONCLUSTERED INDEX [IX_EmailQueue_Status] ON [EmailQueue]([Status], [CreatedAt]);
            END";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();
        await using var cmd = new SqlCommand(sql, conn);
        cmd.CommandTimeout = 30;
        await cmd.ExecuteNonQueryAsync();
    }
}

/// <summary>DI 注册扩展</summary>
public static class EmailServiceExtensions
{
    public static IServiceCollection AddEmailService(this IServiceCollection services,
        SmtpOptions options, string templateDirectory = "EmailTemplates", string? connectionString = null)
    {
        services.AddSingleton<ISmtpClient, MailKitSmtpClient>();
        services.AddSingleton<IEmailService>(sp =>
            new EmailService(sp.GetRequiredService<ISmtpClient>(), options, templateDirectory, connectionString));
        return services;
    }
}
