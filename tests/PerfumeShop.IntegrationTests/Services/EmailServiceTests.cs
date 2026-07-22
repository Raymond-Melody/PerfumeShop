using Moq;
using MimeKit;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.IntegrationTests.Services;

/// <summary>EmailService 单元测试（正常路径 + 异常路径 + 边界条件，Mock SMTP）</summary>
public class EmailServiceTests : IDisposable
{
    private readonly Mock<ISmtpClient> _mockSmtp;
    private readonly SmtpOptions _options;
    private readonly EmailService _service;
    private readonly string _templateDir;

    public EmailServiceTests()
    {
        _mockSmtp = new Mock<ISmtpClient>();
        _options = new SmtpOptions
        {
            Host = "smtp.test.com",
            Port = 587,
            UseSsl = true,
            Username = "test@test.com",
            Password = "password",
            FromAddress = "noreply@test.com",
            FromName = "TestShop"
        };

        // 创建临时模板目录
        _templateDir = Path.Combine(Path.GetTempPath(), $"email_tpl_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_templateDir);

        _service = new EmailService(_mockSmtp.Object, _options, _templateDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_templateDir))
        {
            try { Directory.Delete(_templateDir, true); } catch { }
        }
    }

    [Fact]
    public async Task SendAsync_ValidMessage_CallsSmtpClient()
    {
        // Arrange
        _mockSmtp.Setup(s => s.ConnectAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.AuthenticateAsync(It.IsAny<string>(), It.IsAny<string>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.SendAsync(It.IsAny<MimeMessage>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.DisconnectAsync(It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);

        var msg = new EmailMessage
        {
            To = "user@test.com",
            Subject = "测试邮件",
            Body = "<h1>Hello</h1>",
            IsHtml = true
        };

        // Act
        var result = await _service.SendAsync(msg);

        // Assert
        Assert.True(result.Success);
        _mockSmtp.Verify(s => s.ConnectAsync("smtp.test.com", 587, true, default), Times.Once);
        _mockSmtp.Verify(s => s.AuthenticateAsync("test@test.com", "password", default), Times.Once);
        _mockSmtp.Verify(s => s.SendAsync(It.IsAny<MimeMessage>(), default), Times.Once);
        _mockSmtp.Verify(s => s.DisconnectAsync(true, default), Times.Once);
    }

    [Fact]
    public async Task SendAsync_SmtpFailure_ReturnsErrorResult()
    {
        // Arrange
        _mockSmtp.Setup(s => s.ConnectAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<bool>(), default))
                 .ThrowsAsync(new InvalidOperationException("SMTP 连接失败"));

        var msg = new EmailMessage
        {
            To = "user@test.com",
            Subject = "测试",
            Body = "body"
        };

        // Act
        var result = await _service.SendAsync(msg);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("SMTP 连接失败", result.Error);
    }

    [Fact]
    public async Task SendTemplateAsync_WithTemplate_RendersCorrectly()
    {
        // Arrange — 创建测试模板
        var templateContent = "<html><body><h1>欢迎 {{FULL_NAME}}</h1><p>{{SITE_NAME}}</p></body></html>";
        await File.WriteAllTextAsync(Path.Combine(_templateDir, "welcome.html"), templateContent);

        _mockSmtp.Setup(s => s.ConnectAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.AuthenticateAsync(It.IsAny<string>(), It.IsAny<string>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.SendAsync(It.IsAny<MimeMessage>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.DisconnectAsync(It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);

        var model = new EmailTemplateModel
        {
            Variables = new Dictionary<string, string>
            {
                ["FULL_NAME"] = "张三",
                ["SITE_NAME"] = "香氛定制",
                ["SUBJECT"] = "欢迎邮件"
            }
        };

        // Act
        var result = await _service.SendTemplateAsync("welcome", model, "user@test.com");

        // Assert
        Assert.True(result.Success);
        _mockSmtp.Verify(s => s.SendAsync(It.Is<MimeMessage>(m =>
            m.Subject == "欢迎邮件"), default), Times.Once);
    }

    [Fact]
    public async Task SendTemplateAsync_MissingTemplate_UsesFallback()
    {
        // 模板不存在时使用默认 HTML
        _mockSmtp.Setup(s => s.ConnectAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.AuthenticateAsync(It.IsAny<string>(), It.IsAny<string>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.SendAsync(It.IsAny<MimeMessage>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.DisconnectAsync(It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);

        var model = new EmailTemplateModel
        {
            Variables = new Dictionary<string, string> { ["SUBJECT"] = "通知" }
        };

        // Act
        var result = await _service.SendTemplateAsync("nonexistent", model, "user@test.com");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public async Task QueueAsync_WithoutConnectionString_ReturnsError()
    {
        // 没有连接字符串时队列化发送应返回错误
        var msg = new EmailMessage
        {
            To = "user@test.com",
            Subject = "测试",
            Body = "body"
        };

        var result = await _service.QueueAsync(msg);

        Assert.False(result.Success);
        Assert.Contains("数据库连接字符串", result.Error);
    }

    [Fact]
    public async Task QueueAsync_WithConnectionString_TriesDbInsert()
    {
        // 有连接字符串但连接失败
        var serviceWithDb = new EmailService(
            _mockSmtp.Object, _options, _templateDir,
            "Server=invalid;Connection Timeout=1;");

        var msg = new EmailMessage
        {
            To = "user@test.com",
            Subject = "队列测试",
            Body = "body"
        };

        var result = await serviceWithDb.QueueAsync(msg);
        // DB 连接失败应返回错误
        Assert.False(result.Success);
    }

    [Fact]
    public void SmtpOptions_DefaultValues()
    {
        var opts = new SmtpOptions();
        Assert.Equal("localhost", opts.Host);
        Assert.Equal(25, opts.Port);
        Assert.False(opts.UseSsl);
        Assert.Null(opts.Username);
        Assert.Null(opts.Password);
        Assert.Equal("noreply@yourperfume.com", opts.FromAddress);
        Assert.Equal("香氛定制", opts.FromName);
    }

    [Fact]
    public void EmailMessage_DefaultValues()
    {
        var msg = new EmailMessage();
        Assert.Equal("", msg.To);
        Assert.Equal("", msg.Subject);
        Assert.Equal("", msg.Body);
        Assert.True(msg.IsHtml);
        Assert.Null(msg.FromName);
        Assert.Null(msg.FromAddress);
    }

    [Fact]
    public async Task SendAsync_NoAuth_SkipsAuthentication()
    {
        // 无认证信息时应跳过 AuthenticateAsync
        var noAuthOptions = new SmtpOptions
        {
            Host = "smtp.test.com",
            Port = 25,
            Username = null,
            Password = null
        };
        var service = new EmailService(_mockSmtp.Object, noAuthOptions, _templateDir);

        _mockSmtp.Setup(s => s.ConnectAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.SendAsync(It.IsAny<MimeMessage>(), default))
                 .Returns(Task.CompletedTask);
        _mockSmtp.Setup(s => s.DisconnectAsync(It.IsAny<bool>(), default))
                 .Returns(Task.CompletedTask);

        var msg = new EmailMessage { To = "user@test.com", Subject = "Test", Body = "Hi" };
        var result = await service.SendAsync(msg);

        Assert.True(result.Success);
        _mockSmtp.Verify(s => s.AuthenticateAsync(It.IsAny<string>(), It.IsAny<string>(), default), Times.Never);
    }
}
