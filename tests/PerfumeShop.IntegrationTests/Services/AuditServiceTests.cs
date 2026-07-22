using System.Text.Json;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.IntegrationTests.Services;

/// <summary>AuditService 单元测试（正常路径 + 异常路径 + 边界条件）</summary>
public class AuditServiceTests : IDisposable
{
    private readonly string _testLogDir;
    private readonly AuditService _service;

    public AuditServiceTests()
    {
        // 使用临时目录避免污染
        _testLogDir = Path.Combine(Path.GetTempPath(), $"audit_test_{Guid.NewGuid():N}");
        // 使用无效连接字符串 — 文件写入不需要 DB
        _service = new AuditService(
            "Server=invalid;Database=test;Trusted_Connection=True;Connection Timeout=1;",
            _testLogDir);
    }

    public void Dispose()
    {
        // 清理临时目录
        if (Directory.Exists(_testLogDir))
        {
            try { Directory.Delete(_testLogDir, true); } catch { }
        }
    }

    [Fact]
    public async Task LogAsync_WritesJsonlFile()
    {
        // Arrange
        var entry = new AuditEntry
        {
            UserId = 1,
            UserName = "Admin",
            ActionType = AuditAction.Login,
            IpAddress = "127.0.0.1"
        };

        // Act — LogAsync 双写：文件 + DB（DB 会失败但文件应成功）
        await _service.LogAsync(entry);

        // Assert — 验证 JSONL 文件已创建
        var expectedFile = Path.Combine(_testLogDir, $"audit_{DateTime.Now:yyyyMMdd}.jsonl");
        Assert.True(File.Exists(expectedFile), $"JSONL file should exist: {expectedFile}");

        var content = await File.ReadAllTextAsync(expectedFile);
        Assert.Contains("login", content);
        Assert.Contains("Admin", content);

        // 验证 JSON 格式有效
        var json = JsonDocument.Parse(content.Trim());
        Assert.Equal(1, json.RootElement.GetProperty("UserId").GetInt32());
    }

    [Fact]
    public async Task LogAsync_MultipleEntries_AppendToFile()
    {
        // Arrange
        var entries = new[]
        {
            new AuditEntry { UserId = 1, UserName = "Admin", ActionType = AuditAction.Create },
            new AuditEntry { UserId = 2, UserName = "User", ActionType = AuditAction.Update },
            new AuditEntry { UserId = 3, UserName = "System", ActionType = AuditAction.Delete }
        };

        // Act
        foreach (var entry in entries)
            await _service.LogAsync(entry);

        // Assert
        var expectedFile = Path.Combine(_testLogDir, $"audit_{DateTime.Now:yyyyMMdd}.jsonl");
        var lines = await File.ReadAllLinesAsync(expectedFile);
        Assert.Equal(3, lines.Length);
    }

    [Fact]
    public async Task LogAsync_DbFailure_StillWritesFile()
    {
        // DB 连接失败时，文件写入应不受影响（对齐 V18 On Error Resume Next）
        var entry = new AuditEntry
        {
            UserId = 99,
            UserName = "Test",
            ActionType = AuditAction.Export,
            TargetType = AuditTarget.Order,
            TargetId = 123,
            Details = "批量导出订单"
        };

        // Act — DB 写入会失败但不应抛异常
        await _service.LogAsync(entry);

        // Assert
        var expectedFile = Path.Combine(_testLogDir, $"audit_{DateTime.Now:yyyyMMdd}.jsonl");
        Assert.True(File.Exists(expectedFile));
        var content = await File.ReadAllTextAsync(expectedFile);
        Assert.Contains("export", content);
        // JSON 序列化可能转义 Unicode，验证 Details 字段存在
        Assert.Contains("Details", content);
    }

    [Fact]
    public void AuditAction_Constants_MatchV18()
    {
        // 对齐 V18 includes/audit_utils.asp 常量
        Assert.Equal("login", AuditAction.Login);
        Assert.Equal("logout", AuditAction.Logout);
        Assert.Equal("create", AuditAction.Create);
        Assert.Equal("update", AuditAction.Update);
        Assert.Equal("delete", AuditAction.Delete);
        Assert.Equal("export", AuditAction.Export);
        Assert.Equal("batch_operation", AuditAction.BatchOperation);
        Assert.Equal("view_sensitive", AuditAction.ViewSensitive);
        Assert.Equal("privacy_export", AuditAction.PrivacyExport);
        Assert.Equal("privacy_delete", AuditAction.PrivacyDelete);
        Assert.Equal("privacy_consent", AuditAction.PrivacyConsent);
    }

    [Fact]
    public void AuditTarget_Constants_MatchV18()
    {
        // 对齐 V18 includes/audit_utils.asp 目标类型常量
        Assert.Equal("order", AuditTarget.Order);
        Assert.Equal("product", AuditTarget.Product);
        Assert.Equal("user", AuditTarget.User);
        Assert.Equal("coupon", AuditTarget.Coupon);
        Assert.Equal("settings", AuditTarget.Settings);
        Assert.Equal("finance", AuditTarget.Finance);
        Assert.Equal("inventory", AuditTarget.Inventory);
        Assert.Equal("system", AuditTarget.System);
        Assert.Equal("privacy", AuditTarget.Privacy);
    }

    [Fact]
    public void AuditEntry_DefaultValues()
    {
        var entry = new AuditEntry();
        Assert.Equal(0, entry.UserId);
        Assert.Equal("", entry.UserName);
        Assert.Equal("", entry.ActionType);
        Assert.Null(entry.TargetType);
        Assert.Null(entry.TargetId);
        Assert.Null(entry.TargetName);
        Assert.Null(entry.Details);
    }

    [Fact]
    public void AuditQuery_DefaultValues()
    {
        var query = new AuditQuery();
        Assert.Equal(1, query.Page);
        Assert.Equal(20, query.PageSize);
        Assert.Null(query.ActionFilter);
        Assert.Null(query.DateFrom);
        Assert.Null(query.DateTo);
        Assert.Null(query.UserId);
    }

    [Fact]
    public async Task GetLogsAsync_WithInvalidConnection_ReturnsEmptyResult()
    {
        // DB 连接失败时返回空结果（不抛异常）
        var query = new AuditQuery { ActionFilter = AuditAction.Login };
        var result = await _service.GetLogsAsync(query);

        Assert.NotNull(result);
        Assert.Empty(result.Entries);
        Assert.Equal(0, result.TotalCount);
    }

    [Fact]
    public async Task LogAsync_CreatesLogDirectoryIfNotExists()
    {
        // 目录不存在时应自动创建（对齐 V18: EnsureUploadDir / CreateFolderRecursive）
        var subDir = Path.Combine(_testLogDir, "nested", "logs");
        var nestedService = new AuditService(
            "Server=invalid;Connection Timeout=1;",
            subDir);

        await nestedService.LogAsync(new AuditEntry
        {
            UserId = 1,
            ActionType = AuditAction.Login
        });

        Assert.True(Directory.Exists(subDir));
    }
}
