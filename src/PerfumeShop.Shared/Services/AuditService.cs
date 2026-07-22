using System.Data.SqlClient;
using System.Text.Encodings.Web;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;

namespace PerfumeShop.Shared.Services;

// ============================================
// V19 AuditService — 对齐 V18 includes/audit_utils.asp + includes/logger.asp
// 功能：审计日志双写（文件 + AuditLogs 数据库表）、结构化日志（JSON 格式）、
//       按用户/时间/操作类型查询
// ============================================

/// <summary>审计操作类型（对齐 V18 AUDIT_ACTION_* 常量）</summary>
public static class AuditAction
{
    public const string Login = "login";                       // 对齐 AUDIT_ACTION_LOGIN
    public const string Logout = "logout";                     // 对齐 AUDIT_ACTION_LOGOUT
    public const string Create = "create";                     // 对齐 AUDIT_ACTION_CREATE
    public const string Update = "update";                     // 对齐 AUDIT_ACTION_UPDATE
    public const string Delete = "delete";                     // 对齐 AUDIT_ACTION_DELETE
    public const string Export = "export";                     // 对齐 AUDIT_ACTION_EXPORT
    public const string BatchOperation = "batch_operation";    // 对齐 AUDIT_ACTION_BATCH
    public const string ViewSensitive = "view_sensitive";      // 对齐 AUDIT_ACTION_VIEW
    public const string PrivacyExport = "privacy_export";      // 对齐 AUDIT_ACTION_PRIVACY_EXPORT
    public const string PrivacyDelete = "privacy_delete";      // 对齐 AUDIT_ACTION_PRIVACY_DELETE
    public const string PrivacyConsent = "privacy_consent";    // 对齐 AUDIT_ACTION_PRIVACY_CONSENT
}

/// <summary>审计目标类型（对齐 V18 AUDIT_TARGET_* 常量）</summary>
public static class AuditTarget
{
    public const string Order = "order";           // 对齐 AUDIT_TARGET_ORDER
    public const string Product = "product";       // 对齐 AUDIT_TARGET_PRODUCT
    public const string User = "user";             // 对齐 AUDIT_TARGET_USER
    public const string Coupon = "coupon";         // 对齐 AUDIT_TARGET_COUPON
    public const string Settings = "settings";     // 对齐 AUDIT_TARGET_SETTINGS
    public const string Finance = "finance";       // 对齐 AUDIT_TARGET_FINANCE
    public const string Inventory = "inventory";   // 对齐 AUDIT_TARGET_INVENTORY
    public const string System = "system";         // 对齐 AUDIT_TARGET_SYSTEM
    public const string Privacy = "privacy";       // 对齐 AUDIT_TARGET_PRIVACY
}

/// <summary>审计日志条目 — 对齐 V18 AuditLog 子过程参数</summary>
public class AuditEntry
{
    public int UserId { get; set; }
    public string UserName { get; set; } = "";
    public string ActionType { get; set; } = "";       // 对齐 actionType
    public string? TargetType { get; set; }             // 对齐 targetType
    public int? TargetId { get; set; }                  // 对齐 targetID
    public string? TargetName { get; set; }             // 对齐 targetName
    public string? Details { get; set; }                // 对齐 details
    public string IpAddress { get; set; } = "";
    public string UserAgent { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>审计日志查询条件 — 对齐 V18 GetAuditLogs(page, pageSize, actionFilter, dateFrom, dateTo)</summary>
public class AuditQuery
{
    public int Page { get; set; } = 1;
    public int PageSize { get; set; } = 20;
    public string? ActionFilter { get; set; }       // 对齐 actionFilter
    public DateTime? DateFrom { get; set; }          // 对齐 dateFrom
    public DateTime? DateTo { get; set; }            // 对齐 dateTo
    public int? UserId { get; set; }
}

/// <summary>审计日志查询结果</summary>
public class AuditLogResult
{
    public List<AuditEntry> Entries { get; set; } = new();
    public int TotalCount { get; set; }
}

/// <summary>审计服务接口</summary>
public interface IAuditService
{
    /// <summary>写入审计日志 — 对齐 V18: AuditLog 子过程</summary>
    Task LogAsync(AuditEntry entry);

    /// <summary>查询审计日志（分页）— 对齐 V18: GetAuditLogs</summary>
    Task<AuditLogResult> GetLogsAsync(AuditQuery query);

    /// <summary>按用户查询操作记录 — 对齐 V18: GetAuditLogs 按 AdminID 过滤</summary>
    Task<List<AuditEntry>> GetUserActionsAsync(int userId, int top = 100);
}

/// <summary>
/// 审计服务实现 — 双写文件 + 数据库
/// 对齐 V18 函数映射：
///   - AuditLog(actionType, targetType, targetID, targetName, details) → LogAsync
///   - AuditOrder       → LogAsync(entry.TargetType = "order")
///   - AuditProduct     → LogAsync(entry.TargetType = "product")
///   - AuditUser        → LogAsync(entry.TargetType = "user")
///   - AuditBatch       → LogAsync(entry.ActionType = "batch_operation")
///   - LogPrivacyAction → LogAsync(entry.ActionType = "privacy_*")
///   - LogCookieConsent → LogAsync(entry.ActionType = "privacy_consent")
///   - GetAuditLogs     → GetLogsAsync
///   - GetAuditLogCount → GetLogsAsync.TotalCount
///   - EnsureAuditLogTable → EnsureTableAsync
///   - LOG_WriteToFile  → 文件双写（JSONL 格式）
///   - LOG_WriteToDB    → 数据库写入
/// </summary>
public class AuditService : IAuditService
{
    private readonly string _connectionString;
    private readonly string _logDirectory;

    public AuditService(string connectionString, string logDirectory = "logs")
    {
        _connectionString = connectionString;
        _logDirectory = logDirectory;
    }

    /// <summary>确保 AuditLogs 表存在 — 对齐 V18: EnsureAuditLogTable</summary>
    public async Task EnsureTableAsync()
    {
        const string sql = @"
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='AuditLogs')
            BEGIN
                CREATE TABLE [AuditLogs] (
                    [LogID] INT IDENTITY(1,1) NOT NULL,
                    [UserID] INT NOT NULL DEFAULT 0,
                    [UserName] NVARCHAR(100) NULL,
                    [ActionType] NVARCHAR(50) NOT NULL,
                    [TargetType] NVARCHAR(50) NULL,
                    [TargetID] INT NULL,
                    [TargetName] NVARCHAR(200) NULL,
                    [Details] NVARCHAR(MAX) NULL,
                    [IPAddress] NVARCHAR(50) NULL,
                    [UserAgent] NVARCHAR(500) NULL,
                    [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(),
                    CONSTRAINT [PK_AuditLogs] PRIMARY KEY CLUSTERED ([LogID] ASC)
                );
                CREATE NONCLUSTERED INDEX [IX_AuditLogs_User] ON [AuditLogs]([UserID], [CreatedAt] DESC);
                CREATE NONCLUSTERED INDEX [IX_AuditLogs_Action] ON [AuditLogs]([ActionType], [CreatedAt] DESC);
            END";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();
        await using var cmd = new SqlCommand(sql, conn);
        cmd.CommandTimeout = 30;
        await cmd.ExecuteNonQueryAsync();
    }

    // 对齐 V18: AuditLog + LOG_WriteToFile（双写）
    public async Task LogAsync(AuditEntry entry)
    {
        entry.CreatedAt = DateTime.UtcNow;

        // 1. 写入文件（JSONL 按天滚动）— 对齐 V18: LOG_WriteToFile
        WriteToFile(entry);

        // 2. 写入数据库 — 对齐 V18: AuditLog INSERT INTO AdminAuditLog
        await WriteToDbAsync(entry);
    }

    // 对齐 V18: LOG_WriteToFile 按日轮转
    private void WriteToFile(AuditEntry entry)
    {
        try
        {
            if (!Directory.Exists(_logDirectory))
                Directory.CreateDirectory(_logDirectory);

            // 对齐 V18: 文件名 audit_YYYYMMDD.jsonl（按天滚动）
            var fileName = $"audit_{DateTime.Now:yyyyMMdd}.jsonl";
            var filePath = Path.Combine(_logDirectory, fileName);

            var json = JsonSerializer.Serialize(new
            {
                entry.UserId,
                entry.UserName,
                entry.ActionType,
                entry.TargetType,
                entry.TargetId,
                entry.TargetName,
                entry.Details,
                entry.IpAddress,
                entry.UserAgent,
                entry.CreatedAt
            }, new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping });

            File.AppendAllText(filePath, json + Environment.NewLine);
        }
        catch
        {
            // 文件写入失败不应影响主流程 — 对齐 V18 On Error Resume Next
        }
    }

    // 对齐 V18: INSERT INTO AdminAuditLog（参数化查询）
    private async Task WriteToDbAsync(AuditEntry entry)
    {
        const string sql = @"
            INSERT INTO AuditLogs (UserID, UserName, ActionType, TargetType, TargetID, TargetName, Details, IPAddress, UserAgent, CreatedAt)
            VALUES (@UserID, @UserName, @ActionType, @TargetType, @TargetID, @TargetName, @Details, @IPAddress, @UserAgent, @CreatedAt)";

        try
        {
            await using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();
            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@UserID", entry.UserId);
            cmd.Parameters.AddWithValue("@UserName", (object?)entry.UserName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ActionType", entry.ActionType);
            cmd.Parameters.AddWithValue("@TargetType", (object?)entry.TargetType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TargetID", (object?)entry.TargetId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TargetName", (object?)entry.TargetName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Details", (object?)entry.Details ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IPAddress", (object?)entry.IpAddress ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@UserAgent", (object?)entry.UserAgent ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CreatedAt", entry.CreatedAt);
            await cmd.ExecuteNonQueryAsync();
        }
        catch
        {
            // 数据库写入失败不应影响主流程 — 对齐 V18 On Error Resume Next
        }
    }

    // 对齐 V18: GetAuditLogs(page, pageSize, actionFilter, dateFrom, dateTo)
    public async Task<AuditLogResult> GetLogsAsync(AuditQuery query)
    {
        var result = new AuditLogResult();
        var whereClauses = new List<string> { "1=1" };

        if (!string.IsNullOrEmpty(query.ActionFilter))
            whereClauses.Add("ActionType = @ActionFilter");
        if (query.DateFrom.HasValue)
            whereClauses.Add("CreatedAt >= @DateFrom");
        if (query.DateTo.HasValue)
            whereClauses.Add("CreatedAt <= @DateTo");
        if (query.UserId.HasValue)
            whereClauses.Add("UserID = @UserID");

        var where = string.Join(" AND ", whereClauses);
        var offset = (query.Page - 1) * query.PageSize;

        // 查询总数 — 对齐 V18: GetAuditLogCount
        var countSql = $"SELECT COUNT(*) FROM AuditLogs WHERE {where}";
        // 查询数据 — 对齐 V18: GetAuditLogs OFFSET FETCH
        var dataSql = $@"
            SELECT UserID, UserName, ActionType, TargetType, TargetID, TargetName, Details, IPAddress, CreatedAt
            FROM AuditLogs WHERE {where}
            ORDER BY CreatedAt DESC
            OFFSET {offset} ROWS FETCH NEXT {query.PageSize} ROWS ONLY";

        try
        {
            await using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            // Count
            await using (var countCmd = new SqlCommand(countSql, conn))
            {
                AddQueryParams(countCmd, query);
                result.TotalCount = (int)(await countCmd.ExecuteScalarAsync() ?? 0);
            }

            // Data
            await using (var dataCmd = new SqlCommand(dataSql, conn))
            {
                AddQueryParams(dataCmd, query);
                await using var reader = await dataCmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Entries.Add(new AuditEntry
                    {
                        UserId = reader.GetInt32(0),
                        UserName = reader.IsDBNull(1) ? "" : reader.GetString(1),
                        ActionType = reader.GetString(2),
                        TargetType = reader.IsDBNull(3) ? null : reader.GetString(3),
                        TargetId = reader.IsDBNull(4) ? null : reader.GetInt32(4),
                        TargetName = reader.IsDBNull(5) ? null : reader.GetString(5),
                        Details = reader.IsDBNull(6) ? null : reader.GetString(6),
                        IpAddress = reader.IsDBNull(7) ? "" : reader.GetString(7),
                        CreatedAt = reader.IsDBNull(8) ? DateTime.UtcNow : reader.GetDateTime(8)
                    });
                }
            }
        }
        catch
        {
            // 查询失败返回空结果
        }

        return result;
    }

    // 对齐 V18: GetAuditLogs 按 AdminID 过滤
    public async Task<List<AuditEntry>> GetUserActionsAsync(int userId, int top = 100)
    {
        var query = new AuditQuery { UserId = userId, PageSize = top };
        var result = await GetLogsAsync(query);
        return result.Entries;
    }

    private static void AddQueryParams(SqlCommand cmd, AuditQuery query)
    {
        if (!string.IsNullOrEmpty(query.ActionFilter))
            cmd.Parameters.AddWithValue("@ActionFilter", query.ActionFilter);
        if (query.DateFrom.HasValue)
            cmd.Parameters.AddWithValue("@DateFrom", query.DateFrom.Value);
        if (query.DateTo.HasValue)
            cmd.Parameters.AddWithValue("@DateTo", query.DateTo.Value);
        if (query.UserId.HasValue)
            cmd.Parameters.AddWithValue("@UserID", query.UserId.Value);
    }
}

/// <summary>DI 注册扩展</summary>
public static class AuditServiceExtensions
{
    public static IServiceCollection AddAuditService(this IServiceCollection services, string connectionString, string logDirectory = "logs")
    {
        services.AddSingleton<IAuditService>(new AuditService(connectionString, logDirectory));
        return services;
    }
}
