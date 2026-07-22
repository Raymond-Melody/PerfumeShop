using System.Data;
using System.Data.SqlClient;
using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;

namespace PerfumeShop.Shared.Services;

// ============================================
// V19 TrackingService — 对齐 V18 includes/tracking_utils.asp + api/track.asp
// 功能：事件写库（TrackingEvents 表）、像素追踪响应（1x1 GIF）、按用户查询
// ============================================

/// <summary>追踪事件类型（对齐 V18 TU_LogBehavior actionType 参数）</summary>
public static class TrackingEventType
{
    public const string ProductView = "PRODUCT_VIEW";   // 对齐 TU_LogProductView
    public const string Search = "SEARCH";               // 对齐 TU_LogSearch
    public const string CartAdd = "CART_ADD";            // 对齐 TU_LogCartAdd
    public const string Favorite = "FAVORITE";           // 对齐 TU_LogFavorite
    public const string PageView = "PAGE_VIEW";
    public const string Custom = "CUSTOM";
}

/// <summary>
/// 追踪事件模型 — 对齐 V18 TU_LogBehavior 参数 (userId, actionType, targetId, targetType, extraData)
/// </summary>
public class TrackingEvent
{
    public int Id { get; set; }
    public int UserId { get; set; }
    public string EventType { get; set; } = "";
    public string? EventData { get; set; }
    public string IpAddress { get; set; } = "";
    public string UserAgent { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>追踪服务接口</summary>
public interface ITrackingService
{
    /// <summary>记录追踪事件 — 对齐 V18: TU_LogBehavior</summary>
    Task TrackEventAsync(TrackingEvent evt);

    /// <summary>生成 1x1 透明 GIF — 对齐 V18: TU_RenderTrackingScript 中的像素追踪</summary>
    byte[] GenerateTrackingPixel();

    /// <summary>按用户查询事件 — 对齐 V18: TU_GetDailyStats 管理端查询</summary>
    Task<List<TrackingEvent>> GetEventsByUserAsync(int userId, int top = 100);
}

/// <summary>
/// 追踪服务实现
/// 对齐 V18 函数映射：
///   - TU_LogBehavior          → TrackEventAsync
///   - TU_LogProductView       → TrackEventAsync(evt.EventType = PRODUCT_VIEW)
///   - TU_LogSearch            → TrackEventAsync(evt.EventType = SEARCH)
///   - TU_LogCartAdd           → TrackEventAsync(evt.EventType = CART_ADD)
///   - TU_LogFavorite          → TrackEventAsync(evt.EventType = FAVORITE)
///   - TU_GetDailyStats        → GetEventsByUserAsync（管理端查询）
///   - TU_RenderTrackingScript → GenerateTrackingPixel（1x1 GIF 响应）
///   - TU_GetHotSearches       → 可通过 GetEventsByUserAsync + EventType=SEARCH 聚合
/// </summary>
public class TrackingService : ITrackingService
{
    private readonly string _connectionString;

    // 1x1 透明 GIF（最小有效 GIF 文件，43 字节）
    private static readonly byte[] TransparentGif =
    {
        0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00,
        0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x21, 0xF9, 0x04, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x2C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02,
        0x44, 0x01, 0x00, 0x3B
    };

    public TrackingService(string connectionString)
    {
        _connectionString = connectionString;
    }

    /// <summary>确保 TrackingEvents 表存在（参考 v15_app_logs.sql 的 AppLogs 表结构）</summary>
    public async Task EnsureTableAsync()
    {
        const string sql = @"
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='TrackingEvents')
            BEGIN
                CREATE TABLE [TrackingEvents] (
                    [Id] INT IDENTITY(1,1) NOT NULL,
                    [UserId] INT NOT NULL DEFAULT 0,
                    [EventType] NVARCHAR(50) NOT NULL,
                    [EventData] NVARCHAR(MAX) NULL,
                    [IpAddress] NVARCHAR(50) NULL,
                    [UserAgent] NVARCHAR(500) NULL,
                    [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(),
                    CONSTRAINT [PK_TrackingEvents] PRIMARY KEY CLUSTERED ([Id] ASC)
                );
                CREATE NONCLUSTERED INDEX [IX_TrackingEvents_User] ON [TrackingEvents]([UserId], [CreatedAt] DESC);
                CREATE NONCLUSTERED INDEX [IX_TrackingEvents_Type] ON [TrackingEvents]([EventType], [CreatedAt] DESC);
            END";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();
        await using var cmd = new SqlCommand(sql, conn);
        cmd.CommandTimeout = 30;
        await cmd.ExecuteNonQueryAsync();
    }

    // 对齐 V18: TU_LogBehavior(userId, actionType, targetId, targetType, extraData)
    public async Task TrackEventAsync(TrackingEvent evt)
    {
        const string sql = @"
            INSERT INTO TrackingEvents (UserId, EventType, EventData, IpAddress, UserAgent, CreatedAt)
            VALUES (@UserId, @EventType, @EventData, @IpAddress, @UserAgent, @CreatedAt)";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();
        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@UserId", evt.UserId);
        cmd.Parameters.AddWithValue("@EventType", evt.EventType ?? TrackingEventType.Custom);
        cmd.Parameters.AddWithValue("@EventData", (object?)evt.EventData ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@IpAddress", (object?)evt.IpAddress ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@UserAgent", (object?)evt.UserAgent ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@CreatedAt", evt.CreatedAt);
        await cmd.ExecuteNonQueryAsync();
    }

    // 对齐 V18: TU_RenderTrackingScript 中的像素追踪
    public byte[] GenerateTrackingPixel() => TransparentGif;

    // 对齐 V18: TU_GetDailyStats 查询功能
    public async Task<List<TrackingEvent>> GetEventsByUserAsync(int userId, int top = 100)
    {
        var results = new List<TrackingEvent>();
        const string sql = @"
            SELECT TOP(@Top) Id, UserId, EventType, EventData, IpAddress, UserAgent, CreatedAt
            FROM TrackingEvents
            WHERE UserId = @UserId
            ORDER BY CreatedAt DESC";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync();
        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@Top", top);
        cmd.Parameters.AddWithValue("@UserId", userId);

        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            results.Add(new TrackingEvent
            {
                Id = reader.GetInt32(0),
                UserId = reader.GetInt32(1),
                EventType = reader.GetString(2),
                EventData = reader.IsDBNull(3) ? null : reader.GetString(3),
                IpAddress = reader.IsDBNull(4) ? "" : reader.GetString(4),
                UserAgent = reader.IsDBNull(5) ? "" : reader.GetString(5),
                CreatedAt = reader.IsDBNull(6) ? DateTime.UtcNow : reader.GetDateTime(6)
            });
        }

        return results;
    }
}

/// <summary>DI 注册扩展</summary>
public static class TrackingServiceExtensions
{
    public static IServiceCollection AddTrackingService(this IServiceCollection services, string connectionString)
    {
        services.AddSingleton<ITrackingService>(new TrackingService(connectionString));
        return services;
    }
}
