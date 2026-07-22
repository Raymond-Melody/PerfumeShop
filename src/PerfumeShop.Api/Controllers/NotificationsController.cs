using Microsoft.AspNetCore.Mvc;
using System.Collections.Concurrent;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// SSE 实时通知 + 邮件发送 API — 对应 V18 api/notifications_sse.asp + email_service.asp
/// </summary>
[ApiController]
[Route("api/notifications")]
public class NotificationsController : ControllerBase
{
    private static readonly ConcurrentDictionary<int, StreamWriter> _sseClients = new();

    /// <summary>GET /api/notifications/stream — SSE 实时通知流</summary>
    [HttpGet("stream")]
    public async Task Stream([FromQuery] int userId = 0, CancellationToken ct = default)
    {
        Response.StatusCode = 200;
        Response.ContentType = "text/event-stream";
        Response.Headers["Cache-Control"] = "no-cache";
        Response.Headers["Connection"] = "keep-alive";
        Response.Headers["X-Accel-Buffering"] = "no";

        var writer = new StreamWriter(Response.Body);
        if (userId > 0) _sseClients[userId] = writer;

        try
        {
            // 发送初始连接事件
            await writer.WriteLineAsync($"event: connected\ndata: {{\"userId\":{userId},\"timestamp\":\"{DateTime.UtcNow:o}\"}}\n");
            await writer.FlushAsync(ct);

            // 心跳保持连接
            while (!ct.IsCancellationRequested)
            {
                await Task.Delay(15000, ct);
                await writer.WriteLineAsync($": heartbeat {DateTime.UtcNow:HH:mm:ss}\n");
                await writer.FlushAsync(ct);
            }
        }
        catch (OperationCanceledException) { }
        finally
        {
            if (userId > 0) _sseClients.TryRemove(userId, out _);
        }
    }

    /// <summary>POST /api/notifications/send — 向指定用户发送通知</summary>
    [HttpPost("send")]
    public async Task<IActionResult> Send([FromBody] NotificationRequest req)
    {
        if (req.UserId <= 0 || string.IsNullOrWhiteSpace(req.Message))
            return BadRequest(new { success = false, message = "参数不完整" });

        if (_sseClients.TryGetValue(req.UserId, out var writer))
        {
            try
            {
                var data = System.Text.Json.JsonSerializer.Serialize(new
                {
                    type = req.Type ?? "info",
                    message = req.Message,
                    title = req.Title ?? "",
                    url = req.Url ?? "",
                    timestamp = DateTime.UtcNow.ToString("o")
                });
                await writer.WriteLineAsync($"event: notification\ndata: {data}\n");
                await writer.FlushAsync();
            }
            catch { }
        }

        return Ok(new { success = true, message = "通知已发送" });
    }

    /// <summary>POST /api/notifications/broadcast — 广播通知给所有在线用户</summary>
    [HttpPost("broadcast")]
    public async Task<IActionResult> Broadcast([FromBody] NotificationRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Message))
            return BadRequest(new { success = false, message = "消息不能为空" });

        var data = System.Text.Json.JsonSerializer.Serialize(new
        {
            type = req.Type ?? "broadcast",
            message = req.Message,
            title = req.Title ?? "",
            url = req.Url ?? "",
            timestamp = DateTime.UtcNow.ToString("o")
        });

        foreach (var kvp in _sseClients)
        {
            try
            {
                await kvp.Value.WriteLineAsync($"event: notification\ndata: {data}\n");
                await kvp.Value.FlushAsync();
            }
            catch { }
        }

        return Ok(new { success = true, message = $"已广播给 {_sseClients.Count} 个在线用户" });
    }

    /// <summary>GET /api/notifications/status — SSE 连接状态</summary>
    [HttpGet("status")]
    public IActionResult Status()
    {
        return Ok(new { success = true, activeConnections = _sseClients.Count });
    }
}

public class NotificationRequest
{
    public int UserId { get; set; }
    public string? Type { get; set; }
    public string? Title { get; set; }
    public string Message { get; set; } = "";
    public string? Url { get; set; }
}
