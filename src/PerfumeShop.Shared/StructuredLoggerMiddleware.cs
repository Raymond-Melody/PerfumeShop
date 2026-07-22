using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace PerfumeShop.Shared;

/// <summary>
/// V19 结构化日志 + 增强指标采集中间件
/// 对应 V18 includes/logger.asp + includes/metrics.asp
/// 新增：P50/P95/P99延迟统计、慢请求检测、转化漏斗
/// </summary>
public class StructuredLoggerMiddleware
{
    private readonly RequestDelegate _next;
    private readonly ILogger<StructuredLoggerMiddleware> _logger;

    private static long _totalRequests;
    private static long _totalErrors;
    private static long _slowRequests;
    private static readonly ConcurrentDictionary<string, long> _endpointHits = new();

    // V19: 延迟采样（每端点保留最近100个样本用于P50/P95/P99计算）
    private static readonly ConcurrentDictionary<string, CircularLatencyBuffer> _latencySamples = new();

    // V19: 慢请求阈值
    private const int SlowPageThresholdMs = 2000;
    private const int SlowApiThresholdMs = 500;

    // V19: 转化漏斗计数器（对齐 V18 metrics.asp METRICS_TrackFunnel）
    private static readonly ConcurrentDictionary<string, long> _funnelCounts = new();
    private static readonly string[] FunnelSteps = { "view_product", "add_to_cart", "begin_checkout", "purchase" };

    public StructuredLoggerMiddleware(RequestDelegate next, ILogger<StructuredLoggerMiddleware> logger)
    {
        _next = next;
        _logger = logger;

        // 初始化漏斗计数器
        foreach (var step in FunnelSteps)
            _funnelCounts.TryAdd(step, 0);
    }

    public async Task InvokeAsync(HttpContext context)
    {
        var sw = Stopwatch.StartNew();
        Interlocked.Increment(ref _totalRequests);

        var path = context.Request.Path.Value?.ToLower() ?? "/";
        _endpointHits.AddOrUpdate(path, 1, (_, v) => v + 1);

        try
        {
            await _next(context);

            sw.Stop();
            var elapsed = sw.ElapsedMilliseconds;
            var statusCode = context.Response.StatusCode;

            // V19: 记录延迟样本
            RecordLatency(path, elapsed);

            // V19: 慢请求检测（对齐 V18 metrics.asp 阈值）
            var isApi = path.StartsWith("/api");
            var slowThreshold = isApi ? SlowApiThresholdMs : SlowPageThresholdMs;
            if (elapsed > slowThreshold)
            {
                Interlocked.Increment(ref _slowRequests);
                _logger.LogWarning(
                    "[SLOW] {Method} {Path} → {StatusCode} ({Elapsed}ms, Threshold={Threshold}ms) | IP: {IP}",
                    context.Request.Method, path, statusCode, elapsed, slowThreshold,
                    context.Connection.RemoteIpAddress);
            }

            if (statusCode >= 400)
            {
                Interlocked.Increment(ref _totalErrors);
                _logger.LogWarning(
                    "[{Timestamp:yyyy-MM-dd HH:mm:ss}] {Method} {Path} → {StatusCode} ({Elapsed}ms) | IP: {IP} | UA: {UA}",
                    DateTime.Now, context.Request.Method, path, statusCode, elapsed,
                    context.Connection.RemoteIpAddress, Truncate(context.Request.Headers.UserAgent.ToString(), 100));
            }
            else
            {
                _logger.LogInformation(
                    "[{Timestamp:yyyy-MM-dd HH:mm:ss}] {Method} {Path} → {StatusCode} ({Elapsed}ms)",
                    DateTime.Now, context.Request.Method, path, statusCode, elapsed);
            }
        }
        catch (Exception ex)
        {
            sw.Stop();
            Interlocked.Increment(ref _totalErrors);
            _logger.LogError(ex,
                "[{Timestamp:yyyy-MM-dd HH:mm:ss}] ERROR {Method} {Path} ({Elapsed}ms): {Message}",
                DateTime.Now, context.Request.Method, path, sw.ElapsedMilliseconds, ex.Message);
            throw;
        }
    }

    // V19: 记录请求延迟到环形缓冲区
    private static void RecordLatency(string path, long elapsed)
    {
        var buffer = _latencySamples.GetOrAdd(path, _ => new CircularLatencyBuffer(100));
        buffer.Add(elapsed);
    }

    // V19: 计算延迟百分位数
    private static (long p50, long p95, long p99) CalculatePercentiles(string path)
    {
        if (!_latencySamples.TryGetValue(path, out var buffer))
            return (0, 0, 0);

        var samples = buffer.GetSamples();
        if (samples.Length == 0) return (0, 0, 0);

        Array.Sort(samples);
        return (
            samples[Math.Min((int)(samples.Length * 0.50), samples.Length - 1)],
            samples[Math.Min((int)(samples.Length * 0.95), samples.Length - 1)],
            samples[Math.Min((int)(samples.Length * 0.99), samples.Length - 1)]
        );
    }

    /// <summary>GET /api/metrics — 获取增强系统指标</summary>
    public static IResult GetMetrics()
    {
        var topEndpoints = _endpointHits
            .OrderByDescending(kvp => kvp.Value)
            .Take(20)
            .Select(kvp =>
            {
                var (p50, p95, p99) = CalculatePercentiles(kvp.Key);
                return new
                {
                    endpoint = kvp.Key,
                    hits = kvp.Value,
                    latency = new { p50, p95, p99 }
                };
            });

        // V19: 慢请求统计
        var topSlowEndpoints = _latencySamples
            .OrderByDescending(kvp => kvp.Value.GetAverage())
            .Take(10)
            .Select(kvp => new { endpoint = kvp.Key, avgLatencyMs = kvp.Value.GetAverage() });

        // V19: 漏斗转化率
        var funnelStats = _funnelCounts.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

        return Results.Ok(new
        {
            status = "ok",
            uptime = (DateTime.Now - Process.GetCurrentProcess().StartTime).ToString(@"d\.hh\:mm\:ss"),
            process = new
            {
                memoryMB = Math.Round(GC.GetTotalMemory(false) / 1024.0 / 1024.0, 2),
                gcCollections = new
                {
                    gen0 = GC.CollectionCount(0),
                    gen1 = GC.CollectionCount(1),
                    gen2 = GC.CollectionCount(2)
                },
                threads = Environment.ProcessorCount,
                startTime = Process.GetCurrentProcess().StartTime.ToString("o")
            },
            requests = new
            {
                total = Interlocked.Read(ref _totalRequests),
                errors = Interlocked.Read(ref _totalErrors),
                slow = Interlocked.Read(ref _slowRequests),
                errorRate = Interlocked.Read(ref _totalRequests) > 0
                    ? Math.Round((double)Interlocked.Read(ref _totalErrors) / Interlocked.Read(ref _totalRequests) * 100, 2)
                    : 0,
                slowRate = Interlocked.Read(ref _totalRequests) > 0
                    ? Math.Round((double)Interlocked.Read(ref _slowRequests) / Interlocked.Read(ref _totalRequests) * 100, 2)
                    : 0
            },
            topEndpoints,
            topSlowEndpoints,
            funnel = funnelStats
        });
    }

    /// <summary>POST /api/metrics/funnel — 记录漏斗步骤（对齐 V18 METRICS_TrackFunnel）</summary>
    public static void TrackFunnel(string step)
    {
        _funnelCounts.AddOrUpdate(step, 1, (_, v) => v + 1);
    }

    private static string Truncate(string value, int maxLength)
    {
        if (string.IsNullOrEmpty(value)) return "";
        return value.Length <= maxLength ? value : value[..maxLength];
    }

    /// <summary>
    /// V19: 环形延迟缓冲区 — 固定大小，自动覆盖最旧样本
    /// </summary>
    private class CircularLatencyBuffer
    {
        private readonly long[] _buffer;
        private int _index;
        private int _count;
        private long _sum;

        public CircularLatencyBuffer(int capacity)
        {
            _buffer = new long[capacity];
        }

        public void Add(long value)
        {
            lock (_buffer)
            {
                // 如果覆盖旧值，从 sum 中减去
                if (_count >= _buffer.Length)
                {
                    _sum -= _buffer[_index];
                }
                _buffer[_index] = value;
                _sum += value;
                _index = (_index + 1) % _buffer.Length;
                if (_count < _buffer.Length) _count++;
            }
        }

        public long[] GetSamples()
        {
            lock (_buffer)
            {
                var result = new long[_count];
                for (int i = 0; i < _count; i++)
                {
                    var idx = (_index - _count + i + _buffer.Length) % _buffer.Length;
                    result[i] = _buffer[idx];
                }
                return result;
            }
        }

        public double GetAverage()
        {
            lock (_buffer)
            {
                return _count > 0 ? (double)_sum / _count : 0;
            }
        }
    }
}

/// <summary>扩展方法：注册结构化日志中间件和增强指标端点</summary>
public static class StructuredLoggerExtensions
{
    public static IApplicationBuilder UseStructuredLogging(this IApplicationBuilder builder)
    {
        return builder.UseMiddleware<StructuredLoggerMiddleware>();
    }

    public static void MapMetricsEndpoint(this WebApplication app)
    {
        app.MapGet("/api/metrics", () => StructuredLoggerMiddleware.GetMetrics());

        // V19: 漏斗埋点端点（对齐 V18 metrics.asp METRICS_TrackFunnel）
        app.MapPost("/api/metrics/funnel", (FunnelRequest req) =>
        {
            var validSteps = new[] { "view_product", "add_to_cart", "begin_checkout", "purchase" };
            if (!validSteps.Contains(req.Step))
            {
                return Results.BadRequest(new { error = "Invalid step. Must be one of: " + string.Join(", ", validSteps) });
            }
            StructuredLoggerMiddleware.TrackFunnel(req.Step);
            return Results.Ok(new { success = true, step = req.Step });
        });
    }
}

/// <summary>V19: 漏斗埋点请求体</summary>
public record FunnelRequest(string Step);
