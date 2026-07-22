using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Http;
using System.Collections.Concurrent;
using System.Text.Json;

namespace PerfumeShop.Shared.Services;

/// <summary>
/// V19 缓存服务接口 — 对应 V18 includes/cache_manager.asp + cache_v18_ext.asp
/// L1 IMemoryCache (必选) + L2 IDistributedCache (可选, Redis)
/// </summary>
public interface ICacheService
{
    /// <summary>获取缓存值</summary>
    Task<T?> GetAsync<T>(string key, CancellationToken ct = default);

    /// <summary>设置缓存（绝对过期 + 可选滑动过期）</summary>
    Task SetAsync<T>(string key, T value, TimeSpan? absoluteExpiry = null, TimeSpan? slidingExpiry = null, CancellationToken ct = default);

    /// <summary>删除单个缓存</summary>
    Task RemoveAsync(string key, CancellationToken ct = default);

    /// <summary>按前缀批量删除缓存（级联失效）</summary>
    Task RemoveByPrefixAsync(string prefix, CancellationToken ct = default);

    /// <summary>获取或设置（缓存穿透保护：并发时只执行一次 factory）</summary>
    Task<T> GetOrSetAsync<T>(string key, Func<Task<T>> factory, TimeSpan? absoluteExpiry = null, TimeSpan? slidingExpiry = null, CancellationToken ct = default);

    /// <summary>获取缓存统计</summary>
    CacheStats GetStats();
}

/// <summary>缓存统计</summary>
public class CacheStats
{
    public long Hits { get; set; }
    public long Misses { get; set; }
    public long Sets { get; set; }
    public long Deletes { get; set; }
    public double HitRate => Hits + Misses > 0 ? Math.Round((double)Hits / (Hits + Misses) * 100, 1) : 0;
}

/// <summary>
/// 缓存配置选项
/// </summary>
public class CacheOptions
{
    /// <summary>L1 最大缓存条数（默认 1024）</summary>
    public int L1MaxSize { get; set; } = 1024;

    /// <summary>默认绝对过期时间（秒），默认 600 = 10 分钟</summary>
    public int DefaultTtlSeconds { get; set; } = 600;

    /// <summary>默认滑动过期时间（秒），默认 0 = 不启用</summary>
    public int DefaultSlidingSeconds { get; set; } = 0;

    /// <summary>缓存键前缀</summary>
    public string Prefix { get; set; } = "ps";

    /// <summary>L2 Redis 连接字符串（为空则不启用 L2）</summary>
    public string? RedisConnectionString { get; set; }
}

/// <summary>
/// 两级缓存实现 — L1 MemoryCache + L2 DistributedCache (可选 Redis)
/// 缓存键格式：{prefix}:{key}
/// </summary>
public class CacheService : ICacheService, IDisposable
{
    private readonly IMemoryCache _l1;
    private readonly IDistributedCache? _l2;
    private readonly CacheOptions _options;
    private readonly IHttpContextAccessor? _httpAccessor;

    // 统计
    private long _hits, _misses, _sets, _deletes;

    // 键跟踪（用于前缀删除）
    private readonly ConcurrentDictionary<string, byte> _trackedKeys = new();

    // 并发锁（GetOrSet 防穿透）
    private readonly ConcurrentDictionary<string, SemaphoreSlim> _locks = new();

    public CacheService(IMemoryCache l1, IConfiguration config, IDistributedCache? l2 = null, IHttpContextAccessor? httpAccessor = null)
    {
        _l1 = l1;
        _l2 = l2;
        _httpAccessor = httpAccessor;
        _options = new CacheOptions();
        config.GetSection("Cache").Bind(_options);
    }

    /// <summary>构造（测试用）</summary>
    public CacheService(IMemoryCache l1, CacheOptions options, IDistributedCache? l2 = null, IHttpContextAccessor? httpAccessor = null)
    {
        _l1 = l1;
        _l2 = l2;
        _httpAccessor = httpAccessor;
        _options = options;
    }

    private string FullKey(string key) => $"{_options.Prefix}:{key}";

    /// <summary>标记当前HTTP请求的缓存状态（HIT/MISS），供 CacheHeaderMiddleware 读取</summary>
    private void SetCacheStatus(string status)
    {
        var ctx = _httpAccessor?.HttpContext;
        if (ctx != null)
            ctx.Items["CacheStatus"] = status;
    }

    public Task<T?> GetAsync<T>(string key, CancellationToken ct = default)
    {
        var fk = FullKey(key);

        // L1 查找
        if (_l1.TryGetValue(fk, out T? value))
        {
            Interlocked.Increment(ref _hits);
            SetCacheStatus("HIT");
            return Task.FromResult(value);
        }

        // L2 查找
        if (_l2 != null)
        {
            var json = _l2.GetString(fk);
            if (json != null)
            {
                Interlocked.Increment(ref _hits);
                SetCacheStatus("HIT");
                var deserialized = JsonSerializer.Deserialize<T>(json);
                // 回填 L1
                var entry = _l1.CreateEntry(fk);
                entry.Value = deserialized;
                ApplyExpiration(entry);
                entry.Dispose();
                return Task.FromResult(deserialized);
            }
        }

        Interlocked.Increment(ref _misses);
        SetCacheStatus("MISS");
        return Task.FromResult(default(T?));
    }

    public Task SetAsync<T>(string key, T value, TimeSpan? absoluteExpiry = null, TimeSpan? slidingExpiry = null, CancellationToken ct = default)
    {
        var fk = FullKey(key);

        // L1
        var entry = _l1.CreateEntry(fk);
        entry.Value = value;
        var abs = absoluteExpiry ?? (_options.DefaultTtlSeconds > 0 ? TimeSpan.FromSeconds(_options.DefaultTtlSeconds) : null);
        var sld = slidingExpiry ?? (_options.DefaultSlidingSeconds > 0 ? TimeSpan.FromSeconds(_options.DefaultSlidingSeconds) : null);
        if (abs.HasValue) entry.SetAbsoluteExpiration(abs.Value);
        if (sld.HasValue) entry.SetSlidingExpiration(sld.Value);
        entry.Size = 1;  // SizeLimit 模式下必须设置
        entry.Dispose();

        // L2
        if (_l2 != null)
        {
            var json = JsonSerializer.Serialize(value);
            var distOpts = new DistributedCacheEntryOptions();
            if (abs.HasValue) distOpts.SetAbsoluteExpiration(abs.Value);
            if (sld.HasValue) distOpts.SetSlidingExpiration(sld.Value);
            _l2.SetString(fk, json, distOpts);
        }

        _trackedKeys.TryAdd(fk, 0);
        Interlocked.Increment(ref _sets);
        return Task.CompletedTask;
    }

    public Task RemoveAsync(string key, CancellationToken ct = default)
    {
        var fk = FullKey(key);
        _l1.Remove(fk);
        _l2?.Remove(fk);
        _trackedKeys.TryRemove(fk, out _);
        Interlocked.Increment(ref _deletes);
        return Task.CompletedTask;
    }

    public Task RemoveByPrefixAsync(string prefix, CancellationToken ct = default)
    {
        var fullPrefix = FullKey(prefix);
        // 去掉末尾 * 以支持 "products:*" 模式
        if (fullPrefix.EndsWith("*"))
            fullPrefix = fullPrefix[..^1];

        var toRemove = _trackedKeys.Keys.Where(k => k.StartsWith(fullPrefix, StringComparison.OrdinalIgnoreCase)).ToList();
        foreach (var fk in toRemove)
        {
            _l1.Remove(fk);
            _l2?.Remove(fk);
            _trackedKeys.TryRemove(fk, out _);
            Interlocked.Increment(ref _deletes);
        }
        return Task.CompletedTask;
    }

    public async Task<T> GetOrSetAsync<T>(string key, Func<Task<T>> factory, TimeSpan? absoluteExpiry = null, TimeSpan? slidingExpiry = null, CancellationToken ct = default)
    {
        var existing = await GetAsync<T>(key, ct);
        if (existing is not null)
            return existing;

        var fk = FullKey(key);
        var lockObj = _locks.GetOrAdd(fk, _ => new SemaphoreSlim(1, 1));
        await lockObj.WaitAsync(ct);
        try
        {
            // 双重检查
            existing = await GetAsync<T>(key, ct);
            if (existing is not null)
                return existing;

            var value = await factory();
            await SetAsync(key, value, absoluteExpiry, slidingExpiry, ct);
            SetCacheStatus("SET");
            return value;
        }
        finally
        {
            lockObj.Release();
            _locks.TryRemove(fk, out _);
        }
    }

    public CacheStats GetStats() => new()
    {
        Hits = Interlocked.Read(ref _hits),
        Misses = Interlocked.Read(ref _misses),
        Sets = Interlocked.Read(ref _sets),
        Deletes = Interlocked.Read(ref _deletes)
    };

    private void ApplyExpiration(ICacheEntry entry)
    {
        if (_options.DefaultTtlSeconds > 0)
            entry.SetAbsoluteExpiration(TimeSpan.FromSeconds(_options.DefaultTtlSeconds));
        if (_options.DefaultSlidingSeconds > 0)
            entry.SetSlidingExpiration(TimeSpan.FromSeconds(_options.DefaultSlidingSeconds));
    }

    public void Dispose()
    {
        (_l1 as IDisposable)?.Dispose();
        foreach (var kv in _locks)
            kv.Value.Dispose();
    }
}

/// <summary>DI 扩展方法</summary>
public static class CacheServiceExtensions
{
    public static IServiceCollection AddCacheService(this IServiceCollection services, IConfiguration config)
    {
        var opts = new CacheOptions();
        config.GetSection("Cache").Bind(opts);

        services.AddMemoryCache(o =>
        {
            o.SizeLimit = opts.L1MaxSize;
        });

        // 如果配置了 Redis，注册 IDistributedCache
        if (!string.IsNullOrWhiteSpace(opts.RedisConnectionString))
        {
            services.AddStackExchangeRedisCache(o =>
            {
                o.Configuration = opts.RedisConnectionString;
                o.InstanceName = opts.Prefix + ":";
            });
        }

        services.AddSingleton<ICacheService, CacheService>();
        return services;
    }
}
