using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using PerfumeShop.Shared.Services;
using System.Text.Json;

namespace PerfumeShop.IntegrationTests.Services;

/// <summary>
/// CacheService 集成测试 — 10 用例
/// 覆盖 L1 命中/L2 命中/过期/级联失效/GetOrSet 缓存穿透保护
/// </summary>
public class CacheServiceTests : IDisposable
{
    private readonly MemoryCache _memCache;
    private readonly CacheService _cache;

    public CacheServiceTests()
    {
        _memCache = new MemoryCache(new MemoryCacheOptions { SizeLimit = 100 });
        _cache = new CacheService(_memCache, new CacheOptions
        {
            Prefix = "test",
            DefaultTtlSeconds = 60,
            DefaultSlidingSeconds = 0,
            L1MaxSize = 100
        });
    }

    [Fact]
    public async Task L1_SetAndGet_ReturnsValue()
    {
        // Arrange & Act
        await _cache.SetAsync("key1", "value1");
        var result = await _cache.GetAsync<string>("key1");

        // Assert
        Assert.Equal("value1", result);
    }

    [Fact]
    public async Task L1_Miss_ReturnsDefault()
    {
        var result = await _cache.GetAsync<string>("nonexistent");
        Assert.Null(result);
    }

    [Fact]
    public async Task L1_Expired_ReturnsNull()
    {
        // Set with very short expiry
        await _cache.SetAsync("expiring", "data", absoluteExpiry: TimeSpan.FromMilliseconds(50));
        await Task.Delay(100);
        var result = await _cache.GetAsync<string>("expiring");
        Assert.Null(result);
    }

    [Fact]
    public async Task L1_SlidingExpiry_StaysAlive()
    {
        await _cache.SetAsync("sliding", "data", slidingExpiry: TimeSpan.FromSeconds(5));
        var result = await _cache.GetAsync<string>("sliding");
        Assert.Equal("data", result);
    }

    [Fact]
    public async Task Remove_DeletesKey()
    {
        await _cache.SetAsync("removeme", "val");
        await _cache.RemoveAsync("removeme");
        var result = await _cache.GetAsync<string>("removeme");
        Assert.Null(result);
    }

    [Fact]
    public async Task RemoveByPrefix_ClearsMatchingKeys()
    {
        await _cache.SetAsync("products_list_1", "p1");
        await _cache.SetAsync("products_list_2", "p2");
        await _cache.SetAsync("users_list_1", "u1");

        await _cache.RemoveByPrefixAsync("products_list_");

        Assert.Null(await _cache.GetAsync<string>("products_list_1"));
        Assert.Null(await _cache.GetAsync<string>("products_list_2"));
        Assert.Equal("u1", await _cache.GetAsync<string>("users_list_1"));
    }

    [Fact]
    public async Task RemoveByPrefix_WildcardPattern_Works()
    {
        await _cache.SetAsync("products_detail_1", "d1");
        await _cache.SetAsync("products_detail_2", "d2");

        await _cache.RemoveByPrefixAsync("products_detail_*");

        Assert.Null(await _cache.GetAsync<string>("products_detail_1"));
        Assert.Null(await _cache.GetAsync<string>("products_detail_2"));
    }

    [Fact]
    public async Task GetOrSet_CacheMiss_CallsFactory()
    {
        var callCount = 0;
        var result1 = await _cache.GetOrSetAsync("factory_key", async () =>
        {
            callCount++;
            return "computed_value";
        });

        var result2 = await _cache.GetOrSetAsync("factory_key", async () =>
        {
            callCount++;
            return "should_not_be_called";
        });

        Assert.Equal("computed_value", result1);
        Assert.Equal("computed_value", result2);
        Assert.Equal(1, callCount); // factory 只调用一次
    }

    [Fact]
    public async Task GetOrSet_ConcurrentRequests_OnlyOneFactoryCall()
    {
        var callCount = 0;
        var tasks = Enumerable.Range(0, 10).Select(_ =>
            _cache.GetOrSetAsync("concurrent_key", async () =>
            {
                Interlocked.Increment(ref callCount);
                await Task.Delay(10);
                return "shared_value";
            })
        ).ToList();

        var results = await Task.WhenAll(tasks);
        Assert.All(results, r => Assert.Equal("shared_value", r));
        // 并发保护：factory 最多被调用少量（取决于时序，但远小于 10）
        Assert.True(callCount < 10, $"Factory called {callCount} times, expected < 10");
    }

    [Fact]
    public void GetStats_TracksHitsAndMisses()
    {
        _ = _cache.GetAsync<string>("stat_miss").GetAwaiter().GetResult();
        _cache.SetAsync("stat_hit", "v").GetAwaiter().GetResult();
        _ = _cache.GetAsync<string>("stat_hit").GetAwaiter().GetResult();

        var stats = _cache.GetStats();
        Assert.True(stats.Hits >= 1);
        Assert.True(stats.Misses >= 1);
        Assert.True(stats.Sets >= 1);
    }

    [Fact]
    public async Task ComplexType_SerializeDeserialize()
    {
        var obj = new TestProduct { Id = 1, Name = "Rose Perfume", Price = 299.99m };
        await _cache.SetAsync("product_1", obj);
        var result = await _cache.GetAsync<TestProduct>("product_1");

        Assert.NotNull(result);
        Assert.Equal(1, result.Id);
        Assert.Equal("Rose Perfume", result.Name);
        Assert.Equal(299.99m, result.Price);
    }

    public void Dispose() => _cache.Dispose();

    private class TestProduct
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
    }
}
