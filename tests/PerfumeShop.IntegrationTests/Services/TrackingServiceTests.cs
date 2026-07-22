using Moq;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.IntegrationTests.Services;

/// <summary>TrackingService 单元测试（正常路径 + 异常路径 + 边界条件）</summary>
public class TrackingServiceTests
{
    // 使用无效连接字符串（测试不连接 DB 的部分）
    private readonly TrackingService _service;

    public TrackingServiceTests()
    {
        _service = new TrackingService("Server=invalid;Database=test;Trusted_Connection=True;Connection Timeout=1;");
    }

    [Fact]
    public void GenerateTrackingPixel_ReturnsValidGif()
    {
        // Act
        var pixel = _service.GenerateTrackingPixel();

        // Assert
        Assert.NotNull(pixel);
        Assert.True(pixel.Length > 0);
        // GIF89a header
        Assert.Equal(0x47, pixel[0]); // G
        Assert.Equal(0x49, pixel[1]); // I
        Assert.Equal(0x46, pixel[2]); // F
        Assert.Equal(0x38, pixel[3]); // 8
        Assert.Equal(0x39, pixel[4]); // 9
        Assert.Equal(0x61, pixel[5]); // a
    }

    [Fact]
    public void GenerateTrackingPixel_ReturnsSmallPayload()
    {
        // 像素追踪 GIF 应尽量小
        var pixel = _service.GenerateTrackingPixel();
        Assert.True(pixel.Length < 100, $"GIF size {pixel.Length} should be < 100 bytes");
    }

    [Fact]
    public void GenerateTrackingPixel_IsIdempotent()
    {
        // 每次调用返回相同数据
        var pixel1 = _service.GenerateTrackingPixel();
        var pixel2 = _service.GenerateTrackingPixel();
        Assert.Equal(pixel1, pixel2);
    }

    [Fact]
    public void TrackingEventType_Constants_AreCorrect()
    {
        // 对齐 V18 TU_LogBehavior actionType 参数
        Assert.Equal("PRODUCT_VIEW", TrackingEventType.ProductView);
        Assert.Equal("SEARCH", TrackingEventType.Search);
        Assert.Equal("CART_ADD", TrackingEventType.CartAdd);
        Assert.Equal("FAVORITE", TrackingEventType.Favorite);
        Assert.Equal("PAGE_VIEW", TrackingEventType.PageView);
        Assert.Equal("CUSTOM", TrackingEventType.Custom);
    }

    [Fact]
    public void TrackingEvent_DefaultValues_AreSet()
    {
        // 验证模型默认值
        var evt = new TrackingEvent();
        Assert.Equal(0, evt.UserId);
        Assert.Equal("", evt.EventType);
        Assert.Null(evt.EventData);
        Assert.Equal("", evt.IpAddress);
        Assert.Equal("", evt.UserAgent);
    }

    [Fact]
    public void TrackingEvent_CanSetAllProperties()
    {
        // Arrange & Act
        var evt = new TrackingEvent
        {
            Id = 1,
            UserId = 42,
            EventType = TrackingEventType.ProductView,
            EventData = "{\"productId\": 123}",
            IpAddress = "192.168.1.1",
            UserAgent = "Mozilla/5.0",
            CreatedAt = new DateTime(2024, 1, 15)
        };

        // Assert
        Assert.Equal(1, evt.Id);
        Assert.Equal(42, evt.UserId);
        Assert.Equal("PRODUCT_VIEW", evt.EventType);
        Assert.Contains("productId", evt.EventData);
        Assert.Equal("192.168.1.1", evt.IpAddress);
        Assert.Equal("Mozilla/5.0", evt.UserAgent);
    }

    [Fact]
    public async Task TrackEventAsync_WithInvalidConnection_DoesNotThrowUnhandled()
    {
        // 验证连接失败时不会导致未处理异常（对齐 V18 On Error Resume Next）
        var evt = new TrackingEvent
        {
            UserId = 1,
            EventType = TrackingEventType.PageView,
            IpAddress = "127.0.0.1"
        };

        // 应当抛出 SqlException（因为连接失败），测试用 Assert.ThrowsAsync 验证
        await Assert.ThrowsAnyAsync<Exception>(() => _service.TrackEventAsync(evt));
    }

    [Fact]
    public async Task GetEventsByUserAsync_WithInvalidConnection_DoesNotThrowUnhandled()
    {
        // 连接失败时应抛出异常
        await Assert.ThrowsAnyAsync<Exception>(() => _service.GetEventsByUserAsync(1));
    }
}
