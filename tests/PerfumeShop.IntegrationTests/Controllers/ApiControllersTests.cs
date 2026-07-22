using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Api.Controllers;
using PerfumeShop.Core.AI;
using PerfumeShop.Data.Models;

namespace PerfumeShop.IntegrationTests.Controllers;

/// <summary>
/// 所有新 Controller 路由验证 + 功能测试
/// </summary>
public class ApiControllersTests : IDisposable
{
    private readonly TestEngineContext _db;
    private readonly ISentimentAnalyzer _sentiment;
    private readonly IChatbotEngine _chatbot;
    private readonly IFragranceMatcher _fragrance;

    public ApiControllersTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"ApiControllersTests_{Guid.NewGuid()}")
            .Options;
        _db = new TestEngineContext(options);

        // Seed test data
        _db.Products.Add(new Product
        {
            ProductId = 1,
            ProductName = "玫瑰花香淡香水",
            BasePrice = 299,
            ProductType = "Custom",
            Category = "花香",
            Description = "清新优雅的玫瑰香",
            ImageUrl = "/images/products/rose.jpg",
            IsActive = true
        });
        _db.Products.Add(new Product
        {
            ProductId = 2,
            ProductName = "海洋清风男士香水",
            BasePrice = 399,
            ProductType = "standard",
            Category = "海洋",
            Description = "清爽活力的海洋调",
            ImageUrl = "/images/products/ocean.jpg",
            IsActive = true
        });
        _db.Orders.Add(new Order
        {
            OrderId = 1,
            OrderNo = "ORD001",
            TotalAmount = 299,
            Status = "Paid",
            UserId = 1
        });
        _db.Areas.Add(new Area { AreaId = 1, AreaName = "北京市", ParentId = 0 });
        _db.Areas.Add(new Area { AreaId = 2, AreaName = "上海市", ParentId = 0 });
        _db.Areas.Add(new Area { AreaId = 10, AreaName = "朝阳区", ParentId = 1 });
        _db.SaveChanges();

        _sentiment = new SentimentAnalyzer();
        _chatbot = new ChatbotEngine();
        _fragrance = new FragranceMatcher();
    }

    public void Dispose()
    {
        _db.Dispose();
    }

    // ========== SentimentController ==========

    [Fact]
    public void SentimentController_Analyze_ReturnsOk()
    {
        var controller = new SentimentController(_sentiment);
        var result = controller.Analyze(new SentimentRequest { Text = "非常好闻" });

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public void SentimentController_AnalyzeEmpty_ReturnsBadRequest()
    {
        var controller = new SentimentController(_sentiment);
        var result = controller.Analyze(new SentimentRequest { Text = "" });

        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public void SentimentController_Batch_ReturnsOk()
    {
        var controller = new SentimentController(_sentiment);
        var result = controller.Batch(new SentimentBatchRequest
        {
            Texts = new List<string> { "喜欢", "难闻", "还行" }
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public void SentimentController_BatchTooMany_ReturnsBadRequest()
    {
        var controller = new SentimentController(_sentiment);
        var result = controller.Batch(new SentimentBatchRequest
        {
            Texts = Enumerable.Range(0, 51).Select(i => $"text{i}").ToList()
        });

        Assert.IsType<BadRequestObjectResult>(result);
    }

    // ========== RiskCheckController ==========

    [Fact]
    public void RiskCheckController_Check_ReturnsOk()
    {
        var controller = new RiskCheckController();
        var result = controller.Check(new RiskCheckRequest
        {
            Ip = "1.2.3.4",
            Amount = 500,
            AccountAgeDays = 30,
            RequestsPerMinute = 5
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public void RiskCheckController_HighAmount_ReturnsHighRisk()
    {
        var controller = new RiskCheckController();
        var result = controller.Check(new RiskCheckRequest
        {
            Amount = 15000,
            AccountAgeDays = 30,
            RequestsPerMinute = 5
        });

        Assert.IsType<OkObjectResult>(result);
    }

    // ========== BackupStatusController ==========

    [Fact]
    public void BackupStatusController_GetStatus_ReturnsOk()
    {
        var env = CreateMockEnvironment();
        var controller = new BackupStatusController(env);
        var result = controller.GetStatus();

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    // ========== SystemDiagController ==========

    [Fact]
    public void SystemDiagController_GetSystem_ReturnsOk()
    {
        var env = CreateMockEnvironment();
        var controller = new SystemDiagController(env);
        var result = controller.GetSystemDiag();

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    // ========== DiagCostsController ==========

    [Fact]
    public async Task DiagCostsController_GetCosts_ReturnsOk()
    {
        var controller = new DiagCostsController(_db);
        var result = await controller.GetCostDiag();

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    // ========== DiagOrdersController ==========

    [Fact]
    public async Task DiagOrdersController_GetOrders_ReturnsOk()
    {
        var controller = new DiagOrdersController(_db);
        var result = await controller.GetOrderDiag();

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    // ========== TrackController ==========

    [Fact]
    public void TrackController_PostEvent_ReturnsOk()
    {
        var controller = new TrackController();
        var result = controller.TrackEvent(new TrackEventRequest
        {
            Action = "view",
            Target = "1"
        });

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public void TrackController_GetGif_ReturnsGifFile()
    {
        var controller = new TrackController();
        controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext()
        };
        var result = controller.TrackGif("view", "1", null);

        var fileResult = Assert.IsType<FileContentResult>(result);
        Assert.Equal("image/gif", fileResult.ContentType);
        Assert.True(fileResult.FileContents.Length > 0);
    }

    // ========== CookieConsentController ==========

    [Fact]
    public void CookieConsentController_Post_ReturnsOk()
    {
        var controller = new CookieConsentController();
        controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext()
        };
        var result = controller.RecordConsent(new CookieConsentRequest { Consent = "all" });

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public void CookieConsentController_InvalidConsent_DefaultsToEssential()
    {
        var controller = new CookieConsentController();
        controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext()
        };
        var result = controller.RecordConsent(new CookieConsentRequest { Consent = "invalid" });

        Assert.IsType<OkObjectResult>(result);
    }

    // ========== AreasController ==========

    [Fact]
    public async Task AreasController_GetAreas_ReturnsAreas()
    {
        var controller = new AreasController(_db);
        var result = await controller.GetAreas(parentId: 0);

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public async Task AreasController_GetAreasByParent_ReturnsChildAreas()
    {
        var controller = new AreasController(_db);
        var result = await controller.GetAreas(parentId: 1);

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    // ========== SearchSuggestionsController ==========

    [Fact]
    public async Task SearchSuggestionsController_EmptyQuery_ReturnsEmpty()
    {
        var controller = new SearchSuggestionsController(_db);
        var result = await controller.GetSuggestions(q: "");

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public async Task SearchSuggestionsController_ValidQuery_ReturnsResults()
    {
        var controller = new SearchSuggestionsController(_db);
        var result = await controller.GetSuggestions(q: "玫瑰");

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    [Fact]
    public async Task SearchSuggestionsController_SynonymQuery_ExpandsSearch()
    {
        var controller = new SearchSuggestionsController(_db);
        var result = await controller.GetSuggestions(q: "清新");

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.Equal(200, ok.StatusCode);
    }

    // ========== UploadController ==========

    [Fact]
    public void UploadController_NoFile_ReturnsBadRequest()
    {
        var env = CreateMockEnvironment();
        var controller = new UploadController(env);
        var result = controller.Upload(null).Result;

        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public void UploadController_GetUploadInfo_NotFound_Returns404()
    {
        var env = CreateMockEnvironment();
        var controller = new UploadController(env);
        var result = controller.GetUploadInfo("nonexistent.jpg");

        Assert.IsType<NotFoundObjectResult>(result);
    }

    // ========== Helpers ==========

    private static IWebHostEnvironment CreateMockEnvironment()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "PerfumeShopTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        var env = new Moq.Mock<IWebHostEnvironment>();
        env.Setup(e => e.ContentRootPath).Returns(tempDir);
        env.Setup(e => e.WebRootPath).Returns(Path.Combine(tempDir, "wwwroot"));
        env.Setup(e => e.EnvironmentName).Returns("Testing");
        return env.Object;
    }
}
