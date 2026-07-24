using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Moq;
using PerfumeShop.Api.Controllers;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Shared.Services;

namespace PerfumeShop.IntegrationTests.Pages;

/// <summary>
/// M3-A 用户认证与订单管理 E2E 测试
/// 覆盖: Register、ForgotPassword、ResetPassword、Orders分页、OrderDetail、未登录重定向
/// </summary>
public class UserPagesTests : IDisposable
{
    private readonly TestEngineContext _db;
    private readonly AuthController _authCtrl;
    private readonly OrdersController _ordersCtrl;
    private readonly Mock<IUserRepository> _userRepoMock;
    private readonly Mock<IOrderRepository> _orderRepoMock;
    private readonly Mock<IEmailService> _emailMock;
    private const int TestUserId = 1;

    private static readonly EmailResult OkEmail = new() { Success = true };

    /// <summary>从 OkObjectResult 提取匿名对象为 JsonElement</summary>
    private static JsonElement ToJson(IActionResult result)
    {
        var ok = Assert.IsType<OkObjectResult>(result);
        var json = JsonSerializer.Serialize(ok.Value);
        return JsonDocument.Parse(json).RootElement;
    }

    public UserPagesTests()
    {
        var options = new DbContextOptionsBuilder<PerfumeShopContext>()
            .UseInMemoryDatabase(databaseName: $"UserPagesTests_{Guid.NewGuid()}")
            .ConfigureWarnings(w => w.Ignore(InMemoryEventId.TransactionIgnoredWarning))
            .Options;
        _db = new TestEngineContext(options);

        _userRepoMock = new Mock<IUserRepository>();
        _orderRepoMock = new Mock<IOrderRepository>();
        _emailMock = new Mock<IEmailService>();

        _authCtrl = new AuthController(_userRepoMock.Object, _db, _emailMock.Object);
        _authCtrl.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext
            {
                Request = { Scheme = "https", Host = new HostString("localhost") }
            }
        };

        var paymentMock = new Mock<PerfumeShop.Core.Interfaces.IPaymentHandler>();
        var promoMock = new Mock<PerfumeShop.Core.Interfaces.IPromotionEngine>();
        var costMock = new Mock<PerfumeShop.Core.Interfaces.ICostEngine>();
        _ordersCtrl = new OrdersController(_orderRepoMock.Object, paymentMock.Object, promoMock.Object, costMock.Object, _db);

        SeedData();
    }

    private void SeedData()
    {
        var user = new User
        {
            UserId = TestUserId,
            Username = "testuser",
            Email = "test@example.com",
            Password = PerfumeShop.Shared.Security.PasswordHasher.Hash("correctpassword"),
            FullName = "Test User",
            Phone = "13800138000",
            IsActive = true,
            CustomerTier = "bronze",
            Points = 100,
            CreatedAt = DateTime.Now
        };
        _db.Users.Add(user);

        for (int i = 1; i <= 15; i++)
        {
            _db.Orders.Add(new Order
            {
                OrderId = i,
                OrderNo = $"ORD20240101{i:D4}",
                UserId = TestUserId,
                Status = i <= 5 ? "Pending" : i <= 10 ? "Paid" : "Completed",
                TotalAmount = 100m * i,
                PaymentMethod = "online",
                ShippingName = "Test User",
                ShippingAddress = "Test Address",
                ShippingCity = "Shanghai",
                ShippingPhone = "13800138000",
                CreatedAt = DateTime.Now.AddDays(-i)
            });
        }

        _db.OrderDetails.Add(new OrderDetail
        {
            DetailId = 1,
            OrderId = 1,
            ProductId = 1,
            ProductName = "Chanel No.5",
            Quantity = 2,
            UnitPrice = 500m,
            Subtotal = 1000m,
            VolumeName = "50ml",
            VolumeMl = 50
        });

        _db.SaveChanges();
    }

    public void Dispose() => _db.Dispose();

    // ========== Register 测试 ==========

    [Fact]
    public async Task Register_ValidRequest_ReturnsOk()
    {
        _userRepoMock.Setup(r => r.UsernameExistsAsync("newuser", It.IsAny<CancellationToken>())).ReturnsAsync(false);
        _userRepoMock.Setup(r => r.EmailExistsAsync("new@example.com", It.IsAny<CancellationToken>())).ReturnsAsync(false);
        _userRepoMock.Setup(r => r.AddAsync(It.IsAny<User>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((User u, CancellationToken _) => { u.UserId = 99; return u; });
        _userRepoMock.Setup(r => r.SaveChangesAsync(It.IsAny<CancellationToken>())).ReturnsAsync(1);
        _emailMock.Setup(e => e.SendTemplateAsync(It.IsAny<string>(), It.IsAny<EmailTemplateModel>(), It.IsAny<string>()))
            .ReturnsAsync(OkEmail);

        var el = ToJson(await _authCtrl.Register(new RegisterDto
        {
            Username = "newuser",
            Password = "Test1234",
            Email = "new@example.com",
            FullName = "New User",
            Phone = "13900139000"
        }));
        Assert.Equal("注册成功", el.GetProperty("message").GetString());
    }

    [Fact]
    public async Task Register_WeakPassword_ReturnsBadRequest()
    {
        var result = await _authCtrl.Register(new RegisterDto
        {
            Username = "weakuser",
            Password = "123",
            Email = "weak@example.com"
        });

        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public async Task Register_DuplicateUsername_ReturnsConflict()
    {
        _userRepoMock.Setup(r => r.UsernameExistsAsync("testuser", It.IsAny<CancellationToken>())).ReturnsAsync(true);

        var result = await _authCtrl.Register(new RegisterDto
        {
            Username = "testuser",
            Password = "Test1234",
            Email = "dup@example.com"
        });

        Assert.IsType<ConflictObjectResult>(result);
    }

    [Fact]
    public async Task Register_DuplicateEmail_ReturnsConflict()
    {
        _userRepoMock.Setup(r => r.UsernameExistsAsync("unique", It.IsAny<CancellationToken>())).ReturnsAsync(false);
        _userRepoMock.Setup(r => r.EmailExistsAsync("test@example.com", It.IsAny<CancellationToken>())).ReturnsAsync(true);

        var result = await _authCtrl.Register(new RegisterDto
        {
            Username = "unique",
            Password = "Test1234",
            Email = "test@example.com"
        });

        Assert.IsType<ConflictObjectResult>(result);
    }

    // ========== ForgotPassword 测试 ==========

    [Fact]
    public async Task ForgotPassword_ValidEmail_ReturnsOk()
    {
        var existingUser = _db.Users.First(u => u.Email == "test@example.com");
        _userRepoMock.Setup(r => r.GetByEmailAsync("test@example.com", It.IsAny<CancellationToken>()))
            .ReturnsAsync(existingUser);
        _emailMock.Setup(e => e.SendTemplateAsync(It.IsAny<string>(), It.IsAny<EmailTemplateModel>(), It.IsAny<string>()))
            .ReturnsAsync(OkEmail);

        var result = await _authCtrl.RequestPasswordReset(new ForgotPasswordDto { Email = "test@example.com" });

        Assert.IsType<OkObjectResult>(result);
        var tokens = _db.PasswordResetTokens.Where(t => t.UserId == TestUserId).ToList();
        Assert.NotEmpty(tokens);
        Assert.False(tokens[0].IsUsed);
        Assert.True(tokens[0].ExpiresAt > DateTime.Now);
    }

    [Fact]
    public async Task ForgotPassword_UnknownEmail_ReturnsNotFound()
    {
        _userRepoMock.Setup(r => r.GetByEmailAsync("unknown@example.com", It.IsAny<CancellationToken>()))
            .ReturnsAsync((User?)null);

        var result = await _authCtrl.RequestPasswordReset(new ForgotPasswordDto { Email = "unknown@example.com" });

        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public async Task ForgotPassword_EmptyEmail_ReturnsBadRequest()
    {
        var result = await _authCtrl.RequestPasswordReset(new ForgotPasswordDto { Email = "" });
        Assert.IsType<BadRequestObjectResult>(result);
    }

    // ========== ResetPassword 测试 ==========

    [Fact]
    public async Task ResetPassword_ValidToken_UpdatesPassword()
    {
        // 先通过 ForgotPassword 创建 Token
        var existingUser = _db.Users.First(u => u.Email == "test@example.com");
        _userRepoMock.Setup(r => r.GetByEmailAsync("test@example.com", It.IsAny<CancellationToken>()))
            .ReturnsAsync(existingUser);
        _emailMock.Setup(e => e.SendTemplateAsync(It.IsAny<string>(), It.IsAny<EmailTemplateModel>(), It.IsAny<string>()))
            .ReturnsAsync(OkEmail);
        await _authCtrl.RequestPasswordReset(new ForgotPasswordDto { Email = "test@example.com" });

        var token = _db.PasswordResetTokens.First(t => t.UserId == TestUserId).Token;

        _userRepoMock.Setup(r => r.GetByIdAsync(TestUserId, It.IsAny<CancellationToken>()))
            .ReturnsAsync(existingUser);
        _userRepoMock.Setup(r => r.Update(It.IsAny<User>()));

        var result = await _authCtrl.ResetPassword(new ResetPasswordDto
        {
            Token = token,
            NewPassword = "NewPass123"
        });

        Assert.IsType<OkObjectResult>(result);
        var usedToken = _db.PasswordResetTokens.First(t => t.Token == token);
        Assert.True(usedToken.IsUsed);
    }

    [Fact]
    public async Task ResetPassword_InvalidToken_ReturnsBadRequest()
    {
        var result = await _authCtrl.ResetPassword(new ResetPasswordDto
        {
            Token = "invalid_token_hex",
            NewPassword = "NewPass123"
        });

        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public async Task ResetPassword_ShortPassword_ReturnsBadRequest()
    {
        var result = await _authCtrl.ResetPassword(new ResetPasswordDto
        {
            Token = "some_token",
            NewPassword = "short"
        });

        Assert.IsType<BadRequestObjectResult>(result);
    }

    // ========== Orders 分页测试 ==========

    [Fact]
    public async Task GetUserOrders_DefaultPage_Returns10Items()
    {
        var el = ToJson(await _ordersCtrl.GetUserOrders(TestUserId, page: 1, pageSize: 10));
        Assert.Equal(15, el.GetProperty("total").GetInt32());
        Assert.Equal(10, el.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public async Task GetUserOrders_Page2_Returns5Items()
    {
        var el = ToJson(await _ordersCtrl.GetUserOrders(TestUserId, page: 2, pageSize: 10));
        Assert.Equal(5, el.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public async Task GetUserOrders_StatusFilter_ReturnsFiltered()
    {
        var el = ToJson(await _ordersCtrl.GetUserOrders(TestUserId, page: 1, pageSize: 50, status: "Pending"));
        Assert.Equal(5, el.GetProperty("total").GetInt32());
    }

    [Fact]
    public async Task GetUserOrders_NoResults_ReturnsEmpty()
    {
        var el = ToJson(await _ordersCtrl.GetUserOrders(TestUserId, page: 1, pageSize: 10, status: "Refunded"));
        Assert.Equal(0, el.GetProperty("total").GetInt32());
    }

    // ========== OrderDetail 测试 ==========

    [Fact]
    public async Task GetUserOrderDetail_ValidOrder_ReturnsFullDetail()
    {
        var el = ToJson(await _ordersCtrl.GetUserOrderDetail(TestUserId, 1));
        Assert.Equal(1, el.GetProperty("OrderId").GetInt32());
        Assert.Equal("ORD202401010001", el.GetProperty("OrderNo").GetString());
        Assert.True(el.GetProperty("details").GetArrayLength() > 0);
    }

    [Fact]
    public async Task GetUserOrderDetail_WrongUser_ReturnsNotFound()
    {
        var result = await _ordersCtrl.GetUserOrderDetail(999, 1);
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public async Task GetUserOrderDetail_NonexistentOrder_ReturnsNotFound()
    {
        var result = await _ordersCtrl.GetUserOrderDetail(TestUserId, 9999);
        Assert.IsType<NotFoundObjectResult>(result);
    }

    // ========== DeleteOrder 测试 ==========

    [Fact]
    public async Task DeleteOrder_CompletedOrder_ReturnsOk()
    {
        var completedOrder = _db.Orders.First(o => o.Status == "Completed");
        _orderRepoMock.Setup(r => r.GetByIdAsync(completedOrder.OrderId, It.IsAny<CancellationToken>()))
            .ReturnsAsync(completedOrder);
        _orderRepoMock.Setup(r => r.Update(It.IsAny<Order>()));
        _orderRepoMock.Setup(r => r.SaveChangesAsync(It.IsAny<CancellationToken>())).ReturnsAsync(1);

        var result = await _ordersCtrl.DeleteOrder(completedOrder.OrderId, TestUserId);
        Assert.IsType<OkObjectResult>(result);
    }

    [Fact]
    public async Task DeleteOrder_PendingOrder_ReturnsBadRequest()
    {
        var pendingOrder = _db.Orders.First(o => o.Status == "Pending");
        _orderRepoMock.Setup(r => r.GetByIdAsync(pendingOrder.OrderId, It.IsAny<CancellationToken>()))
            .ReturnsAsync(pendingOrder);

        var result = await _ordersCtrl.DeleteOrder(pendingOrder.OrderId, TestUserId);
        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public async Task DeleteOrder_WrongUser_ReturnsForbid()
    {
        var order = _db.Orders.First();
        _orderRepoMock.Setup(r => r.GetByIdAsync(order.OrderId, It.IsAny<CancellationToken>()))
            .ReturnsAsync(order);

        var result = await _ordersCtrl.DeleteOrder(order.OrderId, 999);
        Assert.IsType<ForbidResult>(result);
    }

    // ========== Login 测试 (AuthController.Login) ==========

    [Fact]
    public async Task Login_ValidCredentials_ReturnsOkWithUser()
    {
        var user = _db.Users.First();
        _userRepoMock.Setup(r => r.AuthenticateAsync("testuser", "correctpassword", It.IsAny<CancellationToken>()))
            .ReturnsAsync(user);

        var el = ToJson(await _authCtrl.Login(new LoginRequest
        {
            Username = "testuser",
            Password = "correctpassword"
        }));
        Assert.Equal(TestUserId, el.GetProperty("userId").GetInt32());
        Assert.Equal("testuser", el.GetProperty("username").GetString());
    }

    [Fact]
    public async Task Login_InvalidCredentials_ReturnsUnauthorized()
    {
        _userRepoMock.Setup(r => r.AuthenticateAsync("testuser", "wrongpassword", It.IsAny<CancellationToken>()))
            .ReturnsAsync((User?)null);

        var result = await _authCtrl.Login(new LoginRequest
        {
            Username = "testuser",
            Password = "wrongpassword"
        });

        Assert.IsType<UnauthorizedObjectResult>(result);
    }

    [Fact]
    public async Task Login_EmptyCredentials_ReturnsBadRequest()
    {
        var result = await _authCtrl.Login(new LoginRequest
        {
            Username = "",
            Password = ""
        });

        Assert.IsType<BadRequestObjectResult>(result);
    }
}
