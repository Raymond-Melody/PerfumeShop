using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 用户认证 API
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class AuthController : ControllerBase
{
    private readonly IUserRepository _userRepo;

    public AuthController(IUserRepository userRepo)
    {
        _userRepo = userRepo;
    }

    /// <summary>用户登录</summary>
    [HttpPost("login")]
    public async Task<IActionResult> Login([FromBody] LoginRequest request)
    {
        if (string.IsNullOrEmpty(request.Username) || string.IsNullOrEmpty(request.Password))
            return BadRequest(new { message = "用户名和密码不能为空" });

        var user = await _userRepo.AuthenticateAsync(request.Username, request.Password);
        if (user == null)
            return Unauthorized(new { message = "用户名或密码错误" });

        return Ok(new
        {
            userId = user.UserId,
            username = user.Username,
            email = user.Email,
            fullName = user.FullName,
            tier = user.CustomerTier,
            points = user.Points
        });
    }

    /// <summary>用户注册</summary>
    [HttpPost("register")]
    public async Task<IActionResult> Register([FromBody] RegisterRequest request)
    {
        if (string.IsNullOrEmpty(request.Username) || string.IsNullOrEmpty(request.Password) ||
            string.IsNullOrEmpty(request.Email))
            return BadRequest(new { message = "用户名、密码和邮箱不能为空" });

        if (await _userRepo.UsernameExistsAsync(request.Username))
            return Conflict(new { message = "用户名已存在" });

        if (await _userRepo.EmailExistsAsync(request.Email))
            return Conflict(new { message = "邮箱已被注册" });

        var user = new Data.Models.User
        {
            Username = request.Username,
            Password = request.Password, // TODO: 生产环境应使用哈希
            Email = request.Email,
            FullName = request.FullName,
            IsActive = true,
            CustomerTier = "bronze",
            Points = 0,
            CreatedAt = DateTime.Now
        };

        await _userRepo.AddAsync(user);
        await _userRepo.SaveChangesAsync();

        return Ok(new { userId = user.UserId, username = user.Username, message = "注册成功" });
    }

    /// <summary>获取当前用户信息</summary>
    [HttpGet("profile/{userId}")]
    public async Task<IActionResult> GetProfile(int userId)
    {
        var user = await _userRepo.GetByIdAsync(userId);
        if (user == null)
            return NotFound(new { message = "用户不存在" });

        return Ok(new
        {
            userId = user.UserId,
            username = user.Username,
            email = user.Email,
            fullName = user.FullName,
            phone = user.Phone,
            address = user.Address,
            city = user.City,
            tier = user.CustomerTier,
            points = user.Points,
            totalSpent = user.TotalSpent,
            orderCount = user.OrderCount
        });
    }
}

public class LoginRequest
{
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
}

public class RegisterRequest
{
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string Email { get; set; } = "";
    public string? FullName { get; set; }
}
