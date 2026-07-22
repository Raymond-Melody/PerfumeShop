using Microsoft.AspNetCore.Mvc;
using System.Text.Json.Serialization;

namespace PerfumeShop.Shared.ApiContracts;

/// <summary>
/// V19 统一 API 响应 — 对应 V18 includes/api_response.asp
/// 泛型 ApiResponse&lt;T&gt;，字段：Success/Data/Error/Message/Pagination/Timestamp
/// </summary>
public class ApiResponse<T>
{
    public bool Success { get; set; }
    public T? Data { get; set; }
    public string? Error { get; set; }
    public string? Message { get; set; }
    public PaginationInfo? Pagination { get; set; }
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;

    /// <summary>请求 ID（对应 V18 API_GetRequestId）</summary>
    public string? RequestId { get; set; }

    // ── 静态工厂 ──────────────────────────────────────────

    public static ApiResponse<T> Ok(T data, string? message = null)
        => new() { Success = true, Data = data, Message = message ?? "success" };

    public static ApiResponse<T> Fail(string error, string? message = null)
        => new() { Success = false, Error = error, Message = message ?? error };

    public static ApiResponse<T> NotFound(string? message = null)
        => new() { Success = false, Error = "not_found", Message = message ?? "请求的资源不存在" };

    public static ApiResponse<T> Paginated(T data, int total, int page, int size, string? message = null)
        => new()
        {
            Success = true,
            Data = data,
            Message = message ?? "success",
            Pagination = new PaginationInfo(page, size, total)
        };
}

/// <summary>
/// 非泛型快捷入口（对应 V18 API_Success / API_Error / API_Message）
/// </summary>
public static class ApiResponse
{
    public static ApiResponse<object> Ok(object? data = null, string? message = null)
        => ApiResponse<object>.Ok(data ?? new { }, message);

    public static ApiResponse<object> Fail(string error, string? message = null)
        => ApiResponse<object>.Fail(error, message);

    public static ApiResponse<object> NotFound(string? message = null)
        => ApiResponse<object>.NotFound(message);

    public static ApiResponse<object> Message(bool success, string message)
        => success ? ApiResponse<object>.Ok(new { }, message) : ApiResponse<object>.Fail("business_error", message);
}

/// <summary>
/// 分页信息（对应 V18 API 分页响应）
/// </summary>
public class PaginationInfo
{
    public int Page { get; set; }
    public int Size { get; set; }
    public int Total { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public int TotalPages => Size > 0 ? (int)Math.Ceiling((double)Total / Size) : 0;

    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool HasPrev => Page > 1;

    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool HasNext => Page < TotalPages;

    public PaginationInfo() { }
    public PaginationInfo(int page, int size, int total) { Page = page; Size = size; Total = total; }
}

/// <summary>
/// V18 标准错误码（对应 includes/api_response.asp 常量）
/// </summary>
public static class ApiErrorCode
{
    public const int Success = 0;

    // 认证/授权 1000-1999
    public const int AuthRequired = 1001;
    public const int AuthExpired = 1002;
    public const int CsrfInvalid = 1003;
    public const int Forbidden = 1004;

    // 参数 2000-2999
    public const int ParamMissing = 2001;
    public const int ParamInvalid = 2002;
    public const int ParamType = 2003;

    // 业务 3000-3999
    public const int NotFound = 3001;
    public const int Duplicate = 3002;
    public const int LimitExceeded = 3003;
    public const int BusinessRule = 3004;

    // 数据库 4000-4999
    public const int DbError = 4001;
    public const int DbTimeout = 4002;
    public const int DbDeadlock = 4003;

    // 文件/上传 5000-5999
    public const int FileUpload = 5001;
    public const int FileType = 5002;
    public const int FileSize = 5003;

    // 服务端 6000-6999
    public const int ServerError = 6001;
    public const int Maintenance = 6002;
}

/// <summary>
/// Controller 扩展方法：将现有 JsonResult 转为 ApiResponse
/// </summary>
public static class ApiResponseExtensions
{
    /// <summary>将 JsonResult 的数据包装为 ApiResponse</summary>
    public static ActionResult ToApiResponse<T>(this T data, string? message = null)
    {
        return new OkObjectResult(ApiResponse<T>.Ok(data, message));
    }

    /// <summary>将分页数据包装为 ApiResponse</summary>
    public static ActionResult ToPaginatedApiResponse<T>(this T data, int total, int page, int size, string? message = null)
    {
        return new OkObjectResult(ApiResponse<T>.Paginated(data, total, page, size, message));
    }
}
