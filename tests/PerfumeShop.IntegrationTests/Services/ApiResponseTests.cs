using PerfumeShop.Shared.ApiContracts;
using System.Text.Json;

namespace PerfumeShop.IntegrationTests.Services;

/// <summary>
/// ApiResponse 测试 — 5 用例
/// 覆盖 Success/Error/Paginated 序列化
/// </summary>
public class ApiResponseTests
{
    [Fact]
    public void Ok_ReturnsSuccessResponse()
    {
        var response = ApiResponse<string>.Ok("hello", "操作成功");

        Assert.True(response.Success);
        Assert.Equal("hello", response.Data);
        Assert.Equal("操作成功", response.Message);
        Assert.Null(response.Error);
        Assert.Null(response.Pagination);
    }

    [Fact]
    public void Fail_ReturnsErrorResponse()
    {
        var response = ApiResponse<string>.Fail("param_invalid", "参数错误");

        Assert.False(response.Success);
        Assert.Null(response.Data);
        Assert.Equal("param_invalid", response.Error);
        Assert.Equal("参数错误", response.Message);
    }

    [Fact]
    public void NotFound_ReturnsNotFoundResponse()
    {
        var response = ApiResponse<object>.NotFound();

        Assert.False(response.Success);
        Assert.Equal("not_found", response.Error);
        Assert.Equal("请求的资源不存在", response.Message);
    }

    [Fact]
    public void Paginated_IncludesPaginationInfo()
    {
        var items = new List<string> { "a", "b", "c" };
        var response = ApiResponse<List<string>>.Paginated(items, total: 25, page: 2, size: 10);

        Assert.True(response.Success);
        Assert.Equal(items, response.Data);
        Assert.NotNull(response.Pagination);
        Assert.Equal(2, response.Pagination.Page);
        Assert.Equal(10, response.Pagination.Size);
        Assert.Equal(25, response.Pagination.Total);
        Assert.Equal(3, response.Pagination.TotalPages);
        Assert.True(response.Pagination.HasPrev);
        Assert.True(response.Pagination.HasNext);
    }

    [Fact]
    public void Paginated_SerializesToJson()
    {
        var response = ApiResponse<string>.Paginated("item", total: 10, page: 1, size: 5);
        var json = JsonSerializer.Serialize(response, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        Assert.Contains("\"success\":true", json);
        Assert.Contains("\"data\":\"item\"", json);
        Assert.Contains("\"page\":1", json);
        Assert.Contains("\"total\":10", json);
        Assert.Contains("\"totalPages\":2", json);
        Assert.Contains("\"hasPrev\":false", json);
        Assert.Contains("\"hasNext\":true", json);
    }

    [Fact]
    public void NonGeneric_Ok_ReturnsObjectResponse()
    {
        var response = ApiResponse.Ok(message: "done");
        Assert.True(response.Success);
        Assert.Equal("done", response.Message);
    }

    [Fact]
    public void PaginationInfo_EdgeCases()
    {
        // 首页无 prev
        var p1 = new PaginationInfo(1, 10, 30);
        Assert.False(p1.HasPrev);
        Assert.True(p1.HasNext);
        Assert.Equal(3, p1.TotalPages);

        // 末页无 next
        var pLast = new PaginationInfo(3, 10, 30);
        Assert.True(pLast.HasPrev);
        Assert.False(pLast.HasNext);

        // size = 0
        var pZero = new PaginationInfo(1, 0, 10);
        Assert.Equal(0, pZero.TotalPages);
        Assert.False(pZero.HasNext);
    }

    [Fact]
    public void ApiErrorCode_MatchesV18Constants()
    {
        Assert.Equal(0, ApiErrorCode.Success);
        Assert.Equal(1001, ApiErrorCode.AuthRequired);
        Assert.Equal(2001, ApiErrorCode.ParamMissing);
        Assert.Equal(3001, ApiErrorCode.NotFound);
        Assert.Equal(4001, ApiErrorCode.DbError);
        Assert.Equal(6001, ApiErrorCode.ServerError);
    }
}
