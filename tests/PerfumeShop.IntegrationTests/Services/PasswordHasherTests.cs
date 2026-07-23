using System.Linq;
using PerfumeShop.Shared.Security;
using Xunit;

namespace PerfumeShop.IntegrationTests.Services;

/// <summary>
/// M1 单元测试 — 统一口令散列器 PasswordHasher。
/// 校验 V3 生成/验证、错误口令拒绝、旧版本(V1)兼容+自动升级、全角/半角标准化、临时口令生成。
/// </summary>
public class PasswordHasherTests
{
    [Fact]
    public void Hash_Produces_V3_Format()
    {
        var hash = PasswordHasher.Hash("Secret@123");
        Assert.StartsWith("V3$", hash);
        Assert.Equal(3, hash.Split('$').Length); // V3 $ salt $ hash
    }

    [Fact]
    public void Hash_Then_Verify_Succeeds_Without_Upgrade()
    {
        var hash = PasswordHasher.Hash("Secret@123");
        var result = PasswordHasher.Verify("Secret@123", hash);
        Assert.True(result.Success);
        Assert.Null(result.UpgradedHash); // 已是最新 V3，无需升级
    }

    [Fact]
    public void Verify_WrongPassword_Fails()
    {
        var hash = PasswordHasher.Hash("Secret@123");
        var result = PasswordHasher.Verify("wrong-password", hash);
        Assert.False(result.Success);
    }

    [Fact]
    public void Verify_EmptyStoredHash_Fails()
    {
        Assert.False(PasswordHasher.Verify("anything", null).Success);
        Assert.False(PasswordHasher.Verify("anything", "").Success);
    }

    [Fact]
    public void Verify_Legacy_V1_Succeeds_And_Upgrades_To_V3()
    {
        // V1 = 每字符 (int)c 的两位十六进制拼接（对齐 Login.razor HashV1）
        const string pwd = "test1234";
        var v1 = string.Concat(pwd.Select(c => ((int)c).ToString("x2")));

        var result = PasswordHasher.Verify(pwd, v1);

        Assert.True(result.Success);
        Assert.NotNull(result.UpgradedHash);
        Assert.StartsWith("V3$", result.UpgradedHash!);

        // 升级后的 V3 散列应能继续验证同一口令
        Assert.True(PasswordHasher.Verify(pwd, result.UpgradedHash!).Success);
    }

    [Fact]
    public void Verify_FullWidth_Input_Matches_HalfWidth_Hash()
    {
        // 半角创建，全角输入应通过标准化匹配（对齐 V18 password_utils 全/半角双向）
        var hash = PasswordHasher.Hash("abc123");
        var result = PasswordHasher.Verify("ａｂｃ１２３", hash); // 全角
        Assert.True(result.Success);
    }

    [Fact]
    public void GenerateTempPassword_Respects_MinimumLength()
    {
        Assert.True(PasswordHasher.GenerateTempPassword(10).Length == 10);
        Assert.True(PasswordHasher.GenerateTempPassword(3).Length >= 6); // 下限 6
    }

    [Fact]
    public void Hash_Uses_Random_Salt_So_Two_Hashes_Differ()
    {
        var a = PasswordHasher.Hash("SamePassword!");
        var b = PasswordHasher.Hash("SamePassword!");
        Assert.NotEqual(a, b); // 随机盐 → 不同密文
        Assert.True(PasswordHasher.Verify("SamePassword!", a).Success);
        Assert.True(PasswordHasher.Verify("SamePassword!", b).Success);
    }
}
