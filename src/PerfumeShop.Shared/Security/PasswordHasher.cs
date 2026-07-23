using System.Security.Cryptography;
using System.Text;

namespace PerfumeShop.Shared.Security;

/// <summary>
/// V19 统一口令散列器 — 从 Login.razor 抽取，保持与登录端完全一致的算法。
/// 兼容 V18 password_utils.asp：全角/半角双向标准化 + V1/V2/V3 版本化散列 + 自动升级。
/// 供登录验证、管理员创建、密码重置共用，消除 SysAdmin 页面裸 SHA256 的不一致缺陷。
/// </summary>
public static class PasswordHasher
{
    // 对齐 V18 config.asp PASSWORD_PEPPER
    private const string PasswordPepper = "P3rfum3Sh0p_S@lt_2026!";

    /// <summary>验证结果：是否成功，以及（若命中旧版本 V1/V2）应升级到的 V3 散列。</summary>
    public readonly record struct VerifyOutcome(bool Success, string? UpgradedHash)
    {
        public static VerifyOutcome Ok() => new(true, null);
        public static VerifyOutcome OkUpgrade(string hash) => new(true, hash);
        public static VerifyOutcome Fail() => new(false, null);
    }

    /// <summary>生成最新版本(V3)散列，供创建/重置口令使用。</summary>
    public static string Hash(string password) => HashPasswordV3(NormalizePasswordInput(password));

    /// <summary>验证口令；返回是否成功及（如命中旧版本）应升级到的 V3 散列。</summary>
    public static VerifyOutcome Verify(string password, string? storedHash)
    {
        if (string.IsNullOrEmpty(storedHash)) return VerifyOutcome.Fail();
        var normalized = NormalizePasswordInput(password);
        var expanded = ExpandToFullwidth(password);

        if (storedHash.StartsWith("V3$")) return VerifyV3Multi(password, normalized, expanded, storedHash);
        if (storedHash.StartsWith("V2_")) return VerifyV2Multi(password, normalized, expanded, storedHash);
        return VerifyV1Multi(password, normalized, expanded, storedHash);
    }

    /// <summary>生成随机临时密码（用于管理员密码重置默认值）。</summary>
    public static string GenerateTempPassword(int length = 10)
    {
        const string chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
        var rng = Random.Shared;
        return new string(Enumerable.Range(0, Math.Max(6, length)).Select(_ => chars[rng.Next(chars.Length)]).ToArray());
    }

    // ===== V17.2 全角/半角密码标准化 (从 V18 password_utils.asp 移植) =====

    private static string NormalizePasswordInput(string pwd)
    {
        if (string.IsNullOrEmpty(pwd)) return "";
        var result = new StringBuilder();
        foreach (char c in pwd)
        {
            int code = c;
            if (code >= 0xFF01 && code <= 0xFF5E) result.Append((char)(code - 0xFEE0));
            else if (code == 0x3000) result.Append((char)0x20);
            else result.Append(c);
        }
        return result.ToString();
    }

    private static string ExpandToFullwidth(string pwd)
    {
        if (string.IsNullOrEmpty(pwd)) return "";
        var result = new StringBuilder();
        foreach (char c in pwd)
        {
            int code = c;
            if ((code >= 0x21 && code <= 0x2F) || (code >= 0x3A && code <= 0x40) ||
                (code >= 0x5B && code <= 0x60) || (code >= 0x7B && code <= 0x7E))
                result.Append((char)(code + 0xFEE0));
            else if (code == 0x20) result.Append((char)0x3000);
            else result.Append(c);
        }
        return result.ToString();
    }

    // ===== V18 完整密码验证 (全角/半角双向 + V1/V2/V3 + 自动升级) =====

    private static VerifyOutcome VerifyV3Multi(string orig, string norm, string exp, string storedHash)
    {
        if (VerifyV3(orig, storedHash)) return VerifyOutcome.Ok();
        if (norm != orig && VerifyV3(norm, storedHash)) return VerifyOutcome.OkUpgrade(HashPasswordV3(norm));
        if (exp != orig && exp != norm && VerifyV3(exp, storedHash)) return VerifyOutcome.OkUpgrade(HashPasswordV3(norm.Length > 0 ? norm : exp));
        return VerifyOutcome.Fail();
    }

    private static VerifyOutcome VerifyV2Multi(string orig, string norm, string exp, string storedHash)
    {
        if (HashV2(orig) == storedHash) return VerifyOutcome.OkUpgrade(HashPasswordV3(orig));
        if (norm != orig && HashV2(norm) == storedHash) return VerifyOutcome.OkUpgrade(HashPasswordV3(norm));
        if (exp != orig && exp != norm && HashV2(exp) == storedHash) return VerifyOutcome.OkUpgrade(HashPasswordV3(norm.Length > 0 ? norm : exp));
        return VerifyOutcome.Fail();
    }

    private static VerifyOutcome VerifyV1Multi(string orig, string norm, string exp, string storedHash)
    {
        if (HashV1(orig) == storedHash) return VerifyOutcome.OkUpgrade(HashPasswordV3(orig));
        if (norm != orig && HashV1(norm) == storedHash) return VerifyOutcome.OkUpgrade(HashPasswordV3(norm));
        if (exp != orig && exp != norm && HashV1(exp) == storedHash) return VerifyOutcome.OkUpgrade(HashPasswordV3(norm.Length > 0 ? norm : exp));
        return VerifyOutcome.Fail();
    }

    private static string HashPasswordV3(string password)
    {
        var salt = GenerateSaltV3(32);
        var hash = HashV3FromSalt(password, salt);
        return $"V3${salt}${hash}";
    }

    private static string GenerateSaltV3(int length)
    {
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        var random = Random.Shared;
        return new string(Enumerable.Range(0, length).Select(_ => chars[random.Next(chars.Length)]).ToArray());
    }

    private static bool VerifyV3(string password, string storedHash)
    {
        var parts = storedHash.Split('$');
        if (parts.Length < 3) return false;
        return string.Equals(HashV3FromSalt(password, parts[1]), parts[2], StringComparison.OrdinalIgnoreCase);
    }

    private static string HashV3FromSalt(string password, string salt)
    {
        using var sha = SHA256.Create();
        var combined = Encoding.ASCII.GetBytes(salt + password + PasswordPepper);
        var hashBytes = sha.ComputeHash(combined);
        var hashHex = BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
        for (int i = 0; i < 999; i++)
        {
            var strInput = hashHex + salt;
            hashBytes = sha.ComputeHash(Encoding.ASCII.GetBytes(strInput));
            hashHex = BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
        }
        return hashHex;
    }

    private static string HashV2(string password)
    {
        var salted = password + PasswordPepper;
        for (int i = 0; i < 10; i++) salted = InternalHash(salted);
        return "V2_" + salted;
    }

    private static string InternalHash(string input)
    {
        var sb = new StringBuilder(); int prev = 0;
        foreach (char c in input) { int cc = (int)c; prev = (prev + cc) % 256; sb.Append(((cc * 7 + prev) % 256).ToString("x2")); }
        return sb.ToString();
    }

    private static string HashV1(string password)
    {
        var sb = new StringBuilder();
        foreach (char c in password) sb.Append(((int)c).ToString("x2"));
        return sb.ToString();
    }
}
