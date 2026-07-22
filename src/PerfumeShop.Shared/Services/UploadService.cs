using Microsoft.AspNetCore.Http;

namespace PerfumeShop.Shared.Services;

// ============================================
// V19 UploadService — 对齐 V18 includes/upload_utils.asp + api/upload.asp
// 功能：multipart/form-data 解析、MIME 类型白名单校验、大小限制、唯一文件名生成、
//       本地存储或 CDN 上传（抽象 IStorageProvider 接口）
// ============================================

/// <summary>允许的文件类型白名单（对齐 V18 ALLOWED_IMAGE_EXTENSIONS）</summary>
[Flags]
public enum AllowedFileTypes
{
    None = 0,
    Images = 1,        // jpg, jpeg, png, gif, webp, svg
    Documents = 2,     // pdf, doc, docx, xls, xlsx
    All = Images | Documents
}

/// <summary>上传结果（对齐 V18 SaveUploadedFile 返回值）</summary>
public class UploadResult
{
    public string Url { get; set; } = "";
    public string FileName { get; set; } = "";
    public long Size { get; set; }
    public string ContentType { get; set; } = "";
    public bool Success { get; set; }
    public string? Error { get; set; }
}

/// <summary>
/// 存储提供者抽象接口 — 对齐 V18 SaveUploadedFile / EnsureUploadDir
/// </summary>
public interface IStorageProvider
{
    /// <summary>保存文件并返回可访问的 URL</summary>
    Task<string> SaveAsync(byte[] data, string fileName, string folder, string contentType);
}

/// <summary>
/// 本地文件存储实现 — 对齐 V18 SaveUploadedFile + EnsureUploadDir + CreateFolderRecursive
/// </summary>
public class LocalStorageProvider : IStorageProvider
{
    private readonly string _basePath;
    private readonly string _baseUrl;

    public LocalStorageProvider(string basePath, string baseUrl = "/uploads")
    {
        _basePath = basePath;
        _baseUrl = baseUrl.TrimEnd('/');
    }

    public async Task<string> SaveAsync(byte[] data, string fileName, string folder, string contentType)
    {
        // 对齐 V18: EnsureUploadDir + CreateFolderRecursive
        var dir = Path.Combine(_basePath, folder);
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        var filePath = Path.Combine(dir, fileName);
        await File.WriteAllBytesAsync(filePath, data);

        return $"{_baseUrl}/{folder}/{fileName}";
    }
}

/// <summary>
/// 上传服务接口
/// </summary>
public interface IUploadService
{
    /// <summary>
    /// 上传文件 — 对齐 V18: MultipartParser.Parse + IsValidImageType + GenerateUploadFileName + SaveUploadedFile
    /// </summary>
    Task<UploadResult> UploadAsync(IFormFile file, string folder, AllowedFileTypes types = AllowedFileTypes.Images);
}

/// <summary>
/// 上传服务实现
/// 对齐 V18 函数映射：
///   - MultipartParser.Parse        → IFormFile 自动解析
///   - IsValidImageType             → ValidateFileType
///   - IsValidImageMagicBytes       → ValidateMagicBytes
///   - GenerateUploadFileName       → GenerateUniqueFileName
///   - SaveUploadedFile             → IStorageProvider.SaveAsync
///   - SanitizeFileName             → SanitizeFileName
///   - EnsureUploadDir              → LocalStorageProvider 内部处理
/// </summary>
public class UploadService : IUploadService
{
    private readonly IStorageProvider _storage;
    private readonly long _maxSizeBytes;

    // 对齐 V18: ALLOWED_IMAGE_EXTENSIONS = ".jpg,.jpeg,.png,.gif,.webp,.svg"
    private static readonly Dictionary<string, string> ImageMimeTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        [".jpg"] = "image/jpeg",
        [".jpeg"] = "image/jpeg",
        [".png"] = "image/png",
        [".gif"] = "image/gif",
        [".webp"] = "image/webp",
        [".svg"] = "image/svg+xml",
    };

    private static readonly Dictionary<string, string> DocumentMimeTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        [".pdf"] = "application/pdf",
        [".doc"] = "application/msword",
        [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        [".xls"] = "application/vnd.ms-excel",
        [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    };

    /// <param name="storage">存储提供者</param>
    /// <param name="maxSizeBytes">最大文件大小，默认 10MB（对齐 V18 MAX_UPLOAD_SIZE）</param>
    public UploadService(IStorageProvider storage, long maxSizeBytes = 10 * 1024 * 1024)
    {
        _storage = storage;
        _maxSizeBytes = maxSizeBytes;
    }

    public async Task<UploadResult> UploadAsync(IFormFile file, string folder, AllowedFileTypes types = AllowedFileTypes.Images)
    {
        // 1. 基本校验
        if (file == null || file.Length == 0)
            return new UploadResult { Error = "文件为空" };

        // 2. 大小校验 — 对齐 V18: MAX_UPLOAD_SIZE 检查
        if (file.Length > _maxSizeBytes)
            return new UploadResult { Error = $"文件大小超过限制（最大 {_maxSizeBytes / 1024 / 1024}MB）" };

        // 3. 文件名安全清理 — 对齐 V18: SanitizeFileName
        var originalName = SanitizeFileName(Path.GetFileName(file.FileName));
        var ext = Path.GetExtension(originalName).ToLowerInvariant();

        if (string.IsNullOrEmpty(ext))
            return new UploadResult { Error = "无法识别文件扩展名" };

        // 4. MIME 类型白名单校验 — 对齐 V18: IsValidImageType
        var allowedTypes = GetAllowedMimeTypes(types);
        if (!allowedTypes.ContainsKey(ext))
            return new UploadResult { Error = $"不允许的文件类型: {ext}" };

        // 5. 读取文件数据
        using var ms = new MemoryStream();
        await file.CopyToAsync(ms);
        var data = ms.ToArray();

        // 6. 魔数验证 — 对齐 V18: IsValidImageMagicBytes
        if (types.HasFlag(AllowedFileTypes.Images) && IsImageExtension(ext))
        {
            if (!ValidateMagicBytes(data, ext))
                return new UploadResult { Error = "文件内容与扩展名不匹配（魔数验证失败）" };
        }

        // 7. 生成唯一文件名 — 对齐 V18: GenerateUploadFileName
        var uniqueName = GenerateUniqueFileName(originalName);

        // 8. 保存 — 对齐 V18: SaveUploadedFile
        var url = await _storage.SaveAsync(data, uniqueName, folder, file.ContentType);

        return new UploadResult
        {
            Url = url,
            FileName = uniqueName,
            Size = data.Length,
            ContentType = allowedTypes[ext],
            Success = true
        };
    }

    // 对齐 V18: SanitizeFileName
    public static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var result = string.Concat(name.Where(c => !invalid.Contains(c)));
        // 额外清理 — 对齐 V18 防止目录遍历
        result = result.Replace("..", "").Replace("/", "").Replace("\\", "");
        result = result.Replace(":", "").Replace("*", "").Replace("?", "");
        result = result.Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "");
        return result.Trim();
    }

    // 对齐 V18: GenerateUploadFileName (日期时间_随机数.扩展名)
    public static string GenerateUniqueFileName(string originalName)
    {
        var ext = Path.GetExtension(originalName);
        var now = DateTime.Now;
        var random = Random.Shared.Next(10000, 99999);
        return $"{now:yyyyMMddHHmmss}_{random}{ext}";
    }

    // 对齐 V18: IsValidImageMagicBytes
    public static bool ValidateMagicBytes(byte[] data, string ext)
    {
        if (data.Length < 4) return false;

        return ext switch
        {
            ".jpg" or ".jpeg" => data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF,
            ".png" => data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47,
            ".gif" => data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38,
            ".webp" => data.Length >= 12 && data[0] == 0x52 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x46
                       && data[8] == 0x57 && data[9] == 0x45 && data[10] == 0x42 && data[11] == 0x50,
            ".svg" => true, // SVG 是文本格式，不做魔数检查
            _ => true  // 非图片类型不做魔数检查
        };
    }

    private static bool IsImageExtension(string ext)
        => ImageMimeTypes.ContainsKey(ext);

    private static Dictionary<string, string> GetAllowedMimeTypes(AllowedFileTypes types)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (types.HasFlag(AllowedFileTypes.Images))
            foreach (var kv in ImageMimeTypes) result[kv.Key] = kv.Value;
        if (types.HasFlag(AllowedFileTypes.Documents))
            foreach (var kv in DocumentMimeTypes) result[kv.Key] = kv.Value;
        return result;
    }
}
