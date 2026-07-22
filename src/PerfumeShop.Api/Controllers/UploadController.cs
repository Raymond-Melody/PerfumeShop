using Microsoft.AspNetCore.Mvc;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 文件上传 API — 对应 V18 includes/upload_utils.asp
/// </summary>
[ApiController]
[Route("api/v2/upload")]
public class UploadController : ControllerBase
{
    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg"
    };

    private const long MaxFileSize = 5 * 1024 * 1024; // 5MB

    private readonly IWebHostEnvironment _env;

    public UploadController(IWebHostEnvironment env)
    {
        _env = env;
    }

    /// <summary>POST /api/v2/upload — 上传文件</summary>
    [HttpPost]
    public async Task<IActionResult> Upload(IFormFile? file)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { success = false, message = "请选择要上传的文件" });

        if (file.Length > MaxFileSize)
            return BadRequest(new { success = false, message = $"文件大小超过限制（最大 {MaxFileSize / 1024 / 1024}MB）" });

        var ext = Path.GetExtension(file.FileName);
        if (!AllowedExtensions.Contains(ext))
            return BadRequest(new { success = false, message = $"不允许的文件类型: {ext}" });

        // Validate magic bytes
        if (!await ValidateImageMagicBytes(file))
            return BadRequest(new { success = false, message = "文件内容不是有效的图片格式" });

        var uploadDir = Path.Combine(_env.ContentRootPath, "wwwroot", "uploads");
        if (!Directory.Exists(uploadDir))
            Directory.CreateDirectory(uploadDir);

        var uniqueName = GenerateUniqueFileName(ext);
        var filePath = Path.Combine(uploadDir, uniqueName);

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        return Ok(new
        {
            success = true,
            data = new
            {
                fileName = uniqueName,
                originalName = file.FileName,
                size = file.Length,
                url = $"/uploads/{uniqueName}"
            }
        });
    }

    /// <summary>GET /api/v2/upload/{id} — 获取上传文件信息</summary>
    [HttpGet("{id}")]
    public IActionResult GetUploadInfo(string id)
    {
        var uploadDir = Path.Combine(_env.ContentRootPath, "wwwroot", "uploads");
        var filePath = Path.Combine(uploadDir, id);

        if (!System.IO.File.Exists(filePath))
            return NotFound(new { success = false, message = "文件不存在" });

        var info = new FileInfo(filePath);
        return Ok(new
        {
            success = true,
            data = new
            {
                fileName = id,
                size = info.Length,
                lastModified = info.LastWriteTimeUtc,
                url = $"/uploads/{id}"
            }
        });
    }

    private static string GenerateUniqueFileName(string ext)
    {
        var now = DateTime.Now;
        var random = Random.Shared.Next(10000, 99999);
        return $"{now:yyyyMMddHHmmss}_{random}{ext}";
    }

    private static async Task<bool> ValidateImageMagicBytes(IFormFile file)
    {
        using var stream = file.OpenReadStream();
        var header = new byte[12];
        var bytesRead = await stream.ReadAsync(header, 0, 12);
        if (bytesRead < 3) return false;

        // JPEG: FF D8 FF
        if (bytesRead >= 3 && header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF) return true;
        // PNG: 89 50 4E 47
        if (bytesRead >= 4 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47) return true;
        // GIF: 47 49 46 38
        if (bytesRead >= 4 && header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x38) return true;
        // WebP: RIFF....WEBP
        if (bytesRead >= 12 && header[0] == 0x52 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x46
            && header[8] == 0x57 && header[9] == 0x45 && header[10] == 0x42 && header[11] == 0x50) return true;
        // SVG (text-based, skip magic check)
        var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (ext == ".svg") return true;

        return false;
    }
}
