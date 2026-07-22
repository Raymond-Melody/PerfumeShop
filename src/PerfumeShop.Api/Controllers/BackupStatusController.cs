using Microsoft.AspNetCore.Mvc;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 备份状态 API — 对应 V18 api/backup_status.asp
/// </summary>
[ApiController]
[Route("api/v2/backup")]
public class BackupStatusController : ControllerBase
{
    private readonly IWebHostEnvironment _env;

    public BackupStatusController(IWebHostEnvironment env)
    {
        _env = env;
    }

    /// <summary>GET /api/v2/backup/status — 获取备份系统状态</summary>
    [HttpGet("status")]
    public IActionResult GetStatus()
    {
        var backupPath = Path.Combine(_env.ContentRootPath, "..", "..", "database", "backups");

        string lastBackupName = "";
        long lastBackupSize = 0;
        DateTime? lastBackupTime = null;
        bool fileVerified = false;
        int totalBackups = 0;
        int recentBackups = 0;

        if (Directory.Exists(backupPath))
        {
            var bakFiles = Directory.GetFiles(backupPath, "*.bak")
                .Select(f => new FileInfo(f))
                .ToList();

            totalBackups = bakFiles.Count;

            var latest = bakFiles.OrderByDescending(f => f.LastWriteTime).FirstOrDefault();
            if (latest != null)
            {
                lastBackupName = latest.Name;
                lastBackupSize = latest.Length;
                lastBackupTime = latest.LastWriteTime;
                fileVerified = latest.Length > 512;
            }

            recentBackups = bakFiles.Count(f => (DateTime.Now - f.LastWriteTime).TotalDays <= 30);
        }

        return Ok(new
        {
            status = "ok",
            lastBackup = new
            {
                fileName = lastBackupName,
                sizeBytes = lastBackupSize,
                sizeMB = Math.Round(lastBackupSize / 1048576.0, 2),
                time = lastBackupTime?.ToString("o") ?? "",
                verified = fileVerified
            },
            totals = new
            {
                totalBackups,
                recent30Days = recentBackups
            },
            schedule = new
            {
                frequency = "daily",
                time = "02:00",
                nextRun = DateTime.Now.AddDays(1).Date.AddHours(2).ToString("o")
            },
            generatedAt = DateTime.UtcNow.ToString("o")
        });
    }
}
