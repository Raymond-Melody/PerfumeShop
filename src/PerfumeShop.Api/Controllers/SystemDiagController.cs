using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 系统诊断 API — 对应 V18 api/system_diag.asp
/// </summary>
[ApiController]
[Route("api/v2/diag")]
public class SystemDiagController : ControllerBase
{
    private readonly IWebHostEnvironment _env;

    public SystemDiagController(IWebHostEnvironment env)
    {
        _env = env;
    }

    /// <summary>GET /api/v2/diag/system — 系统配置诊断</summary>
    [HttpGet("system")]
    public IActionResult GetSystemDiag()
    {
        var process = Process.GetCurrentProcess();

        return Ok(new
        {
            status = "ok",
            timestamp = DateTime.UtcNow,
            version = "V19.0",
            runtime = new
            {
                framework = RuntimeInformation.FrameworkDescription,
                os = RuntimeInformation.OSDescription,
                architecture = RuntimeInformation.OSArchitecture.ToString(),
                processorCount = Environment.ProcessorCount,
                workingSetMB = Math.Round(process.WorkingSet64 / 1048576.0, 2),
                uptimeSeconds = (DateTime.UtcNow - process.StartTime.ToUniversalTime()).TotalSeconds
            },
            environment = new
            {
                environmentName = _env.EnvironmentName,
                contentRootPath = _env.ContentRootPath,
                webRootPath = _env.WebRootPath
            },
            features = new
            {
                MSOLEDBSQL = true,
                DAL_ENABLED = true,
                I18N = true,
                PASSWORD_V3 = true,
                STRUCTURED_LOGGING = true,
                API_V1 = true,
                CACHE_MANAGER = true,
                SSE_NOTIFICATIONS = true,
                AI_SEARCH = true
            },
            session = new
            {
                timeoutMinutes = 20
            }
        });
    }
}
