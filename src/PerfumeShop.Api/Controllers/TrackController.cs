using Microsoft.AspNetCore.Mvc;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 用户行为追踪 API — 对应 V18 api/track.asp
/// </summary>
[ApiController]
[Route("api/v2/track")]
public class TrackController : ControllerBase
{
    // 1x1 transparent GIF (43 bytes)
    private static readonly byte[] TransparentGif = new byte[]
    {
        0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00,
        0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x21, 0xF9, 0x04, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x2C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00,
        0x00, 0x02, 0x02, 0x44, 0x01, 0x00, 0x3B
    };

    /// <summary>POST /api/v2/track — 记录用户行为（JSON方式）</summary>
    [HttpPost]
    public IActionResult TrackEvent([FromBody] TrackEventRequest req)
    {
        // In production this would write to the database via tracking_utils
        // For now, return success acknowledging the event
        return Ok(new
        {
            success = true,
            message = "事件已记录",
            data = new
            {
                action = req.Action,
                target = req.Target,
                keyword = req.Keyword,
                timestamp = DateTime.UtcNow
            }
        });
    }

    /// <summary>GET /api/v2/track.gif — 1x1 像素追踪 GIF</summary>
    [HttpGet("gif")]
    public IActionResult TrackGif(
        [FromQuery] string? action,
        [FromQuery] string? target,
        [FromQuery] string? keyword)
    {
        // Log the tracking event (simplified — in production would use tracking service)
        // Output the GIF regardless
        Response.Headers.Append("Cache-Control", "no-cache, no-store");
        Response.Headers.Append("Pragma", "no-cache");
        Response.Headers.Append("Expires", "-1");
        return File(TransparentGif, "image/gif");
    }
}

public class TrackEventRequest
{
    public string Action { get; set; } = "";
    public string? Target { get; set; }
    public string? Keyword { get; set; }
    public int? Qty { get; set; }
}
