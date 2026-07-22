using Microsoft.AspNetCore.Mvc;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// Cookie 同意 API — 对应 V18 api/cookie_consent.asp
/// </summary>
[ApiController]
[Route("api/v2/cookie-consent")]
public class CookieConsentController : ControllerBase
{
    /// <summary>POST /api/v2/cookie-consent — 记录 Cookie 同意偏好</summary>
    [HttpPost]
    public IActionResult RecordConsent([FromBody] CookieConsentRequest req)
    {
        var consentLevel = string.IsNullOrWhiteSpace(req.Consent) ? "essential" : req.Consent;

        // Validate consent value
        if (consentLevel != "all" && consentLevel != "essential")
            consentLevel = "essential";

        // Set consent cookie
        Response.Cookies.Append("cookie_consent", consentLevel, new CookieOptions
        {
            Expires = DateTime.UtcNow.AddYears(1),
            HttpOnly = false,
            Secure = true,
            SameSite = SameSiteMode.Lax
        });

        return Ok(new
        {
            success = true,
            message = "Cookie 偏好已记录",
            data = new { consent = consentLevel }
        });
    }
}

public class CookieConsentRequest
{
    public string Consent { get; set; } = "essential";
}
