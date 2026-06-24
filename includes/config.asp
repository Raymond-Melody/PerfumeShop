<%
' ============================================
' PerfumeShop Configuration V14.6
' ============================================
Const SITE_NAME = "Custom Fragrance"
Const SITE_EMAIL = "contact@perfumeshop.com"
Const SITE_NOREPLY = "noreply@perfumeshop.com"
Const SITE_PHONE = "400-888-8888"

Function GetSiteURL()
    Dim proto
    If Request.ServerVariables("HTTPS") = "on" Then proto = "https://" Else proto = "http://"
    GetSiteURL = proto & Request.ServerVariables("HTTP_HOST")
End Function

Const SYS_VERSION = "V14.6"
Const SYS_VERSION_NAME = "PerfumeShop V14.6"
Const COOKIE_SECRET = "PerfumeShop_SecKey_2026_X9K3m"
Const COOKIE_SECRET_V10 = "PF_V10_SHA256_Salt_7kM2xP9qR4vN8wL3jH6fD1sA5gK0"
Const PASSWORD_PEPPER = "P3rfum3Sh0p_S@lt_2026!"
Const REFERRAL_SECRET = "PF_V14_Referral_Salt_9mK4xR7vN2wL8jH3fD1sA5"
Const BACKUP_RETENTION_DAYS = 30
Const PAGE_SIZE = 12
Const FREE_SHIPPING_AMOUNT = 299
Const SHIPPING_FEE = 15
Const IMAGE_PATH = "/images/"
Const DEFAULT_PRODUCT_IMAGE = "/images/default-product.svg"
Const DEFAULT_AVATAR = "/images/default-avatar.svg"
Const UPLOAD_PATH_PRODUCTS = "/images/products/"
Const UPLOAD_PATH_NOTES = "/images/notes/"
Const UPLOAD_PATH_BOTTLES = "/images/bottles/"
Const UPLOAD_PATH_AVATARS = "/images/avatars/"
Const UPLOAD_PATH_DEFAULT = "/images/uploads/"
Session.Timeout = 60
Const ORDER_PREFIX = "PF"

Sub GenerateCSPNonce()
    Dim nonce, i, charType, charCode
    Randomize
    nonce = ""
    For i = 1 To 32
        charType = Int(Rnd * 3)
        Select Case charType
            Case 0 : charCode = Int(Rnd * 26) + 65
            Case 1 : charCode = Int(Rnd * 26) + 97
            Case 2 : charCode = Int(Rnd * 10) + 48
        End Select
        nonce = nonce & Chr(charCode)
    Next
    Session("csp_nonce") = nonce
End Sub
Call GenerateCSPNonce()

Response.AddHeader "X-Content-Type-Options", "nosniff"
Response.AddHeader "X-Frame-Options", "SAMEORIGIN"
Response.AddHeader "X-XSS-Protection", "1; mode=block"
Response.AddHeader "Referrer-Policy", "strict-origin-when-cross-origin"
Response.AddHeader "Content-Security-Policy", "default-src 'self'; script-src 'self' 'unsafe-inline' https://cdnjs.cloudflare.com https://code.jquery.com; style-src 'self' 'unsafe-inline' https://cdnjs.cloudflare.com; img-src 'self' data: https:; font-src 'self' https://cdnjs.cloudflare.com; connect-src 'self'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'"
Response.AddHeader "Strict-Transport-Security", "max-age=31536000; includeSubDomains"
Response.AddHeader "Permissions-Policy", "camera=(), microphone=(), geolocation=()"

Function CacheGet(key)
    Dim cacheKey : cacheKey = "CACHE_" & key
    Dim cacheEntry
    cacheEntry = Application(cacheKey)
    If IsEmpty(cacheEntry) Or cacheEntry = "" Then
        CacheGet = Empty
        Exit Function
    End If
    Dim parts : parts = Split(cacheEntry, Chr(1))
    If UBound(parts) < 2 Then
        CacheGet = Empty
        Exit Function
    End If
    Dim cacheTime, cacheTTL
    cacheTime = CDbl(parts(0))
    cacheTTL = CInt(parts(1))
    If DateDiff("s", CDate(cacheTime), Now()) > cacheTTL Then
        CacheGet = Empty
    Else
        CacheGet = parts(2)
    End If
End Function

Sub CacheSet(key, value, ttlSeconds)
    Dim cacheKey : cacheKey = "CACHE_" & key
    Dim cacheEntry
    cacheEntry = CDbl(Now()) & Chr(1) & ttlSeconds & Chr(1) & value
    Application.Lock
    Application(cacheKey) = cacheEntry
    Application.UnLock
End Sub

Function GetCachedSiteSettings(conn)
    Dim cached : cached = CacheGet("SiteSettings_All")
    If Not IsEmpty(cached) Then
        GetCachedSiteSettings = cached
        Exit Function
    End If
    GetCachedSiteSettings = "OK"
End Function

Function GetCachedVolumes(conn)
    Dim cached : cached = CacheGet("Volumes_All")
    If Not IsEmpty(cached) Then
        GetCachedVolumes = cached
        Exit Function
    End If
    GetCachedVolumes = "OK"
End Function

' CDN functions
Dim CDN_DOMAIN
CDN_DOMAIN = "https://cdn.perfumeshop.com"
Dim CDN_ENABLED
CDN_ENABLED = False

Function GetImageCDNURL(localPath)
    If Not CDN_ENABLED Then GetImageCDNURL = localPath : Exit Function
    If InStr(localPath, "http") = 1 Or InStr(localPath, "//") = 1 Then GetImageCDNURL = localPath : Exit Function
    If Left(localPath, 1) <> "/" Then localPath = "/" & localPath
    GetImageCDNURL = CDN_DOMAIN & localPath
End Function

Function GetProductImageCDN(imageURL)
    If imageURL = "" Or IsNull(imageURL) Then
        GetProductImageCDN = GetImageCDNURL("/images/default-product.svg")
    Else
        GetProductImageCDN = GetImageCDNURL(imageURL)
    End If
End Function

Function ConvertImagesToCDN(htmlContent)
    If Not CDN_ENABLED Then ConvertImagesToCDN = htmlContent : Exit Function
    Dim regex, matches, match
    Set regex = New RegExp
    regex.Pattern = "src=""(/images/[^""]*?)"""
    regex.Global = True
    regex.IgnoreCase = True
    Set matches = regex.Execute(htmlContent)
    For Each match In matches
        Dim oldSrc, newSrc
        oldSrc = match.SubMatches(0)
        newSrc = CDN_DOMAIN & oldSrc
        htmlContent = Replace(htmlContent, oldSrc, newSrc)
    Next
    ConvertImagesToCDN = htmlContent
End Function
%>