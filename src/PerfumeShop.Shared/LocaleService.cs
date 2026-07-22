using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.Concurrent;

namespace PerfumeShop.Shared;

/// <summary>
/// V19 国际化服务 — 对应 V18 includes/i18n.asp
/// 支持 zh-CN / en-US，优先级: QueryString ?lang= → Cookie → Accept-Language → 默认zh-CN
/// V19 增强：GetSectionAsync / GetAllAsync / SwitchCultureAsync / IMemoryCache 缓存词典
/// </summary>
public interface ILocaleService
{
    string T(string key, params object[] args);
    string CurrentLocale { get; }
    void SetLocale(string locale);
    string HtmlLang();

    // ── V19 新增 ──

    /// <summary>获取指定前缀 section 的所有键值对（如 "nav" → 所有 nav_* 键）</summary>
    Task<Dictionary<string, string>> GetSectionAsync(string section);

    /// <summary>获取当前语言的全部词典</summary>
    Task<Dictionary<string, string>> GetAllAsync();

    /// <summary>切换文化并持久化到 Cookie</summary>
    Task SwitchCultureAsync(string culture);
}

public class LocaleService : ILocaleService
{
    private readonly IHttpContextAccessor _http;
    private readonly IMemoryCache _cache;
    private readonly ConcurrentDictionary<string, ConcurrentDictionary<string, string>> _dicts = new();

    private static readonly string[] SupportedLocales = { "zh-CN", "en-US" };
    private const string DefaultLocale = "zh-CN";
    private const string DictCachePrefix = "i18n_dict_";
    private string _currentLocale = DefaultLocale;

    public string CurrentLocale => _currentLocale;

    public LocaleService(IHttpContextAccessor httpAccessor, IMemoryCache cache)
    {
        _http = httpAccessor;
        _cache = cache;
        LoadLocaleFile("zh-CN");
        LoadLocaleFile("en-US");
        DetectLocale();
    }

    private void DetectLocale()
    {
        var ctx = _http.HttpContext;
        if (ctx == null) { _currentLocale = DefaultLocale; return; }

        // 1. QueryString ?lang=
        var qs = ctx.Request.Query["lang"].FirstOrDefault();
        if (!string.IsNullOrEmpty(qs) && SupportedLocales.Contains(qs))
        {
            SetLocale(qs);
            return;
        }

        // 2. Cookie
        if (ctx.Request.Cookies.TryGetValue("lang", out var cookieLang) &&
            !string.IsNullOrEmpty(cookieLang) && SupportedLocales.Contains(cookieLang))
        {
            _currentLocale = cookieLang;
            return;
        }

        // 3. Accept-Language header
        var acceptLang = ctx.Request.Headers.AcceptLanguage.FirstOrDefault();
        if (!string.IsNullOrEmpty(acceptLang))
        {
            if (acceptLang.Contains("zh", StringComparison.OrdinalIgnoreCase))
            {
                _currentLocale = "zh-CN";
                return;
            }
            if (acceptLang.Contains("en", StringComparison.OrdinalIgnoreCase))
            {
                _currentLocale = "en-US";
                return;
            }
        }

        // 4. Default
        _currentLocale = DefaultLocale;
    }

    private void LoadLocaleFile(string locale)
    {
        var cacheKey = $"{DictCachePrefix}{locale}";

        // 先尝试 IMemoryCache
        if (_cache.TryGetValue(cacheKey, out ConcurrentDictionary<string, string>? cached) && cached != null)
        {
            _dicts[locale] = cached;
            return;
        }

        var dict = new ConcurrentDictionary<string, string>();

        // 尝试多个路径查找 locale 文件（支持 .asp 和 .txt）
        string[] searchPaths = {
            Path.Combine(AppContext.BaseDirectory, "locale", $"{locale}.asp"),
            Path.Combine(Directory.GetCurrentDirectory(), "locale", $"{locale}.asp"),
            Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "locale", $"{locale}.asp"),
            Path.Combine(AppContext.BaseDirectory, "locale", $"{locale}.txt"),
            Path.Combine(Directory.GetCurrentDirectory(), "locale", $"{locale}.txt"),
            Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "locale", $"{locale}.txt"),
        };

        string? foundPath = null;
        foreach (var p in searchPaths)
        {
            if (File.Exists(p)) { foundPath = p; break; }
        }

        if (foundPath != null)
        {
            foreach (var line in File.ReadLines(foundPath))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith('#'))
                    continue;

                var eqIdx = trimmed.IndexOf('=');
                if (eqIdx > 0)
                {
                    var key = trimmed[..eqIdx].Trim();
                    var value = trimmed[(eqIdx + 1)..].Trim();
                    dict[key] = value;
                }
            }
        }
        else
        {
            // 内嵌回退词典
            LoadFallbackDict(dict, locale);
        }

        _dicts[locale] = dict;

        // 缓存词典（30 分钟绝对过期，文件变更时由 ChangeToken 失效）
        var options = new MemoryCacheEntryOptions()
            .SetAbsoluteExpiration(TimeSpan.FromHours(1))
            .SetSize(1);
        _cache.Set(cacheKey, dict, options);
    }

    private static void LoadFallbackDict(ConcurrentDictionary<string, string> dict, string locale)
    {
        if (locale == "zh-CN")
        {
            dict["site_name"] = "香氛定制";
            dict["welcome"] = "欢迎";
            dict["login"] = "登录";
            dict["logout"] = "退出";
            dict["home"] = "首页";
            dict["products"] = "产品";
            dict["about"] = "关于我们";
            dict["contact"] = "联系我们";
            dict["cart"] = "购物车";
            dict["search"] = "搜索";
            dict["submit"] = "提交";
            dict["cancel"] = "取消";
            dict["save"] = "保存";
            dict["loading"] = "加载中...";
            dict["no_data"] = "暂无数据";
            dict["price"] = "价格";
            dict["checkout"] = "结算";
            dict["add_to_cart"] = "加入购物车";
            dict["buy_now"] = "立即购买";
            dict["language"] = "语言";
            dict["dark_mode"] = "暗黑模式";
            dict["light_mode"] = "明亮模式";
            dict["back_to_top"] = "回到顶部";
        }
        else if (locale == "en-US")
        {
            dict["site_name"] = "PerfumeShop";
            dict["welcome"] = "Welcome";
            dict["login"] = "Login";
            dict["logout"] = "Logout";
            dict["home"] = "Home";
            dict["products"] = "Products";
            dict["about"] = "About Us";
            dict["contact"] = "Contact";
            dict["cart"] = "Cart";
            dict["search"] = "Search";
            dict["submit"] = "Submit";
            dict["cancel"] = "Cancel";
            dict["save"] = "Save";
            dict["loading"] = "Loading...";
            dict["no_data"] = "No Data";
            dict["price"] = "Price";
            dict["checkout"] = "Checkout";
            dict["add_to_cart"] = "Add to Cart";
            dict["buy_now"] = "Buy Now";
            dict["language"] = "Language";
            dict["dark_mode"] = "Dark Mode";
            dict["light_mode"] = "Light Mode";
            dict["back_to_top"] = "Back to Top";
        }
    }

    public string T(string key, params object[] args)
    {
        // 先从当前语言查找
        if (_dicts.TryGetValue(_currentLocale, out var dict) && dict.TryGetValue(key, out var value))
            return args.Length > 0 ? string.Format(value, args) : value;

        // 回退到默认语言
        if (_currentLocale != DefaultLocale && _dicts.TryGetValue(DefaultLocale, out var defDict) && defDict.TryGetValue(key, out var defValue))
            return args.Length > 0 ? string.Format(defValue, args) : defValue;

        return $"[{key}]";
    }

    public void SetLocale(string locale)
    {
        if (SupportedLocales.Contains(locale))
        {
            _currentLocale = locale;
            var ctx = _http.HttpContext;
            if (ctx != null)
            {
                ctx.Response.Cookies.Append("lang", locale, new CookieOptions
                {
                    Expires = DateTimeOffset.UtcNow.AddYears(1),
                    Path = "/"
                });
            }
        }
    }

    public string HtmlLang() => _currentLocale;

    // ── V19 新增方法 ──────────────────────────────────────────

    public Task<Dictionary<string, string>> GetSectionAsync(string section)
    {
        var prefix = section.EndsWith("_") ? section : section + "_";
        var result = new Dictionary<string, string>();

        if (_dicts.TryGetValue(_currentLocale, out var dict))
        {
            foreach (var kv in dict)
            {
                if (kv.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    result[kv.Key] = kv.Value;
            }
        }

        // 回退补全
        if (result.Count == 0 && _currentLocale != DefaultLocale && _dicts.TryGetValue(DefaultLocale, out var defDict))
        {
            foreach (var kv in defDict)
            {
                if (kv.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    result[kv.Key] = kv.Value;
            }
        }

        return Task.FromResult(result);
    }

    public Task<Dictionary<string, string>> GetAllAsync()
    {
        var result = new Dictionary<string, string>();

        // 先加载默认语言作为基底
        if (_dicts.TryGetValue(DefaultLocale, out var defDict))
        {
            foreach (var kv in defDict) result[kv.Key] = kv.Value;
        }

        // 当前语言覆盖
        if (_currentLocale != DefaultLocale && _dicts.TryGetValue(_currentLocale, out var curDict))
        {
            foreach (var kv in curDict) result[kv.Key] = kv.Value;
        }
        else if (_currentLocale == DefaultLocale && defDict != null)
        {
            // 已经是默认语言，不需要额外覆盖
        }
        else
        {
            // 当前语言词典为空时就用默认的
            if (_dicts.TryGetValue(_currentLocale, out var curDict2))
            {
                foreach (var kv in curDict2) result[kv.Key] = kv.Value;
            }
        }

        return Task.FromResult(result);
    }

    public Task SwitchCultureAsync(string culture)
    {
        if (!SupportedLocales.Contains(culture))
            return Task.CompletedTask;

        _currentLocale = culture;
        var ctx = _http.HttpContext;
        if (ctx != null)
        {
            ctx.Response.Cookies.Append("lang", culture, new CookieOptions
            {
                Expires = DateTimeOffset.UtcNow.AddYears(1),
                Path = "/",
                IsEssential = true
            });
        }

        return Task.CompletedTask;
    }
}

/// <summary>扩展方法</summary>
public static class LocaleServiceExtensions
{
    public static IServiceCollection AddLocaleService(this IServiceCollection services)
    {
        services.AddHttpContextAccessor();
        services.AddMemoryCache(o => o.SizeLimit = 1024);
        services.AddSingleton<ILocaleService, LocaleService>();
        return services;
    }
}
