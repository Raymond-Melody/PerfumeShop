using Microsoft.AspNetCore.Mvc;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 智能香氛匹配 API — 对应 V18 api/fragrance_match.asp
/// </summary>
[ApiController]
[Route("api/fragrance")]
public class FragranceMatchController : ControllerBase
{
    private static readonly Dictionary<string, FragranceProfile> _profiles = new()
    {
        ["floral"] = new FragranceProfile
        {
            Name = "花香调",
            Description = "浪漫优雅，柔和甜美",
            TopNotes = new[] { "佛手柑", "橙花", "紫丁香" },
            MidNotes = new[] { "玫瑰", "茉莉", "依兰" },
            BaseNotes = new[] { "麝香", "琥珀", "檀香" },
            RecommendedCategory = "floral"
        },
        ["oriental"] = new FragranceProfile
        {
            Name = "东方调",
            Description = "温暖浓郁，神秘性感",
            TopNotes = new[] { "肉桂", "豆蔻", "香柠檬" },
            MidNotes = new[] { "檀香", "广藿香", "琥珀" },
            BaseNotes = new[] { "香草", "沉香", "皮革" },
            RecommendedCategory = "oriental"
        },
        ["woody"] = new FragranceProfile
        {
            Name = "木质调",
            Description = "沉稳深邃，自然醇厚",
            TopNotes = new[] { "雪松", "香根草", "柏木" },
            MidNotes = new[] { "檀木", "愈创木", "岩兰草" },
            BaseNotes = new[] { "沉香", "琥珀木", "苔藓" },
            RecommendedCategory = "woody"
        },
        ["fresh"] = new FragranceProfile
        {
            Name = "清新调",
            Description = "清爽活力，自然通透",
            TopNotes = new[] { "柠檬", "薄荷", "青草" },
            MidNotes = new[] { "绿茶", "白松香", "铃兰" },
            BaseNotes = new[] { "白麝香", "雪松", "橡苔" },
            RecommendedCategory = "fresh"
        },
        ["oceanic"] = new FragranceProfile
        {
            Name = "海洋调",
            Description = "清新海洋，自由奔放",
            TopNotes = new[] { "海风", "柑橘", "水生花" },
            MidNotes = new[] { "海藻", "龙涎香", "紫罗兰" },
            BaseNotes = new[] { "浮木", "白麝香", "琥珀" },
            RecommendedCategory = "oceanic"
        },
        ["fruity"] = new FragranceProfile
        {
            Name = "果香调",
            Description = "甜美活泼，青春洋溢",
            TopNotes = new[] { "蜜桃", "荔枝", "黑加仑" },
            MidNotes = new[] { "苹果花", "樱花", "小苍兰" },
            BaseNotes = new[] { "香草", "麝香", "琥珀" },
            RecommendedCategory = "fruity"
        }
    };

    /// <summary>GET /api/fragrance/profiles — 获取所有香调类型</summary>
    [HttpGet("profiles")]
    public IActionResult GetProfiles()
    {
        return Ok(new { success = true, data = _profiles.Select(p => new
        {
            key = p.Key,
            name = p.Value.Name,
            description = p.Value.Description
        }) });
    }

    /// <summary>GET /api/fragrance/profile/{key} — 获取具体香调详情</summary>
    [HttpGet("profile/{key}")]
    public IActionResult GetProfile(string key)
    {
        if (_profiles.TryGetValue(key.ToLower(), out var profile))
            return Ok(new { success = true, data = profile });
        return NotFound(new { success = false, message = "未找到该香调类型" });
    }

    /// <summary>POST /api/fragrance/match — 根据偏好匹配香调</summary>
    [HttpPost("match")]
    public IActionResult Match([FromBody] FragranceMatchRequest req)
    {
        if (req.Answers == null || req.Answers.Count == 0)
            return BadRequest(new { success = false, message = "请提供香氛偏好答案" });

        // 简单的评分匹配算法
        var scores = new Dictionary<string, int>();
        foreach (var profile in _profiles)
            scores[profile.Key] = 0;

        // Q1: 场合偏好
        var q1 = req.Answers.GetValueOrDefault("occasion", "").ToLower();
        switch (q1)
        {
            case "daily": scores["fresh"] += 3; scores["fruity"] += 2; break;
            case "work": scores["woody"] += 3; scores["floral"] += 2; break;
            case "date": scores["floral"] += 4; scores["oriental"] += 3; break;
            case "party": scores["oriental"] += 4; scores["fruity"] += 2; break;
            case "sport": scores["oceanic"] += 4; scores["fresh"] += 3; break;
        }

        // Q2: 性格偏好
        var q2 = req.Answers.GetValueOrDefault("personality", "").ToLower();
        switch (q2)
        {
            case "elegant": scores["floral"] += 4; break;
            case "mysterious": scores["oriental"] += 4; break;
            case "natural": scores["woody"] += 3; scores["fresh"] += 2; break;
            case "energetic": scores["oceanic"] += 3; scores["fruity"] += 3; break;
            case "romantic": scores["floral"] += 3; scores["fruity"] += 2; break;
        }

        // Q3: 季节偏好
        var q3 = req.Answers.GetValueOrDefault("season", "").ToLower();
        switch (q3)
        {
            case "spring": scores["floral"] += 3; scores["fruity"] += 2; break;
            case "summer": scores["fresh"] += 4; scores["oceanic"] += 4; break;
            case "autumn": scores["woody"] += 3; scores["oriental"] += 2; break;
            case "winter": scores["oriental"] += 4; scores["woody"] += 3; break;
        }

        // Q4: 浓度偏好
        var q4 = req.Answers.GetValueOrDefault("intensity", "").ToLower();
        switch (q4)
        {
            case "light": scores["fresh"] += 3; scores["oceanic"] += 2; break;
            case "moderate": scores["floral"] += 2; scores["fruity"] += 2; break;
            case "strong": scores["oriental"] += 4; scores["woody"] += 3; break;
        }

        // 取最高分
        var topMatch = scores.OrderByDescending(s => s.Value).First();
        var runnerUp = scores.OrderByDescending(s => s.Value).Skip(1).First();
        var matchedProfile = _profiles[topMatch.Key];

        return Ok(new
        {
            success = true,
            data = new
            {
                primaryKey = topMatch.Key,
                primary = new { matchedProfile.Name, matchedProfile.Description, matchedProfile.TopNotes, matchedProfile.MidNotes, matchedProfile.BaseNotes },
                secondaryKey = runnerUp.Key,
                secondaryName = _profiles[runnerUp.Key].Name,
                scores = scores.Select(s => new { key = s.Key, name = _profiles[s.Key].Name, score = s.Value }).OrderByDescending(s => s.score)
            }
        });
    }
}

public class FragranceMatchRequest
{
    public Dictionary<string, string> Answers { get; set; } = new();
}

public class FragranceProfile
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public string[] TopNotes { get; set; } = Array.Empty<string>();
    public string[] MidNotes { get; set; } = Array.Empty<string>();
    public string[] BaseNotes { get; set; } = Array.Empty<string>();
    public string RecommendedCategory { get; set; } = "";
}
