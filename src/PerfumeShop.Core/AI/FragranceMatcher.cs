namespace PerfumeShop.Core.AI;

/// <summary>
/// 智能香氛匹配引擎（移植自 ai-service/fragrance_matcher.py，增强评分聚合）
/// </summary>
public interface IFragranceMatcher
{
    bool IsReady { get; }
    List<QuizQuestion> GetQuizQuestions();
    FragranceMatchResult Match(Dictionary<string, string> answers);
}

public class QuizQuestion
{
    public string Id { get; set; } = "";
    public string Question { get; set; } = "";
    public List<QuizOption> Options { get; set; } = new();
}

public class QuizOption
{
    public string Value { get; set; } = "";
    public string Label { get; set; } = "";
}

public class FragranceMatchResult
{
    public List<MatchedFamily> MatchedFamilies { get; set; } = new();
    public NoteRecommendations RecommendedNotes { get; set; } = new();
    public string IntensityAdvice { get; set; } = "";
    public string BudgetLevel { get; set; } = "";
}

public class MatchedFamily
{
    public string Family { get; set; } = "";
    public double Score { get; set; }
    public List<string> Keywords { get; set; } = new();
}

public class NoteRecommendations
{
    public List<string> Top { get; set; } = new();
    public List<string> Middle { get; set; } = new();
    public List<string> Base { get; set; } = new();
}

public class FragranceMatcher : IFragranceMatcher
{
    private static readonly Dictionary<string, NoteFamilyInfo> NoteFamilies = new()
    {
        ["floral"] = new NoteFamilyInfo { Keywords = new() { "花香", "玫瑰", "茉莉", "百合", "温柔", "浪漫", "女性" }, Weight = 0.9 },
        ["citrus"] = new NoteFamilyInfo { Keywords = new() { "柑橘", "柠檬", "清新", "活力", "阳光", "清爽", "夏天" }, Weight = 0.9 },
        ["woody"] = new NoteFamilyInfo { Keywords = new() { "木质", "檀香", "雪松", "沉稳", "成熟", "中性", "秋天" }, Weight = 0.8 },
        ["oriental"] = new NoteFamilyInfo { Keywords = new() { "东方", "琥珀", "香草", "神秘", "性感", "浓郁", "夜晚" }, Weight = 0.8 },
        ["fresh"] = new NoteFamilyInfo { Keywords = new() { "海洋", "水生", "绿叶", "运动", "干净", "春天", "日常" }, Weight = 0.85 },
        ["fruity"] = new NoteFamilyInfo { Keywords = new() { "果香", "桃子", "莓果", "甜美", "活泼", "年轻", "派对" }, Weight = 0.75 },
        ["green"] = new NoteFamilyInfo { Keywords = new() { "青草", "绿茶", "自然", "素雅", "文艺", "中性" }, Weight = 0.7 },
    };

    private static readonly Dictionary<string, List<string>> OccasionMap = new()
    {
        ["daily"] = new() { "fresh", "citrus", "green" },
        ["work"] = new() { "woody", "green", "fresh" },
        ["date"] = new() { "floral", "oriental", "fruity" },
        ["party"] = new() { "oriental", "fruity", "floral" },
        ["sport"] = new() { "fresh", "citrus" },
        ["formal"] = new() { "woody", "oriental", "floral" },
    };

    private static readonly Dictionary<string, List<string>> SeasonMap = new()
    {
        ["spring"] = new() { "floral", "green", "fruity" },
        ["summer"] = new() { "citrus", "fresh", "fruity" },
        ["autumn"] = new() { "woody", "oriental" },
        ["winter"] = new() { "oriental", "woody", "floral" },
    };

    private static readonly Dictionary<string, List<string>> GenderMap = new()
    {
        ["female"] = new() { "floral", "fruity", "oriental" },
        ["male"] = new() { "woody", "fresh", "citrus" },
        ["unisex"] = new() { "citrus", "green", "woody" },
    };

    public bool IsReady { get; } = true;

    public List<QuizQuestion> GetQuizQuestions()
    {
        return new List<QuizQuestion>
        {
            new() { Id = "style", Question = "你更喜欢哪种风格的香水？", Options = new()
            {
                new() { Value = "floral", Label = "花香调 - 浪漫优雅" },
                new() { Value = "fresh", Label = "清新调 - 干净自然" },
                new() { Value = "woody", Label = "木质调 - 沉稳大气" },
                new() { Value = "oriental", Label = "东方调 - 神秘性感" },
                new() { Value = "citrus", Label = "柑橘调 - 活力阳光" },
            }},
            new() { Id = "occasion", Question = "你主要在什么场合使用香水？", Options = new()
            {
                new() { Value = "daily", Label = "日常通勤" },
                new() { Value = "work", Label = "工作办公" },
                new() { Value = "date", Label = "约会聚会" },
                new() { Value = "party", Label = "派对晚宴" },
                new() { Value = "sport", Label = "运动休闲" },
            }},
            new() { Id = "season", Question = "你最喜欢哪个季节？", Options = new()
            {
                new() { Value = "spring", Label = "春天" },
                new() { Value = "summer", Label = "夏天" },
                new() { Value = "autumn", Label = "秋天" },
                new() { Value = "winter", Label = "冬天" },
            }},
            new() { Id = "gender", Question = "你偏好的香水类型？", Options = new()
            {
                new() { Value = "female", Label = "女士香水" },
                new() { Value = "male", Label = "男士香水" },
                new() { Value = "unisex", Label = "中性香水" },
            }},
            new() { Id = "intensity", Question = "你喜欢多浓的香水？", Options = new()
            {
                new() { Value = "light", Label = "淡雅清新 (EDT)" },
                new() { Value = "medium", Label = "适中持久 (EDP)" },
                new() { Value = "strong", Label = "浓郁持久 (Parfum)" },
            }},
            new() { Id = "budget", Question = "你的预算范围？", Options = new()
            {
                new() { Value = "entry", Label = "入门级 (200-500元)" },
                new() { Value = "mid", Label = "中端 (500-1500元)" },
                new() { Value = "premium", Label = "高端 (1500元以上)" },
            }},
        };
    }

    public FragranceMatchResult Match(Dictionary<string, string> answers)
    {
        var scores = new Dictionary<string, double>();

        // Score from direct style preference (weight 2.0)
        var style = answers.GetValueOrDefault("style", "");
        if (NoteFamilies.ContainsKey(style))
            AddScore(scores, style, 2.0);

        // Score from occasion (weight 1.5)
        var occasion = answers.GetValueOrDefault("occasion", "");
        if (OccasionMap.TryGetValue(occasion, out var occFamilies))
            foreach (var f in occFamilies) AddScore(scores, f, 1.5);

        // Score from season (weight 1.0)
        var season = answers.GetValueOrDefault("season", "");
        if (SeasonMap.TryGetValue(season, out var seaFamilies))
            foreach (var f in seaFamilies) AddScore(scores, f, 1.0);

        // Score from gender (weight 1.2)
        var gender = answers.GetValueOrDefault("gender", "");
        if (GenderMap.TryGetValue(gender, out var genFamilies))
            foreach (var f in genFamilies) AddScore(scores, f, 1.2);

        // Apply family weights
        foreach (var family in scores.Keys.ToList())
        {
            if (NoteFamilies.TryGetValue(family, out var info))
                scores[family] *= info.Weight;
        }

        // Sort by score descending
        var sorted = scores.OrderByDescending(x => x.Value).ToList();

        var noteRecs = GetNoteRecommendations(sorted);

        var intensity = answers.GetValueOrDefault("intensity", "medium");
        var intensityAdvice = intensity switch
        {
            "light" => "建议选择EDT淡香水，清新不张扬",
            "medium" => "建议选择EDP淡香精，持久适中",
            "strong" => "建议选择Parfum浓香精，持久浓郁",
            _ => ""
        };

        return new FragranceMatchResult
        {
            MatchedFamilies = sorted.Take(3).Select(x => new MatchedFamily
            {
                Family = x.Key,
                Score = Math.Round(x.Value, 2),
                Keywords = NoteFamilies.TryGetValue(x.Key, out var nf) ? nf.Keywords : new()
            }).ToList(),
            RecommendedNotes = noteRecs,
            IntensityAdvice = intensityAdvice,
            BudgetLevel = answers.GetValueOrDefault("budget", "mid")
        };
    }

    private static void AddScore(Dictionary<string, double> scores, string family, double value)
    {
        scores[family] = scores.TryGetValue(family, out var existing) ? existing + value : value;
    }

    private static NoteRecommendations GetNoteRecommendations(List<KeyValuePair<string, double>> sortedFamilies)
    {
        var noteMap = new Dictionary<string, NoteLayers>
        {
            ["floral"] = new() { Top = new() { "佛手柑", "粉红胡椒" }, Middle = new() { "玫瑰", "茉莉", "鸢尾花" }, Base = new() { "麝香", "琥珀" } },
            ["citrus"] = new() { Top = new() { "柠檬", "葡萄柚", "佛手柑" }, Middle = new() { "橙花", "薄荷" }, Base = new() { "雪松", "白麝香" } },
            ["woody"] = new() { Top = new() { "香柠檬", "胡椒" }, Middle = new() { "雪松", "檀香木" }, Base = new() { "香根草", "广藿香", "皮革" } },
            ["oriental"] = new() { Top = new() { "肉桂", "小豆蔻" }, Middle = new() { "琥珀", "香草", "零陵香豆" }, Base = new() { "檀香", "麝香", "广藿香" } },
            ["fresh"] = new() { Top = new() { "柑橘", "海洋" }, Middle = new() { "薰衣草", "迷迭香" }, Base = new() { "白麝香", "苔藓" } },
            ["fruity"] = new() { Top = new() { "桃子", "黑加仑" }, Middle = new() { "玫瑰", "紫罗兰" }, Base = new() { "香草", "麝香" } },
            ["green"] = new() { Top = new() { "佛手柑", "绿叶" }, Middle = new() { "绿茶", "茉莉" }, Base = new() { "白麝香", "雪松" } },
        };

        var recommendations = new NoteRecommendations();

        foreach (var (family, _) in sortedFamilies.Take(2))
        {
            if (!noteMap.TryGetValue(family, out var layers)) continue;
            foreach (var note in layers.Top)
                if (!recommendations.Top.Contains(note)) recommendations.Top.Add(note);
            foreach (var note in layers.Middle)
                if (!recommendations.Middle.Contains(note)) recommendations.Middle.Add(note);
            foreach (var note in layers.Base)
                if (!recommendations.Base.Contains(note)) recommendations.Base.Add(note);
        }

        recommendations.Top = recommendations.Top.Take(3).ToList();
        recommendations.Middle = recommendations.Middle.Take(3).ToList();
        recommendations.Base = recommendations.Base.Take(3).ToList();

        return recommendations;
    }

    private class NoteFamilyInfo
    {
        public List<string> Keywords { get; set; } = new();
        public double Weight { get; set; }
    }

    private class NoteLayers
    {
        public List<string> Top { get; set; } = new();
        public List<string> Middle { get; set; } = new();
        public List<string> Base { get; set; } = new();
    }
}
