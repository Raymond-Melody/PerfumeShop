using PerfumeShop.Core.AI;

namespace PerfumeShop.IntegrationTests.AI;

/// <summary>
/// 香氛匹配引擎测试 — 评分聚合
/// </summary>
public class FragranceMatcherTests
{
    private readonly IFragranceMatcher _matcher = new FragranceMatcher();

    [Fact]
    public void Match_FloralStyle_ReturnsFloralAsTop()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "floral",
            ["occasion"] = "date",
            ["season"] = "spring",
            ["gender"] = "female",
            ["intensity"] = "medium",
            ["budget"] = "mid"
        };

        var result = _matcher.Match(answers);

        Assert.NotNull(result);
        Assert.True(result.MatchedFamilies.Count > 0);
        Assert.Equal("floral", result.MatchedFamilies[0].Family);
        Assert.True(result.MatchedFamilies[0].Score > 0);
    }

    [Fact]
    public void Match_WoodyMale_ReturnsWoodyAsTop()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "woody",
            ["occasion"] = "work",
            ["season"] = "autumn",
            ["gender"] = "male",
            ["intensity"] = "strong",
            ["budget"] = "premium"
        };

        var result = _matcher.Match(answers);

        Assert.NotNull(result);
        Assert.Equal("woody", result.MatchedFamilies[0].Family);
    }

    [Fact]
    public void Match_CitrusFreshSummer_ReturnsHighCitrusScore()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "citrus",
            ["occasion"] = "sport",
            ["season"] = "summer",
            ["gender"] = "unisex",
            ["intensity"] = "light",
            ["budget"] = "entry"
        };

        var result = _matcher.Match(answers);

        Assert.NotNull(result);
        // Citrus should be in top matches
        Assert.Contains(result.MatchedFamilies, f => f.Family == "citrus");
        Assert.Equal("建议选择EDT淡香水，清新不张扬", result.IntensityAdvice);
    }

    [Fact]
    public void Match_OrientalParty_ReturnsHighOrientalScore()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "oriental",
            ["occasion"] = "party",
            ["season"] = "winter",
            ["gender"] = "female",
            ["intensity"] = "strong",
            ["budget"] = "premium"
        };

        var result = _matcher.Match(answers);

        Assert.NotNull(result);
        Assert.Equal("oriental", result.MatchedFamilies[0].Family);
        Assert.Equal("建议选择Parfum浓香精，持久浓郁", result.IntensityAdvice);
    }

    [Fact]
    public void Match_ReturnsNoteRecommendations()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "floral",
            ["occasion"] = "daily",
            ["season"] = "spring",
            ["gender"] = "female"
        };

        var result = _matcher.Match(answers);

        Assert.NotNull(result.RecommendedNotes);
        Assert.True(result.RecommendedNotes.Top.Count > 0);
        Assert.True(result.RecommendedNotes.Middle.Count > 0);
        Assert.True(result.RecommendedNotes.Base.Count > 0);
    }

    [Fact]
    public void Match_NoteRecommendationsMaxThreePerLayer()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "floral",
            ["occasion"] = "date",
            ["season"] = "spring",
            ["gender"] = "female"
        };

        var result = _matcher.Match(answers);

        Assert.True(result.RecommendedNotes.Top.Count <= 3);
        Assert.True(result.RecommendedNotes.Middle.Count <= 3);
        Assert.True(result.RecommendedNotes.Base.Count <= 3);
    }

    [Fact]
    public void Match_ScoreAggregation_MultipleInputsCombine()
    {
        // When style + occasion + season all point to floral, score should be high
        var answers = new Dictionary<string, string>
        {
            ["style"] = "floral",   // +2.0
            ["occasion"] = "date",  // +1.5 for floral
            ["season"] = "spring",  // +1.0 for floral
            ["gender"] = "female"   // +1.2 for floral
        };

        var result = _matcher.Match(answers);

        // Floral score should be significantly high due to multiple inputs
        var floralScore = result.MatchedFamilies.FirstOrDefault(f => f.Family == "floral")?.Score ?? 0;
        Assert.True(floralScore > 2.0, $"Floral score should be > 2.0, got {floralScore}");
    }

    [Fact]
    public void Match_EmptyAnswers_ReturnsEmptyResult()
    {
        var answers = new Dictionary<string, string>();

        var result = _matcher.Match(answers);

        Assert.NotNull(result);
        Assert.Empty(result.MatchedFamilies);
    }

    [Fact]
    public void Match_ReturnsTopThreeFamilies()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "fresh",
            ["occasion"] = "daily",
            ["season"] = "summer",
            ["gender"] = "unisex"
        };

        var result = _matcher.Match(answers);

        Assert.True(result.MatchedFamilies.Count <= 3);
        Assert.True(result.MatchedFamilies.Count > 0);
    }

    [Fact]
    public void GetQuizQuestions_ReturnsSixQuestions()
    {
        var questions = _matcher.GetQuizQuestions();

        Assert.NotNull(questions);
        Assert.Equal(6, questions.Count);
        Assert.Contains(questions, q => q.Id == "style");
        Assert.Contains(questions, q => q.Id == "occasion");
        Assert.Contains(questions, q => q.Id == "season");
        Assert.Contains(questions, q => q.Id == "gender");
        Assert.Contains(questions, q => q.Id == "intensity");
        Assert.Contains(questions, q => q.Id == "budget");
    }

    [Fact]
    public void GetQuizQuestions_EachHasOptions()
    {
        var questions = _matcher.GetQuizQuestions();

        foreach (var q in questions)
        {
            Assert.NotEmpty(q.Options);
            Assert.All(q.Options, o =>
            {
                Assert.NotEmpty(o.Value);
                Assert.NotEmpty(o.Label);
            });
        }
    }

    [Fact]
    public void Match_BudgetLevel_IsPreserved()
    {
        var answers = new Dictionary<string, string>
        {
            ["style"] = "floral",
            ["budget"] = "premium"
        };

        var result = _matcher.Match(answers);

        Assert.Equal("premium", result.BudgetLevel);
    }

    [Fact]
    public void Matcher_IsReady()
    {
        Assert.True(_matcher.IsReady);
    }

    [Fact]
    public void Match_FamilyWeightApplied()
    {
        // Floral has weight 0.9, green has 0.7
        // Both should have their raw scores multiplied by weight
        var answers = new Dictionary<string, string>
        {
            ["style"] = "floral",
        };

        var result = _matcher.Match(answers);
        var floral = result.MatchedFamilies.FirstOrDefault(f => f.Family == "floral");

        Assert.NotNull(floral);
        // 2.0 * 0.9 = 1.8
        Assert.Equal(1.8, floral.Score);
    }
}
