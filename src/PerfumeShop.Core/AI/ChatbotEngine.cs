using System.Collections.Concurrent;
using System.Text.RegularExpressions;

namespace PerfumeShop.Core.AI;

/// <summary>
/// 智能客服引擎（移植自 ai-service/chatbot_handler.py，增强意图识别和多轮对话状态管理）
/// </summary>
public interface IChatbotEngine
{
    bool IsReady { get; }
    ChatbotResponse ProcessMessage(string message, string sessionId = "", Dictionary<string, object>? context = null);
    List<FaqCategory> GetFaqList();
}

public class ChatbotResponse
{
    public string Reply { get; set; } = "";
    public string Category { get; set; } = "";
    public bool NeedsHuman { get; set; }
    public List<string> Suggestions { get; set; } = new();
}

public class FaqCategory
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
    public List<string> Questions { get; set; } = new();
}

public class ChatbotEngine : IChatbotEngine
{
    private const int MaxSessionHistory = 20;
    private static readonly ConcurrentDictionary<string, List<ChatMessage>> Sessions = new();

    private readonly Dictionary<string, List<FaqEntry>> _faq;

    public bool IsReady { get; } = true;

    public ChatbotEngine()
    {
        _faq = BuildFaq();
    }

    public List<FaqCategory> GetFaqList()
    {
        return _faq.Select(kvp => new FaqCategory
        {
            Category = kvp.Key,
            Name = GetCategoryName(kvp.Key),
            Questions = kvp.Value.Select(q => q.Question).ToList()
        }).ToList();
    }

    public ChatbotResponse ProcessMessage(string message, string sessionId = "", Dictionary<string, object>? context = null)
    {
        message = (message ?? "").Trim();
        context ??= new Dictionary<string, object>();

        // Store in session
        if (!string.IsNullOrEmpty(sessionId))
        {
            var history = Sessions.GetOrAdd(sessionId, _ => new List<ChatMessage>());
            history.Add(new ChatMessage { Role = "user", Content = message });
            if (history.Count > MaxSessionHistory)
                history.RemoveRange(0, history.Count - MaxSessionHistory);
        }

        // 1. Check greetings
        var greeting = CheckGreeting(message);
        if (greeting != null)
            return Respond(greeting, "greeting", sessionId);

        // 2. FAQ knowledge base match
        var faqResult = MatchFaq(message);
        if (faqResult != null)
            return Respond(faqResult.Answer, faqResult.Category, sessionId, faqResult.Related);

        // 3. Intent-based pattern matching (enhanced)
        var intentResponse = MatchIntent(message, sessionId);
        if (intentResponse != null)
            return intentResponse;

        // 4. Fallback
        return Respond(
            "感谢您的咨询！这个问题比较复杂，建议您联系在线客服获取更详细的帮助。\n\n客服电话：400-888-8888\n邮箱：contact@perfumeshop.com\n工作时间：周一至周日 9:00-21:00",
            "fallback", sessionId,
            needsHuman: true,
            suggestions: new List<string> { "联系客服", "发送邮件", "查看帮助中心" });
    }

    private static string? CheckGreeting(string message)
    {
        var greetings = new (string pattern, string reply)[]
        {
            (@"^(你好|您好|hi|hello|hey|嗨|哈喽)",
             "您好！我是香氛定制的智能客服小香。\n\n我可以帮您解答：\n• 订单查询与物流跟踪\n• 退换货政策\n• 香水定制流程\n• 会员与积分\n• 产品推荐\n\n请问有什么可以帮您的？"),
            (@"(谢谢|感谢|thanks|thank)",
             "不客气！如果还有其他问题，随时找我哦~"),
            (@"(再见|拜拜|bye|晚安)",
             "再见！祝您有美好的一天\n如有需要随时回来找我~"),
        };

        foreach (var (pattern, reply) in greetings)
        {
            if (Regex.IsMatch(message, pattern, RegexOptions.IgnoreCase))
                return reply;
        }
        return null;
    }

    private FaqEntry? MatchFaq(string message)
    {
        FaqEntry? bestMatch = null;
        double bestScore = 0;

        foreach (var (category, entries) in _faq)
        {
            foreach (var entry in entries)
            {
                double score = KeywordMatchScore(message, entry.Keywords);
                if (score > bestScore)
                {
                    bestScore = score;
                    bestMatch = entry;
                    bestMatch.Category = category;
                }
            }
        }

        return bestScore >= 0.3 ? bestMatch : null;
    }

    private static double KeywordMatchScore(string message, List<string> keywords)
    {
        if (keywords.Count == 0) return 0.0;
        int matches = 0;
        foreach (var kw in keywords)
        {
            if (kw.Length <= 1)
            {
                // Single-char keywords: require word-boundary match to avoid false positives
                if (Regex.IsMatch(message, $@"(?<!\w){Regex.Escape(kw)}(?!\w)"))
                    matches++;
            }
            else
            {
                if (message.Contains(kw, StringComparison.OrdinalIgnoreCase))
                    matches++;
            }
        }
        return (double)matches / keywords.Count;
    }

    private ChatbotResponse? MatchIntent(string message, string sessionId)
    {
        // Order tracking
        if (Regex.IsMatch(message, @"(订单|查询|跟踪|物流|快递|发货|什么时候|还没收到)"))
            return Respond(
                "请提供您的订单号，我可以帮您查询订单状态。您也可以在\"我的订单\"页面查看最新物流信息。",
                "order", sessionId,
                suggestions: new List<string> { "如何查询订单？", "发货时间是多久？", "如何修改收货地址？" });

        // Return/refund
        if (Regex.IsMatch(message, @"(退货|退款|退换|不想要|取消)"))
            return Respond(
                "我们支持7天无理由退货。如需退换货，请在\"我的订单\"中选择对应订单申请退换。\n\n注意事项：\n• 香水类产品开封后恕不退换\n• 定制香水不支持无理由退货\n• 退款将在3-7个工作日内退回原支付方式",
                "order", sessionId,
                suggestions: new List<string> { "退货流程是什么？", "退款多久到账？", "定制香水能退吗？" });

        // Customization
        if (Regex.IsMatch(message, @"(定制|调配|香调|配方|前调|中调|后调)"))
            return Respond(
                "我们提供专业的香水定制服务！您可以：\n\n1. 前往\"香水定制\"页面选择香调组合\n2. 参加\"香氛测试\"获取AI推荐配方\n3. 联系我们的调香师获取专业建议\n\n定制周期通常为3-5个工作日。",
                "customize", sessionId,
                suggestions: new List<string> { "定制流程是怎样的？", "如何选择香调？", "定制香水多少钱？" });

        // Membership/points
        if (Regex.IsMatch(message, @"(会员|积分|等级|优惠|折扣)"))
            return Respond(
                "我们的会员体系包含：\n\n银卡会员：消费满3000元解锁，享95折\n金卡会员：消费满10000元解锁，享9折\n钻石会员：消费满30000元解锁，享85折\n\n积分规则：每消费1元=1积分，积分可兑换优惠券和小样。",
                "account", sessionId,
                suggestions: new List<string> { "如何获取积分？", "会员有什么权益？", "积分会过期吗？" });

        return null;
    }

    private ChatbotResponse Respond(string reply, string category, string sessionId,
        List<string>? suggestions = null, bool needsHuman = false)
    {
        if (!string.IsNullOrEmpty(sessionId))
        {
            var history = Sessions.GetOrAdd(sessionId, _ => new List<ChatMessage>());
            history.Add(new ChatMessage { Role = "bot", Content = reply });
        }

        return new ChatbotResponse
        {
            Reply = reply,
            Category = category,
            NeedsHuman = needsHuman,
            Suggestions = suggestions ?? new List<string>()
        };
    }

    private static string GetCategoryName(string category) => category switch
    {
        "order" => "订单相关",
        "shipping" => "配送相关",
        "product" => "产品相关",
        "customize" => "定制流程",
        "account" => "账户相关",
        _ => category
    };

    private static Dictionary<string, List<FaqEntry>> BuildFaq()
    {
        return new Dictionary<string, List<FaqEntry>>
        {
            ["order"] = new()
            {
                new FaqEntry
                {
                    Question = "如何查询订单状态？",
                    Keywords = new() { "订单", "查询", "状态", "跟踪" },
                    Answer = "您可以在\"我的订单\"页面查看所有订单状态。订单状态包括：待付款 → 已支付 → 生产中 → 已发货 → 已完成。点击订单号可查看详细信息。",
                    Related = new() { "发货时间是多久？", "如何取消订单？", "修改收货地址" }
                },
                new FaqEntry
                {
                    Question = "发货时间是多久？",
                    Keywords = new() { "发货", "多久", "时间", "什么时候" },
                    Answer = "常规商品付款后1-2个工作日发货，定制香水需要3-5个工作日调配制香后发货。全国包邮，顺丰速运配送。",
                    Related = new() { "如何查询物流？", "可以加急吗？", "配送范围" }
                },
                new FaqEntry
                {
                    Question = "如何取消订单？",
                    Keywords = new() { "取消", "订单", "不想要" },
                    Answer = "在订单状态为\"待付款\"或\"已支付\"时，您可以联系客服取消订单。定制香水进入生产环节后无法取消，请谅解。",
                    Related = new() { "退款多久到账？", "退货流程" }
                },
            },
            ["shipping"] = new()
            {
                new FaqEntry
                {
                    Question = "配送范围和费用？",
                    Keywords = new() { "配送", "包邮", "运费", "范围" },
                    Answer = "全国包邮（港澳台及偏远地区除外），使用顺丰速运配送。满299元免运费，未满299元收取15元运费。",
                },
                new FaqEntry
                {
                    Question = "如何修改收货地址？",
                    Keywords = new() { "修改地址", "地址", "收货", "改地址" },
                    Answer = "在订单未发货前，您可以在\"我的订单\"中修改收货地址。如已发货，请联系客服协助处理。",
                },
            },
            ["product"] = new()
            {
                new FaqEntry
                {
                    Question = "香水能保存多久？",
                    Keywords = new() { "保存", "保质期", "过期", "多久" },
                    Answer = "我们的香水准确保质期为3年（未开封）。开封后建议12-24个月内使用完毕。请存放于阴凉干燥处，避免阳光直射。",
                },
                new FaqEntry
                {
                    Question = "如何选择适合自己的香水？",
                    Keywords = new() { "选择", "适合", "推荐", "怎么选" },
                    Answer = "建议您参加我们的\"香氛测试\"，通过6道简单的问题即可获得AI个性化香调推荐。也可以浏览产品页面按香调筛选。",
                    Related = new() { "香氛测试", "热门推荐", "新品上市" }
                },
            },
            ["customize"] = new()
            {
                new FaqEntry
                {
                    Question = "定制流程是怎样的？",
                    Keywords = new() { "定制", "流程", "步骤", "怎么定制" },
                    Answer = "定制流程：1) 选择香调组合（前调/中调/后调）→ 2) 选择容量和瓶型 → 3) 提交定制需求 → 4) 调香师配制 → 5) 发货配送。全程约3-5个工作日。",
                    Related = new() { "定制香水多少钱？", "可以自己调配吗？", "不满意可以重调吗？" }
                },
                new FaqEntry
                {
                    Question = "定制香水多少钱？",
                    Keywords = new() { "定制", "价格", "多少钱", "费用" },
                    Answer = "定制香水价格取决于容量和瓶型：30ml ¥299起、50ml ¥499起、100ml ¥899起。特殊瓶型和原料可能会有额外费用。",
                },
            },
            ["account"] = new()
            {
                new FaqEntry
                {
                    Question = "如何获取积分？",
                    Keywords = new() { "积分", "获取", "怎么得", "赚" },
                    Answer = "积分获取方式：消费1元=1积分、每日签到+5积分、发表评价+20积分、分享产品+10积分。积分有效期12个月。",
                    Related = new() { "积分能做什么？", "积分会过期吗？" }
                },
                new FaqEntry
                {
                    Question = "会员等级和权益？",
                    Keywords = new() { "会员", "等级", "权益", "银卡", "金卡", "钻石" },
                    Answer = "银卡(消费3000+)享95折、金卡(10000+)享9折+生日礼、钻石(30000+)享85折+专属客服+优先发货。",
                    Related = new() { "如何升级会员？", "会员折扣能叠加吗？" }
                },
            },
        };
    }

    private class ChatMessage
    {
        public string Role { get; set; } = "";
        public string Content { get; set; } = "";
    }

    private class FaqEntry
    {
        public string Question { get; set; } = "";
        public List<string> Keywords { get; set; } = new();
        public string Answer { get; set; } = "";
        public string Category { get; set; } = "";
        public List<string> Related { get; set; } = new();
    }
}
