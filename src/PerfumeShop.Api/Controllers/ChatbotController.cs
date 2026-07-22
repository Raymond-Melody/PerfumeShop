using Microsoft.AspNetCore.Mvc;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 智能客服 API — 对应 V18 api/chatbot.asp
/// </summary>
[ApiController]
[Route("api/chatbot")]
public class ChatbotController : ControllerBase
{
    private static readonly Dictionary<string, string[]> _faq = new(StringComparer.OrdinalIgnoreCase)
    {
        ["运费"] = new[] { "全场满299元包邮，不满299元收取15元运费，偏远地区可能略有不同。" },
        ["配送"] = new[] { "下单后1-3个工作日发货，快递一般3-7天送达。您可以在「我的订单」中查看物流信息。" },
        ["退货"] = new[] { "商品未拆封可在7天内申请退货退款，定制香水因个性化制作不支持无理由退货哦。" },
        ["定制"] = new[] { "您可以在「定制香氛」页面选择前调、中调、后调香调，我们为您调配专属香水！" },
        ["秒杀"] = new[] { "限时秒杀活动每天更新，价格超低，手慢无！关注首页秒杀专区获取最新活动。" },
        ["拼团"] = new[] { "拼团活动：邀请好友一起购买即可享受超低拼团价，2人成团，超时未成团自动退款。" },
        ["订阅"] = new[] { "订阅盒子：月度/季度/年度订阅计划，每月自动配送精选香水，可随时取消。" },
        ["积分"] = new[] { "每消费1元获得1积分，积分可兑换礼品或抵扣现金。每日签到也可获得积分哦！" },
        ["优惠券"] = new[] { "新用户注册可获得新人优惠券，也可关注平台活动获取限时优惠券。" },
        ["会员"] = new[] { "会员等级分为青铜、白银、黄金、钻石，等级越高享受的折扣和权益越多。" },
        ["支付"] = new[] { "支持微信支付、支付宝、银行卡支付。支付过程中如遇问题请联系客服。" },
        ["订单"] = new[] { "您可以在「我的订单」中查看所有订单状态。如有问题请联系客服处理。" },
        ["联系"] = new[] { "客服电话：400-888-8888，工作时间：周一至周五 9:00-18:00。也可发送邮件至 contact@perfumeshop.com。" },
        ["香氛"] = new[] { "您可以在「香氛测试」页面进行简单的问答测试，我们会根据您的偏好推荐最适合的香水。" },
        ["成分"] = new[] { "我们的香水使用法国进口天然香精原料，不含酒精替代品，安全温和，敏感肌肤也可使用。" },
        ["注册"] = new[] { "点击页面右上角「注册」按钮，填写邮箱和密码即可完成注册。注册即送新人优惠券！" },
        ["密码"] = new[] { "忘记密码可点击登录页面「忘记密码」链接，通过注册邮箱重置密码。" },
        ["热门"] = new[] { "热门商品有：晨露玫瑰淡香水、琥珀木香男士淡香水、海洋清风中性香等，欢迎到商品页面查看。" },
    };

    private static readonly string[] _greetings = {
        "您好！我是您的香氛顾问小C，请问有什么可以帮您的？",
        "欢迎来到Custom Fragrance！我可以帮您解答关于香水定制、订单、配送等任何问题~",
        "嗨！有什么香氛方面的问题想问我吗？"
    };

    private static readonly string[] _fallbacks = {
        "抱歉，我不太确定这个问题的答案。您可以拨打客服电话 400-888-8888 或发送邮件至 contact@perfumeshop.com 获取帮助。",
        "这个问题我需要了解一下，建议您联系客服获取更准确的答案哦～"
    };

    /// <summary>POST /api/chatbot/message — 发送客服消息</summary>
    [HttpPost("message")]
    public IActionResult Message([FromBody] ChatbotRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Message))
            return Ok(new { success = true, reply = "请告诉我您的问题，我会尽力帮您解答~" });

        var msg = req.Message.Trim();

        // 问候检测
        if (ContainsAny(msg, "你好", "嗨", "hello", "hi", "在吗", "您好"))
            return Ok(new { success = true, reply = GetRandom(_greetings) });

        // 感谢检测
        if (ContainsAny(msg, "谢谢", "thank", "感谢", "多谢"))
            return Ok(new { success = true, reply = "不客气！很高兴能帮到您，还有其他问题随时问我~" });

        // FAQ 关键词匹配
        foreach (var kvp in _faq)
        {
            if (msg.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase))
                return Ok(new { success = true, reply = GetRandom(kvp.Value) });
        }

        // 模糊推荐
        if (ContainsAny(msg, "推荐", "建议", "适合", "什么", "哪个", "怎么选"))
        {
            return Ok(new
            {
                success = true,
                reply = "根据不同的香调偏好，我建议您：\n• 喜欢清新花香 → 试试花香调香水\n• 喜欢温暖木质 → 木质调香水很适合\n• 喜欢清爽海洋 → 海洋调香水不会错\n\n您也可以试试我们的「香氛测试」功能，会帮您精准推荐！",
                suggestions = new[]
                {
                    new { text = "香氛测试", url = "/fragrance-quiz" },
                    new { text = "热门商品", url = "/products" }
                }
            });
        }

        return Ok(new { success = true, reply = GetRandom(_fallbacks) });
    }

    [HttpGet("faq")]
    public IActionResult GetFaq()
    {
        var faqs = new List<object>();
        foreach (var kvp in _faq)
        {
            faqs.Add(new { question = kvp.Key, answer = kvp.Value[0] });
        }
        return Ok(new { success = true, data = faqs });
    }

    private static bool ContainsAny(string text, params string[] keywords)
    {
        return keywords.Any(k => text.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetRandom(string[] arr)
    {
        return arr[Random.Shared.Next(arr.Length)];
    }
}

public class ChatbotRequest
{
    public string Message { get; set; } = "";
    public int? UserId { get; set; }
}
