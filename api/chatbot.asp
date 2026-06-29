<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/api_response.asp"-->
<!--#include file="../includes/api_guard.asp"-->
<!--#include file="../includes/ai_client.asp"-->
<%
' ============================================
' V18.0 智能客服机器人 API (AI Chatbot API)
' 三层架构: FAQ规则引擎 → AI智能回复 → 人工转接标记
' 用法: POST /api/chatbot.asp
'   body: message=如何退换货&session_id=xxx
' 返回: {code, message, data: {reply, confidence, source, handoff, session_id, suggestions}}
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

Call OpenConnection()

' V18: API 守卫（速率限制，不要求登录）
If Not API_Guard("chatbot", False) Then Response.End

Dim message, sessionId
message = Trim(Request.Form("message"))
sessionId = Trim(Request.Form("session_id"))

' 生成会话ID
If sessionId = "" Then
    Randomize
    sessionId = "cb_" & Hex(Int(Rnd * 4294967295))
End If

If message = "" Then
    Call API_Error(API_ERR_PARAM_MISSING, "请输入您的问题，我会尽力帮您解答")
    Response.End
End If

' 限制消息长度
If Len(message) > 500 Then message = Left(message, 500)

' ============================================
' Layer 1: FAQ 规则引擎（关键词匹配）
' ============================================
Dim faqReply, faqConfidence, faqCategory, faqSuggestions
faqReply = CHAT_FAQSearch(message, faqConfidence, faqCategory, faqSuggestions)

If faqReply <> "" And faqConfidence >= 0.6 Then
    Call API_Success(CHAT_BuildResponseObj(faqReply, "faq", faqConfidence, False, sessionId, faqCategory, faqSuggestions), "FAQ匹配")
    Call CloseConnection()
    Response.End
End If

' ============================================
' Layer 2: AI 智能回复
' ============================================
If FEATURE_AI_CHATBOT Then
    Dim aiResult, aiReply, aiConfidence
    aiReply = ""
    aiConfidence = 0
    
    On Error Resume Next
    Set aiResult = AI_ChatbotMessage(message, sessionId)
    If Err.Number = 0 And Not IsEmpty(aiResult) And IsObject(aiResult) Then
        If aiResult.Exists("reply") Then aiReply = CStr(aiResult("reply"))
        If aiResult.Exists("confidence") Then
            aiConfidence = CDbl(aiResult("confidence"))
        Else
            aiConfidence = 0.5
        End If
    End If
    On Error GoTo 0
    
    If aiReply <> "" Then
        Dim needHandoff, aiNote
        needHandoff = False
        aiNote = ""
        
        If aiConfidence < 0.5 Then
            needHandoff = True
            aiNote = CHAT_GetHandoffNote()
        End If
        
        Call API_Success(CHAT_BuildResponseObj(aiReply & aiNote, "ai", aiConfidence, needHandoff, sessionId, "", Empty), "AI回复")
        Call CloseConnection()
        Response.End
    End If
End If

' ============================================
' Layer 3: 兜底回复（人工转接）
' ============================================
Dim fallbackReply, fallbackSuggestions
fallbackReply = CHAT_GetFallbackReply()
fallbackSuggestions = CHAT_GetQuickReplies()

Call API_Success(CHAT_BuildResponseObj(fallbackReply, "fallback", 0, True, sessionId, "general", fallbackSuggestions), "默认回复")
Call CloseConnection()
%>
<%
' ============================================
' V18 客服FAQ规则引擎
' ============================================

' FAQ条目结构: Category|Keywords(逗号分隔)|Reply|Suggestions(逗号分隔,可选)
' 共12个FAQ条目，覆盖退换货、配送、定制、支付、会员、产品、订单等场景

Function CHAT_LoadFAQ()
    Dim faq(11, 3)  ' 12条FAQ × 4列 (Category, Keywords, Reply, Suggestions)
    
    ' 1. 退换货政策
    faq(0, 0) = "return"
    faq(0, 1) = "退换货,退货,换货,退款,退钱,不喜欢,不满意,退,换"
    faq(0, 2) = "我们支持7天无理由退换货哦！✨" & vbCrLf & vbCrLf & _
                 "📋 退换条件：商品未经使用、包装完好、不影响二次销售" & vbCrLf & _
                 "📅 期限：签收后7天内申请" & vbCrLf & _
                 "🚚 运费：质量问题我们承担，非质量问题由您承担" & vbCrLf & _
                 "💰 退款：审核通过后3-5个工作日原路退回" & vbCrLf & vbCrLf & _
                 "如需申请，请前往【我的订单】页面操作，或联系在线客服协助。"
    faq(0, 3) = "如何申请退换货,退款什么时候到账"
    
    ' 2. 定制香水退换
    faq(1, 0) = "return"
    faq(1, 1) = "定制退货,定制换货,定制退款,定制不满意"
    faq(1, 2) = "关于定制香水的退换问题：" & vbCrLf & vbCrLf & _
                 "由于定制香水是根据您的个性化选择调配的专属产品，属于「个性定制商品」。根据规定，定制商品不支持7天无理由退换。" & vbCrLf & vbCrLf & _
                 "但如遇以下情况，我们全力保障您的权益：" & vbCrLf & _
                 "🔸 商品与订单配方不符 → 免费重做或全额退款" & vbCrLf & _
                 "🔸 商品破损/泄漏 → 免费补发" & vbCrLf & vbCrLf & _
                 "建议定制前先做【香氛测试】，帮助您更准确地找到喜欢的香调组合！"
    faq(1, 3) = "定制前如何测试香调,香氛测试在哪里"
    
    ' 3. 配送物流
    faq(2, 0) = "shipping"
    faq(2, 1) = "配送,物流,快递,发货,送货,多久到,几天到,什么时候发,运费,包邮,免邮"
    faq(2, 2) = "📦 配送信息如下：" & vbCrLf & vbCrLf & _
                 "🚚 合作快递：顺丰速运 / 京东物流" & vbCrLf & _
                 "⏱ 发货时效：下单后24-48小时内发货（定制香水需额外1-2天调配）" & vbCrLf & _
                 "📍 配送时效：一线城市1-2天，其他地区2-4天" & vbCrLf & _
                 "💰 运费标准：满299元包邮，不满299元运费15元" & vbCrLf & _
                 "🔍 物流跟踪：可在【我的订单】中实时查看物流状态"
    faq(2, 3) = "如何查看物流,我的订单在哪里"
    
    ' 4. 定制流程
    faq(3, 0) = "custom"
    faq(3, 1) = "定制,DIY,调配,怎么做,怎么定制,步骤,流程,选香调"
    faq(3, 2) = "🎨 定制专属香水的步骤很简单：" & vbCrLf & vbCrLf & _
                 "第1步 → 选择基香（花香/果香/木质/东方/清新等）" & vbCrLf & _
                 "第2步 → 搭配前中后调（前调10-15分钟、中调2-4小时、后调4-8小时）" & vbCrLf & _
                 "第3步 → 选择浓度（淡香水EDT / 香水EDP / 香精Parfum）" & vbCrLf & _
                 "第4步 → 选择瓶型与容量（30ml / 50ml / 100ml）" & vbCrLf & _
                 "第5步 → 提交订单，我们的调香师为您精心调配" & vbCrLf & vbCrLf & _
                 "💡 小贴士：不确定怎么选？先试试【AI香氛测试】，只需回答6个问题，60秒找到最适合您的香调组合！" & vbCrLf & _
                 "👉 点击这里开始：<a href='/fragrance_quiz.asp'>香氛测试</a>"
    faq(3, 3) = "香氛测试怎么用,有哪些香调可以选"
    
    ' 5. 支付方式
    faq(4, 0) = "payment"
    faq(4, 1) = "支付,付款,怎么付,微信,支付宝,银行卡,信用卡,安全,扣款"
    faq(4, 2) = "💳 我们支持以下支付方式：" & vbCrLf & vbCrLf & _
                 "🔹 微信支付" & vbCrLf & _
                 "🔹 支付宝" & vbCrLf & _
                 "🔹 银行卡（借记卡/信用卡）" & vbCrLf & vbCrLf & _
                 "🔒 支付安全：所有支付均通过PCI-DSS认证的支付网关处理，我们不会存储您的银行卡信息。" & vbCrLf & _
                 "📱 支持手机端和电脑端完成支付。"
    faq(4, 3) = "支付安全吗,支持货到付款吗"
    
    ' 6. 会员与积分
    faq(5, 0) = "member"
    faq(5, 1) = "会员,积分,等级,银卡,金卡,钻石,黑金,升级,权益,优惠,折扣"
    faq(5, 2) = "👑 会员等级体系：" & vbCrLf & vbCrLf & _
                 "🥈 银卡会员（0-3000元消费）→ 享受9.5折" & vbCrLf & _
                 "🥇 金卡会员（3000-10000元）→ 享受9折 + 免运费" & vbCrLf & _
                 "💎 钻石会员（10000-30000元）→ 享受8.5折 + 免运费 + 生日礼" & vbCrLf & _
                 "🖤 黑金会员（30000元以上）→ 享受8折 + 专属客服 + 优先发货 + 限量新品优先体验" & vbCrLf & vbCrLf & _
                 "⭐ 积分规则：每消费1元=1积分，积分可在结算时抵扣（100积分=1元）" & vbCrLf & _
                 "积分有效期：滚动12个月"
    faq(5, 3) = "积分怎么用,如何升级会员"
    
    ' 7. 香调知识
    faq(6, 0) = "product"
    faq(6, 1) = "香调,前调,中调,后调,花香,果香,木质,东方,清新,柑橘,留香,持香,浓度,EDT,EDP"
    faq(6, 2) = "🌸 香水香调小知识：" & vbCrLf & vbCrLf & _
                 "香水的香气分为三个层次（金字塔结构）：" & vbCrLf & _
                 "🌅 前调（Top Notes）：喷后10-15分钟，是第一印象，通常较清新（柑橘、绿叶、薄荷）" & vbCrLf & _
                 "☀ 中调（Heart Notes）：持续2-4小时，是香水的核心（花香、果香、辛香）" & vbCrLf & _
                 "🌙 后调（Base Notes）：持续4-8小时，是香水的余韵（木质、麝香、琥珀、香草）" & vbCrLf & vbCrLf & _
                 "💧 浓度选择：" & vbCrLf & _
                 "EDT淡香水（5-15%香精）→ 清新日常，持香3-4小时" & vbCrLf & _
                 "EDP香水（15-20%香精）→ 优雅持久，持香5-8小时" & vbCrLf & _
                 "Parfum香精（20-30%香精）→ 浓郁奢华，持香8-12小时"
    faq(6, 3) = "什么香调适合夏天,什么香调适合约会"
    
    ' 8. 订单查询
    faq(7, 0) = "order"
    faq(7, 1) = "订单,查订单,我的订单,订单状态,取消订单,修改订单,改地址"
    faq(7, 2) = "📋 订单管理指南：" & vbCrLf & vbCrLf & _
                 "🔍 查看订单：登录后进入【我的订单】即可查看所有订单及状态" & vbCrLf & _
                 "❌ 取消订单：未发货的订单可直接在订单详情页取消" & vbCrLf & _
                 "📝 修改地址：未发货前可联系客服修改收货地址" & vbCrLf & _
                 "📊 订单状态说明：" & vbCrLf & _
                 "待付款 → 已付款/调配中 → 已发货 → 已完成" & vbCrLf & vbCrLf & _
                 "如有问题，请提供订单号，我们为您快速查询。"
    faq(7, 3) = "订单号在哪里看,发货后还能改地址吗"
    
    ' 9. 联系客服
    faq(8, 0) = "contact"
    faq(8, 1) = "联系,客服,电话,微信,人工,转人工,找客服,投诉,建议"
    faq(8, 2) = "📞 联系客服方式：" & vbCrLf & vbCrLf & _
                 "📱 客服电话：400-888-8888（工作日9:00-21:00）" & vbCrLf & _
                 "💬 在线客服：您正在使用的人工智能客服7×24小时在线" & vbCrLf & _
                 "📧 邮箱：contact@perfumeshop.com" & vbCrLf & _
                 "💚 微信：请搜索关注公众号「香氛定制」" & vbCrLf & vbCrLf & _
                 "如需转接人工客服，请回复「转人工」，我们将尽快为您安排。"
    faq(8, 3) = "客服几点上班,投诉在哪里"
    
    ' 10. 关于我们
    faq(9, 0) = "general"
    faq(9, 1) = "品牌,关于,公司,介绍,靠谱,正规,质量,安全,原料"
    faq(9, 2) = "🌟 关于「香氛定制」：" & vbCrLf & vbCrLf & _
                 "我们是一家专注于个性化香氛定制的品牌，成立于2020年。" & vbCrLf & vbCrLf & _
                 "✅ 所有香料均来自国际顶级香精公司（Givaudan、Firmenich、IFF）" & vbCrLf & _
                 "✅ 通过IFRA国际香精协会安全标准认证" & vbCrLf & _
                 "✅ 专业调香师团队精心调配每一瓶香水" & vbCrLf & _
                 "✅ 已服务超过100,000位热爱香氛的用户" & vbCrLf & vbCrLf & _
                 "了解更多，请访问<a href='/about.asp'>品牌故事</a>页面。"
    faq(9, 3) = "原料从哪里来,有没有实体店"
    
    ' 11. 产品适用场景
    faq(10, 0) = "product"
    faq(10, 1) = "场合,场景,上班,约会,派对,运动,夏天,冬天,季节,送人,送礼,女友,男友,礼物"
    faq(10, 2) = "🎯 场景香氛推荐：" & vbCrLf & vbCrLf & _
                 "💼 上班/商务：清新柑橘调、淡雅绿茶调 → 低调不扰人" & vbCrLf & _
                 "💕 约会/晚宴：玫瑰花香调、东方琥珀调 → 浪漫迷人" & vbCrLf & _
                 "🏃 运动/户外：海洋调、柑橘薄荷调 → 清爽活力" & vbCrLf & _
                 "☀ 夏季：清新海洋调、果香柑橘调 → 清爽宜人" & vbCrLf & _
                 "❄ 冬季：温暖木质调、琥珀香草调 → 温暖厚重" & vbCrLf & _
                 "🎁 送礼：建议使用【香氛测试】了解收礼人的偏好，或查看我们的礼品推荐！" & vbCrLf & vbCrLf & _
                 "不确定选什么？试试AI香氛测试，根据您的场合和偏好精准匹配！"
    faq(10, 3) = "送女朋友什么香水好,上班用什么香水"
    
    ' 12. 安全与隐私
    faq(11, 0) = "general"
    faq(11, 1) = "隐私,安全,信息,数据,密码,账户,泄露"
    faq(11, 2) = "🔒 您的安全与隐私是我们的首要任务：" & vbCrLf & vbCrLf & _
                 "✅ 密码使用SHA-512加密存储，任何人（包括我们的员工）都无法查看您的明文密码" & vbCrLf & _
                 "✅ 支付信息通过PCI-DSS认证网关处理，本网站不存储银行卡信息" & vbCrLf & _
                 "✅ 个人数据仅用于订单处理和产品推荐，不会出售给第三方" & vbCrLf & _
                 "✅ 您可以随时导出或删除您的个人数据（前往【账户设置】操作）" & vbCrLf & vbCrLf & _
                 "了解更多：<a href='/user/privacy.asp'>隐私政策</a>"
    faq(11, 3) = "如何删除账户,如何导出数据"
    
    CHAT_LoadFAQ = faq
End Function

' ============================================
' FAQ搜索：关键词权重匹配
' 返回: Reply文本, 置信度, 分类, 追问建议数组
' ============================================
Function CHAT_FAQSearch(keyword, ByRef confidence, ByRef category, ByRef suggestions)
    Dim faq, i, kw, kwArr, j, score, bestIdx, bestScore, reply
    Dim normKw, matchedKwCount, totalKwWeight
    
    faq = CHAT_LoadFAQ()
    normKw = LCase(keyword)
    bestIdx = -1
    bestScore = 0
    
    ' 特殊处理：转人工
    If InStr(normKw, "转人工") > 0 Or InStr(normKw, "人工客服") > 0 Or InStr(normKw, "找人工") > 0 Then
        confidence = 1.0
        category = "handoff"
        suggestions = Empty
        CHAT_FAQSearch = "正在为您转接人工客服，请稍候..." & vbCrLf & vbCrLf & _
                         "⏰ 人工客服工作时间：工作日 9:00-21:00" & vbCrLf & _
                         "📞 您也可以直接拨打电话：400-888-8888" & vbCrLf & vbCrLf & _
                         "在等待期间，您可以继续向我提问其他问题。"
        Exit Function
    End If
    
    ' 遍历FAQ条目进行关键词匹配
    For i = 0 To UBound(faq, 1)
        kwArr = Split(faq(i, 1), ",")
        score = 0
        matchedKwCount = 0
        
        For j = 0 To UBound(kwArr)
            kw = LCase(Trim(kwArr(j)))
            If kw <> "" And InStr(normKw, kw) > 0 Then
                ' 关键词越长权重越高
                score = score + Len(kw)
                matchedKwCount = matchedKwCount + 1
            End If
        Next
        
        ' 至少匹配1个关键词
        If matchedKwCount > 0 Then
            ' 归一化分值: 匹配的关键词总长度 / 条目关键词总长度
            totalKwWeight = 0
            For j = 0 To UBound(kwArr)
                totalKwWeight = totalKwWeight + Len(Trim(kwArr(j)))
            Next
            If totalKwWeight > 0 Then
                score = score / totalKwWeight
            End If
            
            If score > bestScore Then
                bestScore = score
                bestIdx = i
            End If
        End If
    Next
    
    If bestIdx >= 0 And bestScore > 0 Then
        reply = faq(bestIdx, 2)
        confidence = bestScore
        If confidence > 1.0 Then confidence = 1.0
        category = faq(bestIdx, 0)
        
        ' 构建追问建议
        If faq(bestIdx, 3) <> "" Then
            suggestions = Split(faq(bestIdx, 3), ",")
        Else
            suggestions = Empty
        End If
        
        CHAT_FAQSearch = reply
    Else
        confidence = 0
        category = ""
        suggestions = Empty
        CHAT_FAQSearch = ""
    End If
End Function

' ============================================
' 兜底回复（FAQ和AI都失败时）
' ============================================
Function CHAT_GetFallbackReply()
    CHAT_GetFallbackReply = "很抱歉，我暂时没能完全理解您的问题 😊" & vbCrLf & vbCrLf & _
                            "您可以尝试：" & vbCrLf & _
                            "🔹 换个方式描述您的问题" & vbCrLf & _
                            "🔹 点击下方快捷按钮获取常见问题解答" & vbCrLf & _
                            "🔹 回复「转人工」联系人工客服" & vbCrLf & vbCrLf & _
                            "我们的客服热线：400-888-8888（工作日9:00-21:00）"
End Function

' ============================================
' 人工转接提示语
' ============================================
Function CHAT_GetHandoffNote()
    CHAT_GetHandoffNote = vbCrLf & vbCrLf & _
                          "---" & vbCrLf & _
                          "🤔 以上回复由AI生成，如未完全解决您的问题，可以回复「转人工」联系我们的客服团队。"
End Function

' ============================================
' 快捷提问按钮列表
' ============================================
Function CHAT_GetQuickReplies()
    Dim qr(6)
    qr(0) = "如何退换货？"
    qr(1) = "配送多久能到？"
    qr(2) = "怎么定制香水？"
    qr(3) = "有哪些香调？"
    qr(4) = "会员有什么权益？"
    qr(5) = "香水怎么选浓度？"
    qr(6) = "转人工客服"
    CHAT_GetQuickReplies = qr
End Function

' ============================================
' 构建响应对象（Dictionary）
' ============================================
Function CHAT_BuildResponseObj(reply, source, confidence, handoff, sessionId, category, suggestions)
    Dim obj
    Set obj = Server.CreateObject("Scripting.Dictionary")
    obj.Add "reply", reply
    obj.Add "source", source
    obj.Add "confidence", confidence
    obj.Add "handoff", handoff
    obj.Add "session_id", sessionId
    If category <> "" Then obj.Add "category", category
    
    ' 添加追问建议
    If IsArray(suggestions) Then
        Dim sugArr, si
        ReDim sugArr(UBound(suggestions))
        For si = 0 To UBound(suggestions)
            sugArr(si) = Trim(suggestions(si))
        Next
        obj.Add "suggestions", sugArr
    ElseIf Not IsEmpty(suggestions) And Not IsNull(suggestions) Then
        obj.Add "suggestions", suggestions
    End If
    
    Set CHAT_BuildResponseObj = obj
End Function
%>
