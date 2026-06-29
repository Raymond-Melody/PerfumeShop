<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<%
If Not FEATURE_SUBSCRIPTION Then Response.Redirect "/index.asp"

Call OpenConnection()

Dim subMsg, subMsgType
subMsg = ""
subMsgType = ""

' 处理订阅
If Request.Form("action") = "subscribe" And Session("UserID") <> "" Then
    Dim newUserID : newUserID = CLng(Session("UserID"))
    Dim newPlanID : newPlanID = CLng(Request.Form("plan_id"))
    Dim autoRenew : autoRenew = IIf(Request.Form("auto_renew") = "1", 1, 0)

    ' 检查是否已有活跃订阅
    Dim existingSub
    existingSub = CLng(DAL_GetScalar("SELECT COUNT(*) FROM UserSubscriptions WHERE UserID = @UserID AND Status = 0", _
                    Array(Array("@UserID", DAL_adInteger, 0, newUserID)), 0))
    If existingSub > 0 Then
        subMsg = "您已有一个活跃的订阅，请先管理现有订阅"
        subMsgType = "error"
    Else
        ' 获取计划信息
        Dim planRow : Set planRow = DAL_GetRow("SELECT * FROM SubscriptionPlans WHERE PlanID = @PlanID AND IsActive = 1", _
                                    Array(Array("@PlanID", DAL_adInteger, 0, newPlanID)))
        If planRow Is Nothing Then
            subMsg = "所选计划不存在或已下架"
            subMsgType = "error"
        Else
            Dim planPeriod : planPeriod = planRow("Period")
            Dim nextDate : nextDate = Now()
            Select Case LCase(planPeriod)
                Case "monthly": nextDate = DateAdd("m", 1, Now())
                Case "quarterly": nextDate = DateAdd("m", 3, Now())
                Case "yearly": nextDate = DateAdd("yyyy", 1, Now())
            End Select

            Dim subID
            subID = DAL_Insert("UserSubscriptions", _
                Array("UserID", "PlanID", "Status", "StartDate", "NextDeliveryDate", "AutoRenew"), _
                Array(Array("@UserID", DAL_adInteger, 0, newUserID), _
                      Array("@PlanID", DAL_adInteger, 0, newPlanID), _
                      Array("@Status", DAL_adInteger, 0, 0), _
                      Array("@StartDate", DAL_adDate, 0, Now()), _
                      Array("@NextDeliveryDate", DAL_adDate, 0, nextDate), _
                      Array("@AutoRenew", DAL_adInteger, 0, autoRenew)))

            If subID > 0 Then
                ' 创建首期配送记录
                DAL_Execute "INSERT INTO SubscriptionDeliveries (SubscriptionID, DeliveryDate, Status) VALUES (@SubID, @Date, 0)", _
                    Array(Array("@SubID", DAL_adInteger, 0, subID), Array("@Date", DAL_adDate, 0, nextDate))

                subMsg = "订阅成功！您的首次配送将在 " & FormatDateTime(nextDate, 1) & " 左右发出"
                subMsgType = "success"
            Else
                subMsg = "订阅失败，请重试"
                subMsgType = "error"
            End If
        End If
    End If
End If

' 获取所有活跃计划
Dim rsPlans : Set rsPlans = DAL_GetList("SELECT * FROM SubscriptionPlans WHERE IsActive = 1 ORDER BY SortOrder ASC, Price ASC", Null)

' 检查用户是否有活跃订阅
Dim hasActiveSub : hasActiveSub = False
If Session("UserID") <> "" Then
    hasActiveSub = CLng(DAL_GetScalar("SELECT COUNT(*) FROM UserSubscriptions WHERE UserID = @UserID AND Status = 0", _
                    Array(Array("@UserID", DAL_adInteger, 0, CLng(Session("UserID")))), 0)) > 0
End If
%>
<!--#include file="includes/header.asp"-->

<section class="page-hero subscribe-hero">
    <div class="container">
        <div class="hero-content text-center">
            <div class="sub-icon"><i class="fas fa-box-open"></i></div>
            <h1>香氛订阅盒</h1>
            <p>每月为你甄选新香，开启香氛探索之旅</p>
        </div>
    </div>
</section>

<% If subMsg <> "" Then %>
<div class="container" style="margin-top:20px;">
    <div class="alert <%= IIf(subMsgType = "success", "alert-success", "alert-error") %>">
        <i class="fas fa-<%= IIf(subMsgType = "success", "check-circle", "info-circle") %>"></i> <%= subMsg %>
        <% If subMsgType = "success" Then %>
        <br><a href="/user/subscription.asp" style="color:inherit;font-weight:600;">前往管理订阅 &raquo;</a>
        <% End If %>
    </div>
</div>
<% End If %>

<section class="how-it-works">
    <div class="container">
        <div class="section-header text-center">
            <h2>订阅流程</h2>
            <p>简单三步，开启香氛探索之旅</p>
        </div>
        <div class="steps-row">
            <div class="hiw-step">
                <div class="hiw-number">1</div>
                <div class="hiw-icon"><i class="fas fa-clipboard-check"></i></div>
                <h4>选择计划</h4>
                <p>月/季/年三种灵活周期</p>
            </div>
            <div class="hiw-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="hiw-step">
                <div class="hiw-number">2</div>
                <div class="hiw-icon"><i class="fas fa-magic"></i></div>
                <h4>AI 选品</h4>
                <p>根据偏好智能匹配香氛</p>
            </div>
            <div class="hiw-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="hiw-step">
                <div class="hiw-number">3</div>
                <div class="hiw-icon"><i class="fas fa-truck"></i></div>
                <h4>定期配送</h4>
                <p>每月准时送到你手中</p>
            </div>
        </div>
    </div>
</section>

<section class="plans-section">
    <div class="container">
        <div class="section-header text-center">
            <h2>选择你的订阅计划</h2>
        </div>

        <% If rsPlans Is Nothing Or rsPlans.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-box-open"></i>
            <p>暂无可用订阅计划</p>
        </div>
        <% Else %>
        <div class="plans-grid">
            <%
            Do While Not rsPlans.EOF
                Dim pID, pName, pPeriod, pPrice, pSample, pFull, pFree, pCancel, pDesc
                pID = rsPlans("PlanID")
                pName = rsPlans("PlanName")
                pPeriod = rsPlans("Period")
                pPrice = rsPlans("Price")
                pSample = rsPlans("SampleCount")
                pFull = rsPlans("FullSizeCount")
                pFree = CBool(rsPlans("FreeShipping"))
                pCancel = rsPlans("CancellationFee")
                pDesc = rsPlans("Description")

                Dim pPeriodLabel, pPerDayPrice
                Select Case LCase(pPeriod)
                    Case "monthly": pPeriodLabel = "月付" : pPerDayPrice = FormatNumber(pPrice / 30, 2)
                    Case "quarterly": pPeriodLabel = "季付" : pPerDayPrice = FormatNumber(pPrice / 90, 2)
                    Case "yearly": pPeriodLabel = "年付" : pPerDayPrice = FormatNumber(pPrice / 365, 2)
                    Case Else: pPeriodLabel = pPeriod : pPerDayPrice = "—"
                End Select

                ' 判断推荐标签
                Dim recommendTag : recommendTag = ""
                If pFull >= 3 Then recommendTag = "最超值"
                If pSample >= 5 Then recommendTag = "推荐"
            %>
            <div class="plan-card<%= IIf(recommendTag <> "", " plan-recommended", "") %>">
                <% If recommendTag <> "" Then %>
                <div class="plan-badge"><%= recommendTag %></div>
                <% End If %>
                <div class="plan-header">
                    <h3><%= Server.HTMLEncode(pName) %></h3>
                    <span class="plan-period"><%= pPeriodLabel %></span>
                </div>
                <div class="plan-price">
                    <span class="price-amount">&yen;<%= FormatNumber(pPrice, 0) %></span>
                    <span class="price-per">/期</span>
                    <div class="price-daily">约 &yen;<%= pPerDayPrice %>/天</div>
                </div>
                <div class="plan-features">
                    <div class="plan-feature">
                        <i class="fas fa-gift"></i>
                        <span><strong><%= pSample %></strong> 款精选小样</span>
                    </div>
                    <div class="plan-feature">
                        <i class="fas fa-spray-can"></i>
                        <span><strong><%= pFull %></strong> 款正装香水</span>
                    </div>
                    <div class="plan-feature">
                        <i class="fas <%= IIf(pFree, "fa-check-circle plan-yes", "fa-times-circle plan-no") %>"></i>
                        <span><% If pFree Then %>免费配送<% Else %>运费另计<% End If %></span>
                    </div>
                    <div class="plan-feature">
                        <i class="fas fa-redo-alt"></i>
                        <span>随时可暂停/取消</span>
                    </div>
                    <% If CDbl(pCancel) > 0 Then %>
                    <div class="plan-feature">
                        <i class="fas fa-coins"></i>
                        <span>取消费 &yen;<%= FormatNumber(pCancel, 0) %></span>
                    </div>
                    <% End If %>
                </div>
                <div class="plan-desc"><%= Server.HTMLEncode(pDesc & "") %></div>
                <div class="plan-action">
                    <% If Session("UserID") <> "" Then %>
                        <% If hasActiveSub Then %>
                        <a href="/user/subscription.asp" class="btn btn-outline btn-block">管理现有订阅</a>
                        <% Else %>
                        <form method="post" onsubmit="return confirm('确认订阅 <%= Server.HTMLEncode(pName) %>？&yen;<%= FormatNumber(pPrice, 0) %>/期')">
                            <input type="hidden" name="action" value="subscribe">
                            <input type="hidden" name="plan_id" value="<%= pID %>">
                            <input type="hidden" name="auto_renew" value="1">
                            <button type="submit" class="btn btn-primary btn-block">立即订阅</button>
                        </form>
                        <% End If %>
                    <% Else %>
                    <a href="/user/login.asp" class="btn btn-primary btn-block">登录后订阅</a>
                    <% End If %>
                </div>
            </div>
            <%
                rsPlans.MoveNext
            Loop
            %>
        </div>
        <% End If %>
    </div>
</section>

<section class="subscribe-faq">
    <div class="container">
        <div class="section-header text-center">
            <h2>常见问题</h2>
        </div>
        <div class="faq-grid">
            <div class="faq-item">
                <h4><i class="fas fa-question-circle"></i> 可以随时取消吗？</h4>
                <p>当然！您可以随时在用户中心暂停或取消订阅，部分计划可能有取消费用。</p>
            </div>
            <div class="faq-item">
                <h4><i class="fas fa-question-circle"></i> 每期配什么产品？</h4>
                <p>我们的 AI 会根据您的香氛偏好、季节和最新产品线为您个性化选品，确保每次都是惊喜。</p>
            </div>
            <div class="faq-item">
                <h4><i class="fas fa-question-circle"></i> 可以跳过某期配送吗？</h4>
                <p>可以！在每期配送前3天，您可以选择跳过本期，不会产生费用。</p>
            </div>
            <div class="faq-item">
                <h4><i class="fas fa-question-circle"></i> 不喜欢收到的产品怎么办？</h4>
                <p>7天内支持无理由退换，正装未拆封可退，小样开封后也支持更换口味。</p>
            </div>
        </div>
    </div>
</section>

<style nonce="<%= Session("csp_nonce") %>">
.subscribe-hero {
    background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    color: #fff; padding: 60px 0 40px; text-align: center;
}
.subscribe-hero .sub-icon { font-size: 3rem; margin-bottom: 10px; }
.subscribe-hero h1 { font-size: 2.5rem; margin: 10px 0; }
.subscribe-hero p { font-size: 1.1rem; opacity: 0.9; }

.how-it-works { padding: 50px 0; background: #fafafa; }
.steps-row { display: flex; align-items: center; justify-content: center; gap: 20px; flex-wrap: wrap; }
.hiw-step { text-align: center; flex: 0 0 160px; }
.hiw-number { width: 40px; height: 40px; background: linear-gradient(135deg, #11998e, #38ef7d); color: #fff; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 10px; font-weight: 700; font-size: 18px; }
.hiw-icon { font-size: 1.8rem; color: #11998e; margin-bottom: 8px; }
.hiw-step h4 { margin: 0 0 4px; font-size: 15px; }
.hiw-step p { margin: 0; font-size: 12px; color: #999; }
.hiw-arrow { color: #ccc; font-size: 1.5rem; }

.plans-section { padding: 50px 0; }
.plans-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 24px; max-width: 1000px; margin: 0 auto; }
.plan-card { background: #fff; border-radius: 16px; padding: 32px 24px; box-shadow: 0 4px 16px rgba(0,0,0,0.08); position: relative; transition: transform 0.2s, box-shadow 0.2s; display: flex; flex-direction: column; }
.plan-card:hover { transform: translateY(-4px); box-shadow: 0 8px 24px rgba(0,0,0,0.15); }
.plan-recommended { border: 2px solid #11998e; }
.plan-badge { position: absolute; top: -12px; left: 50%; transform: translateX(-50%); background: linear-gradient(135deg, #11998e, #38ef7d); color: #fff; padding: 4px 20px; border-radius: 20px; font-size: 13px; font-weight: 700; }
.plan-header { text-align: center; margin-bottom: 16px; }
.plan-header h3 { margin: 0 0 4px; font-size: 20px; }
.plan-period { font-size: 13px; color: #999; }
.plan-price { text-align: center; margin-bottom: 20px; padding-bottom: 20px; border-bottom: 1px solid #eee; }
.price-amount { font-size: 40px; font-weight: 700; color: #11998e; }
.price-per { font-size: 16px; color: #999; }
.price-daily { font-size: 12px; color: #bbb; margin-top: 4px; }
.plan-features { margin-bottom: 16px; flex: 1; }
.plan-feature { display: flex; align-items: center; gap: 10px; padding: 8px 0; font-size: 14px; }
.plan-feature i { width: 20px; text-align: center; color: #11998e; }
.plan-yes { color: #11998e !important; }
.plan-no { color: #ccc !important; }
.plan-desc { font-size: 13px; color: #888; line-height: 1.6; margin-bottom: 20px; }
.plan-action { margin-top: auto; }

.subscribe-faq { padding: 50px 0; background: #fafafa; }
.faq-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; max-width: 900px; margin: 0 auto; }
.faq-item { background: #fff; padding: 20px; border-radius: 10px; box-shadow: 0 1px 6px rgba(0,0,0,0.06); }
.faq-item h4 { margin: 0 0 8px; font-size: 15px; color: #11998e; }
.faq-item h4 i { margin-right: 6px; }
.faq-item p { margin: 0; font-size: 13px; color: #666; line-height: 1.6; }

.empty-state { text-align: center; padding: 60px 20px; color: #999; }
.empty-state i { font-size: 3rem; margin-bottom: 15px; display: block; }
.alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 0; display: flex; align-items: center; gap: 10px; }
.alert-success { background: #d4edda; color: #155724; }
.alert-error { background: #f8d7da; color: #721c24; }
.alert i { font-size: 1.2rem; }

@media (max-width: 768px) {
    .plans-grid { grid-template-columns: 1fr; }
    .hiw-arrow { display: none; }
    .steps-row { flex-direction: column; }
    .faq-grid { grid-template-columns: 1fr; }
}
</style>

<!--#include file="includes/footer.asp"-->
<%
If Not rsPlans Is Nothing Then
    If rsPlans.State = 1 Then rsPlans.Close
    Set rsPlans = Nothing
End If
Call CloseConnection()
%>
