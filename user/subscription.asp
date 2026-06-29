<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/member_utils.asp"-->
<%
If Session("UserID") = "" Or Not FEATURE_SUBSCRIPTION Then
    Response.Redirect "login.asp"
End If

Dim userID : userID = CLng(Session("UserID"))
Call OpenConnection()

Dim subMsg, subMsgType
subMsg = ""
subMsgType = ""

' 处理操作
Dim subAction : subAction = Request.QueryString("action")
If subAction = "" Then subAction = Request.Form("action")

If subAction = "pause" Then
    Dim pauseID : pauseID = Request.QueryString("id")
    If IsNumeric(pauseID) Then
        DAL_Execute "UPDATE UserSubscriptions SET Status = 1, PauseNote = @Note, UpdatedAt = GETDATE() WHERE SubscriptionID = @ID AND UserID = @UID", _
            Array(Array("@ID", DAL_adInteger, 0, CLng(pauseID)), _
                  Array("@UID", DAL_adInteger, 0, userID), _
                  Array("@Note", DAL_adVarWChar, 200, "用户主动暂停"))
        subMsg = "已暂停订阅": subMsgType = "success"
    End If
End If

If subAction = "resume" Then
    Dim resumeID : resumeID = Request.QueryString("id")
    If IsNumeric(resumeID) Then
        Dim nextDate : nextDate = DateAdd("d", 30, Now())
        DAL_Execute "UPDATE UserSubscriptions SET Status = 0, NextDeliveryDate = @Next, PauseNote = NULL, UpdatedAt = GETDATE() WHERE SubscriptionID = @ID AND UserID = @UID AND Status = 1", _
            Array(Array("@Next", DAL_adDate, 0, nextDate), _
                  Array("@ID", DAL_adInteger, 0, CLng(resumeID)), _
                  Array("@UID", DAL_adInteger, 0, userID))
        DAL_Execute "INSERT INTO SubscriptionDeliveries (SubscriptionID, DeliveryDate, Status) VALUES (@SubID, @Date, 0)", _
            Array(Array("@SubID", DAL_adInteger, 0, CLng(resumeID)), Array("@Date", DAL_adDate, 0, nextDate))
        subMsg = "已恢复订阅": subMsgType = "success"
    End If
End If

If subAction = "cancel" Then
    Dim cancelID : cancelID = Request.QueryString("id")
    Dim cancelReason : cancelReason = Trim(Request.Form("cancel_reason"))
    If IsNumeric(cancelID) Then
        DAL_Execute "UPDATE UserSubscriptions SET Status = 2, CancelReason = @Reason, EndDate = GETDATE(), UpdatedAt = GETDATE(), AutoRenew = 0 WHERE SubscriptionID = @ID AND UserID = @UID AND Status IN (0,1)", _
            Array(Array("@ID", DAL_adInteger, 0, CLng(cancelID)), _
                  Array("@UID", DAL_adInteger, 0, userID), _
                  Array("@Reason", DAL_adVarWChar, 200, cancelReason))
        subMsg = "已取消订阅": subMsgType = "success"
    End If
End If

If subAction = "skip" Then
    Dim skipID : skipID = Request.QueryString("delivery_id")
    Dim skipSubID : skipSubID = Request.QueryString("id")
    Dim skipReason : skipReason = Trim(Request.Form("skip_reason"))
    If IsNumeric(skipID) And IsNumeric(skipSubID) Then
        DAL_Execute "UPDATE SubscriptionDeliveries SET Status = 3, SkippedAt = GETDATE(), SkipReason = @Reason WHERE DeliveryID = @DID AND SubscriptionID = @SID", _
            Array(Array("@DID", DAL_adInteger, 0, CLng(skipID)), _
                  Array("@SID", DAL_adInteger, 0, CLng(skipSubID)), _
                  Array("@Reason", DAL_adVarWChar, 200, skipReason))
        ' 生成下次配送
        Dim subRow : Set subRow = DAL_GetRow("SELECT * FROM UserSubscriptions WHERE SubscriptionID = @ID AND UserID = @UID", _
                                    Array(Array("@ID", DAL_adInteger, 0, CLng(skipSubID)), Array("@UID", DAL_adInteger, 0, userID)))
        If Not subRow Is Nothing Then
            Dim nextD : nextD = DateAdd("m", 1, Now())
            Select Case LCase(subRow("Period") & "")
                Case "quarterly": nextD = DateAdd("m", 3, Now())
                Case "yearly": nextD = DateAdd("yyyy", 1, Now())
            End Select
            DAL_Execute "UPDATE UserSubscriptions SET NextDeliveryDate = @Next, UpdatedAt = GETDATE() WHERE SubscriptionID = @ID", _
                Array(Array("@Next", DAL_adDate, 0, nextD), Array("@ID", DAL_adInteger, 0, CLng(skipSubID)))
            DAL_Execute "INSERT INTO SubscriptionDeliveries (SubscriptionID, DeliveryDate, Status) VALUES (@SubID, @Date, 0)", _
                Array(Array("@SubID", DAL_adInteger, 0, CLng(skipSubID)), Array("@Date", DAL_adDate, 0, nextD))
        End If
        subMsg = "已跳过本期配送": subMsgType = "success"
    End If
End If

' 获取用户的活跃/暂停订阅
Dim rsMySubs : Set rsMySubs = DAL_GetList("SELECT us.*, sp.PlanName, sp.Period, sp.Price, sp.SampleCount, sp.FullSizeCount, sp.FreeShipping, sp.CancellationFee " & _
    "FROM UserSubscriptions us INNER JOIN SubscriptionPlans sp ON us.PlanID = sp.PlanID " & _
    "WHERE us.UserID = @UserID AND us.Status IN (0,1) ORDER BY us.Status ASC, us.StartDate DESC", _
    Array(Array("@UserID", DAL_adInteger, 0, userID)))

' 获取已结束的订阅历史
Dim rsSubHistory : Set rsSubHistory = DAL_GetList("SELECT us.*, sp.PlanName, sp.Period FROM UserSubscriptions us INNER JOIN SubscriptionPlans sp ON us.PlanID = sp.PlanID " & _
    "WHERE us.UserID = @UserID AND us.Status IN (2,3) ORDER BY us.EndDate DESC", _
    Array(Array("@UserID", DAL_adInteger, 0, userID)))

Function StatusLabel(status)
    Select Case CInt(status)
        Case 0: StatusLabel = "活跃中"
        Case 1: StatusLabel = "已暂停"
        Case 2: StatusLabel = "已取消"
        Case 3: StatusLabel = "已过期"
        Case Else: StatusLabel = "未知"
    End Select
End Function

Function StatusBadge(status)
    Select Case CInt(status)
        Case 0: StatusBadge = "badge-active"
        Case 1: StatusBadge = "badge-paused"
        Case 2: StatusBadge = "badge-cancelled"
        Case 3: StatusBadge = "badge-expired"
        Case Else: StatusBadge = ""
    End Select
End Function

Function DeliveryStatusLabel(status)
    Select Case CInt(status)
        Case 0: DeliveryStatusLabel = "待配送"
        Case 1: DeliveryStatusLabel = "已发货"
        Case 2: DeliveryStatusLabel = "已签收"
        Case 3: DeliveryStatusLabel = "已跳过"
        Case 4: DeliveryStatusLabel = "已退回"
    End Select
End Function
%>
<!--#include file="../includes/header.asp"-->

<div class="container user-section">
    <div class="page-title-section">
        <h1><i class="fas fa-box-open"></i> 我的订阅</h1>
    </div>

    <% If subMsg <> "" Then %>
    <div class="alert <% If subMsgType = "success" Then %>alert-success<% Else %>alert-error<% End If %>">
        <i class="fas fa-<% If subMsgType = "success" Then %>check-circle<% Else %>info-circle<% End If %>"></i> <%= subMsg %>
    </div>
    <% End If %>

    <% If rsMySubs Is Nothing Or rsMySubs.EOF Then %>
    <div class="empty-state">
        <i class="fas fa-box-open"></i>
        <p>你还没有订阅哦～</p>
        <a href="/subscribe.asp" class="btn btn-primary">探索订阅计划</a>
    </div>
    <% Else %>
        <%
        ' 活跃订阅详情
        Do While Not rsMySubs.EOF
            Dim mySubID, myPlanID, myStatus, myStart, myNext, myEnd, myCount, myAuto, myPause, myCancel, myPlanName, myPeriod, myPrice, mySample, myFull, myFreeShip, myCancelFee
            mySubID = rsMySubs("SubscriptionID")
            myPlanID = rsMySubs("PlanID")
            myStatus = CInt(rsMySubs("Status"))
            myStart = rsMySubs("StartDate")
            myNext = rsMySubs("NextDeliveryDate")
            myEnd = rsMySubs("EndDate")
            myCount = rsMySubs("TotalDeliveries")
            myAuto = CBool(rsMySubs("AutoRenew"))
            myPause = rsMySubs("PauseNote")
            myCancel = rsMySubs("CancelReason")
            myPlanName = rsMySubs("PlanName")
            myPeriod = rsMySubs("Period")
            myPrice = rsMySubs("Price")
            mySample = rsMySubs("SampleCount")
            myFull = rsMySubs("FullSizeCount")
            myFreeShip = CBool(rsMySubs("FreeShipping"))
            myCancelFee = rsMySubs("CancellationFee")

            Dim myPerLabel
            Select Case LCase(myPeriod)
                Case "monthly": myPerLabel = "月度"
                Case "quarterly": myPerLabel = "季度"
                Case "yearly": myPerLabel = "年度"
                Case Else: myPerLabel = myPeriod
            End Select
        %>
        <div class="sub-card">
            <div class="sub-card-header">
                <div class="sub-card-title">
                    <h3><%= Server.HTMLEncode(myPlanName) %></h3>
                    <span class="badge <%= StatusBadge(myStatus) %>"><%= StatusLabel(myStatus) %></span>
                </div>
                <div class="sub-price">&yen;<%= FormatNumber(myPrice, 0) %><span>/期</span></div>
            </div>

            <div class="sub-meta">
                <div class="meta-item">
                    <span class="meta-label">周期</span>
                    <span class="meta-value"><%= myPerLabel %></span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">每期内容</span>
                    <span class="meta-value"><%= mySample %> 款小样 + <%= myFull %> 款正装</span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">开始日期</span>
                    <span class="meta-value"><%= FormatDateTime(myStart, 1) %></span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">下次配送</span>
                    <span class="meta-value" style="color:#11998e;font-weight:600;"><%= FormatDateTime(myNext, 1) %></span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">已完成配送</span>
                    <span class="meta-value"><%= myCount %> 次</span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">配送</span>
                    <span class="meta-value"><% If myFreeShip Then %>包邮<% Else %>运费另计<% End If %></span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">自动续费</span>
                    <span class="meta-value"><% If myAuto Then %><i class="fas fa-check-circle" style="color:#11998e;"></i> 开启<% Else %><i class="fas fa-times-circle" style="color:#ccc;"></i> 关闭<% End If %></span>
                </div>
                <% If myStatus = 1 And Not IsNull(myPause) And myPause <> "" Then %>
                <div class="meta-item meta-full">
                    <span class="meta-label">暂停原因</span>
                    <span class="meta-value"><%= Server.HTMLEncode(myPause) %></span>
                </div>
                <% End If %>
                <% If myStatus = 2 And Not IsNull(myCancel) And myCancel <> "" Then %>
                <div class="meta-item meta-full">
                    <span class="meta-label">取消原因</span>
                    <span class="meta-value"><%= Server.HTMLEncode(myCancel) %></span>
                </div>
                <% End If %>
            </div>

            <!-- 配送历史 -->
            <%
            Dim rsDeliveries : Set rsDeliveries = DAL_GetList("SELECT * FROM SubscriptionDeliveries WHERE SubscriptionID = @SubID ORDER BY DeliveryDate DESC", _
                                                    Array(Array("@SubID", DAL_adInteger, 0, mySubID)))
            If Not rsDeliveries Is Nothing And Not rsDeliveries.EOF Then
            %>
            <div class="delivery-section">
                <h4><i class="fas fa-history"></i> 配送记录</h4>
                <table class="delivery-table">
                    <thead><tr><th>日期</th><th>状态</th><th>内容</th><th>物流</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Do While Not rsDeliveries.EOF
                        Dim dID, dDate, dStatus, dContents, dTracking, dShip, dDeliver, dSkipReason
                        dID = rsDeliveries("DeliveryID")
                        dDate = rsDeliveries("DeliveryDate")
                        dStatus = CInt(rsDeliveries("Status"))
                        dContents = rsDeliveries("Contents")
                        dTracking = rsDeliveries("TrackingNumber")
                        dShip = rsDeliveries("ShippedAt")
                        dDeliver = rsDeliveries("DeliveredAt")
                        dSkipReason = rsDeliveries("SkipReason")
                    %>
                    <tr>
                        <td><%= FormatDateTime(dDate, 1) %></td>
                        <td><span class="badge badge-d<%= dStatus %>"><%= DeliveryStatusLabel(dStatus) %></span></td>
                        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;"><%= IIf(IsNull(dContents) Or dContents = "", "—", Server.HTMLEncode(dContents)) %></td>
                        <td>
                            <% If Not IsNull(dTracking) And dTracking <> "" Then %>
                            <%= Server.HTMLEncode(dTracking) %>
                            <% Else %>—<% End If %>
                        </td>
                        <td>
                            <% If dStatus = 0 And myStatus = 0 Then %>
                            <button class="btn btn-sm btn-outline" onclick="document.getElementById('skipForm_<%= dID %>').style.display='block'">
                                <i class="fas fa-forward"></i> 跳过
                            </button>
                            <form method="post" id="skipForm_<%= dID %>" style="display:none;margin-top:8px;" onsubmit="return confirm('确认跳过本期配送？')">
                                <input type="hidden" name="action" value="skip">
                                <input type="hidden" name="delivery_id" value="<%= dID %>">
                                <input type="hidden" name="id" value="<%= mySubID %>">
                                <input type="text" name="skip_reason" placeholder="跳过原因（选填）" style="width:100%;margin-bottom:4px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;">
                                <button type="submit" class="btn btn-sm btn-primary">确认跳过</button>
                                <button type="button" class="btn btn-sm btn-outline" onclick="this.parentElement.style.display='none'">取消</button>
                            </form>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                        rsDeliveries.MoveNext
                    Loop
                    %>
                    </tbody>
                </table>
            </div>
            <%
            End If
            If Not rsDeliveries Is Nothing Then
                If rsDeliveries.State = 1 Then rsDeliveries.Close
                Set rsDeliveries = Nothing
            End If
            %>

            <!-- 操作按钮 -->
            <div class="sub-actions">
                <% If myStatus = 0 Then %>
                <a href="?action=pause&id=<%= mySubID %>" class="btn btn-outline" onclick="return confirm('确认暂停订阅？')">
                    <i class="fas fa-pause-circle"></i> 暂停订阅
                </a>
                <button class="btn btn-outline" onclick="document.getElementById('cancelForm_<%= mySubID %>').style.display='block'">
                    <i class="fas fa-times-circle"></i> 取消订阅
                </button>
                <div id="cancelForm_<%= mySubID %>" style="display:none;width:100%;margin-top:8px;">
                    <form method="post" class="inline-form" onsubmit="return confirm('确认取消订阅？<% If CDbl(myCancelFee) > 0 Then %>取消费用 &yen;<%= FormatNumber(myCancelFee, 0) %><% End If %>')">
                        <input type="hidden" name="action" value="cancel">
                        <input type="hidden" name="id" value="<%= mySubID %>">
                        <input type="text" name="cancel_reason" placeholder="取消原因（选填）" style="padding:6px 10px;border:1px solid #ddd;border-radius:4px;min-width:200px;">
                        <button type="submit" class="btn btn-danger">确认取消</button>
                        <button type="button" class="btn btn-outline" onclick="this.parentElement.parentElement.style.display='none'">我再想想</button>
                    </form>
                    <% If CDbl(myCancelFee) > 0 Then %>
                    <div style="margin-top:4px;font-size:12px;color:#999;">取消费用: &yen;<%= FormatNumber(myCancelFee, 2) %></div>
                    <% End If %>
                </div>
                <% ElseIf myStatus = 1 Then %>
                <a href="?action=resume&id=<%= mySubID %>" class="btn btn-primary" onclick="return confirm('确认恢复订阅？')">
                    <i class="fas fa-play-circle"></i> 恢复订阅
                </a>
                <a href="?action=cancel&id=<%= mySubID %>" class="btn btn-outline" onclick="return confirm('确认取消订阅？')">
                    <i class="fas fa-times-circle"></i> 取消订阅
                </a>
                <% ElseIf myStatus = 2 Then %>
                <a href="/subscribe.asp" class="btn btn-primary">
                    <i class="fas fa-redo"></i> 重新订阅
                </a>
                <% End If %>
            </div>
        </div>
        <%
            rsMySubs.MoveNext
        Loop
        %>
    <% End If %>

    <!-- 历史订阅 -->
    <% If Not rsSubHistory Is Nothing And Not rsSubHistory.EOF Then %>
    <div class="sub-card" style="margin-top:30px;">
        <h4><i class="fas fa-history"></i> 历史订阅</h4>
        <table class="delivery-table">
            <thead><tr><th>计划</th><th>周期</th><th>开始</th><th>结束</th><th>配送次数</th><th>状态</th></tr></thead>
            <tbody>
            <%
            Do While Not rsSubHistory.EOF
                Dim hID, hPlan, hPer, hStart, hEnd, hCount, hStatus
                hID = rsSubHistory("SubscriptionID")
                hPlan = rsSubHistory("PlanName")
                hPer = rsSubHistory("Period")
                hStart = rsSubHistory("StartDate")
                hEnd = rsSubHistory("EndDate")
                hCount = rsSubHistory("TotalDeliveries")
                hStatus = CInt(rsSubHistory("Status"))
            %>
            <tr>
                <td><%= Server.HTMLEncode(hPlan) %></td>
                <td><%= myPerLabel %></td>
                <td><%= FormatDateTime(hStart, 1) %></td>
                <td><%= FormatDateTime(hEnd, 1) %></td>
                <td><%= hCount %> 次</td>
                <td><span class="badge <%= StatusBadge(hStatus) %>"><%= StatusLabel(hStatus) %></span></td>
            </tr>
            <%
                rsSubHistory.MoveNext
            Loop
            %>
            </tbody>
        </table>
    </div>
    <% End If %>
</div>

<style>
.user-section { max-width: 800px; margin: 40px auto; padding: 0 20px; }
.page-title-section { margin-bottom: 24px; }
.page-title-section h1 { font-size: 1.6rem; margin: 0; display: flex; align-items: center; gap: 10px; color: #333; }

.sub-card { background: #fff; border-radius: 14px; padding: 28px; box-shadow: 0 2px 12px rgba(0,0,0,0.06); margin-bottom: 20px; }
.sub-card-header { display: flex; justify-content: space-between; align-items: center; padding-bottom: 20px; border-bottom: 1px solid #f0f0f0; margin-bottom: 20px; }
.sub-card-title h3 { margin: 0 0 6px; font-size: 1.1rem; display: flex; align-items: center; gap: 8px; }
.sub-price { font-size: 1.6rem; font-weight: 700; color: #11998e; }
.sub-price span { font-size: 0.8rem; color: #999; font-weight: 400; }

.sub-meta { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 20px; }
.meta-item { padding: 10px 0; }
.meta-full { grid-column: 1 / -1; }
.meta-label { display: block; font-size: 11px; color: #999; margin-bottom: 3px; text-transform: uppercase; }
.meta-value { display: block; font-size: 14px; color: #333; }

.sub-actions { display: flex; gap: 10px; margin-top: 20px; flex-wrap: wrap; padding-top: 20px; border-top: 1px solid #f0f0f0; }
.inline-form { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }

.delivery-section { margin-top: 20px; }
.delivery-section h4 { font-size: 1rem; margin: 0 0 12px; color: #555; display: flex; align-items: center; gap: 8px; }
.delivery-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.delivery-table th { text-align: left; padding: 8px 12px; background: #f5f5f5; color: #666; font-weight: 500; }
.delivery-table td { padding: 8px 12px; border-bottom: 1px solid #f0f0f0; }
.delivery-table tr:hover td { background: #fafafa; }

.badge { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 11px; font-weight: 500; }
.badge-active { background: #e8f5e9; color: #2e7d32; }
.badge-paused { background: #fff3e0; color: #e65100; }
.badge-cancelled { background: #fce4ec; color: #c62828; }
.badge-expired { background: #f5f5f5; color: #999; }
.badge-d0 { background: #e3f2fd; color: #1565c0; }
.badge-d1 { background: #f3e5f5; color: #6a1b9a; }
.badge-d2 { background: #e8f5e9; color: #2e7d32; }
.badge-d3 { background: #fafafa; color: #bbb; }
.badge-d4 { background: #fce4ec; color: #c62828; }

.btn { padding: 8px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; background: none; }
.btn-primary { background: linear-gradient(135deg, #11998e, #38ef7d); color: #fff; }
.btn-outline { background: #fff; border: 1px solid #ddd; color: #555; }
.btn-danger { background: #c62828; color: #fff; }
.btn-sm { padding: 4px 10px; font-size: 11px; }
.btn-block { display: block; width: 100%; justify-content: center; }

.empty-state { text-align: center; padding: 80px 20px; }
.empty-state i { font-size: 4rem; color: #ddd; display: block; margin-bottom: 15px; }
.empty-state p { color: #999; margin-bottom: 20px; }

.alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
.alert-success { background: #d4edda; color: #155724; }
.alert-error { background: #f8d7da; color: #721c24; }

@media (max-width: 600px) {
    .sub-meta { grid-template-columns: 1fr 1fr; }
    .sub-card-header { flex-direction: column; align-items: flex-start; gap: 12px; }
    .sub-actions { flex-direction: column; }
    .inline-form { flex-direction: column; align-items: stretch; }
}
</style>

<!--#include file="../includes/footer.asp"-->
<%
If Not rsMySubs Is Nothing Then
    If rsMySubs.State = 1 Then rsMySubs.Close
    Set rsMySubs = Nothing
End If
If Not rsSubHistory Is Nothing Then
    If rsSubHistory.State = 1 Then rsSubHistory.Close
    Set rsSubHistory = Nothing
End If
Call CloseConnection()
%>
