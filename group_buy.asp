<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<%
If Not FEATURE_GROUP_BUY Then Response.Redirect "/index.asp"

Call OpenConnection()

Dim gbPage, gbPageSize, gbPageInfo
gbPage = CLng(GetQParam("page", "1"))
gbPageSize = 12
If gbPage < 1 Then gbPage = 1

' ============================================
' 拼团引擎辅助函数
' ============================================
Function GB_GetActivePlans()
    Dim sql
    sql = "SELECT gp.PlanID, gp.ProductID, gp.TeamSize, gp.GroupPrice, " & _
          "gp.StartTime, gp.EndTime, gp.DurationHours, " & _
          "p.ProductName, p.ImageURL, p.BasePrice, p.Description " & _
          "FROM GroupBuyPlans gp " & _
          "INNER JOIN Products p ON gp.ProductID = p.ProductID " & _
          "WHERE gp.IsActive = 1 AND GETDATE() >= gp.StartTime AND GETDATE() <= gp.EndTime " & _
          "ORDER BY gp.SortOrder ASC, gp.EndTime ASC"
    Set GB_GetActivePlans = DAL_GetList(sql, Null)
End Function

Function GB_GetOpenGroups(planID)
    ' 获取该计划下正在进行中的团（可加入的）
    Dim sql
    sql = "SELECT g.GroupID, g.GroupSN, g.CurrentSize, g.TargetSize, g.CreatedAt, " & _
          "u.Username, " & _
          "DATEDIFF(hour, g.CreatedAt, GETDATE()) AS HoursPassed " & _
          "FROM GroupBuyOrders g " & _
          "LEFT JOIN Users u ON g.InitiatorID = u.UserID " & _
          "WHERE g.PlanID = @PlanID AND g.Status = 0 " & _
          "ORDER BY g.CurrentSize DESC, g.CreatedAt ASC"
    Set GB_GetOpenGroups = DAL_GetList(sql, Array(Array("@PlanID", DAL_adInteger, 0, planID)))
End Function

Function GB_GetPlanStats(planID)
    Dim dict, sql
    Set dict = Server.CreateObject("Scripting.Dictionary")
    ' 正在进行中的团数
    sql = "SELECT COUNT(*) FROM GroupBuyOrders WHERE PlanID = @PlanID AND Status = 0"
    dict.Add "openGroups", CLng(DAL_GetScalar(sql, Array(Array("@PlanID", DAL_adInteger, 0, planID)), 0))
    ' 已成团数
    sql = "SELECT COUNT(*) FROM GroupBuyOrders WHERE PlanID = @PlanID AND Status = 1"
    dict.Add "successGroups", CLng(DAL_GetScalar(sql, Array(Array("@PlanID", DAL_adInteger, 0, planID)), 0))
    ' 参团人数
    sql = "SELECT COUNT(*) FROM GroupBuyParticipants gp2 " & _
          "INNER JOIN GroupBuyOrders g ON gp2.GroupID = g.GroupID " & _
          "WHERE g.PlanID = @PlanID"
    dict.Add "totalParticipants", CLng(DAL_GetScalar(sql, Array(Array("@PlanID", DAL_adInteger, 0, planID)), 0))
    Set GB_GetPlanStats = dict
End Function

Function GB_CanUserJoinGroup(groupID, userID)
    ' 检查用户是否已在该团中
    Dim sql, cnt
    sql = "SELECT COUNT(*) FROM GroupBuyParticipants WHERE GroupID = @GroupID AND UserID = @UserID"
    cnt = CLng(DAL_GetScalar(sql, Array(Array("@GroupID", DAL_adInteger, 0, groupID), Array("@UserID", DAL_adInteger, 0, userID)), 0))
    GB_CanUserJoinGroup = (cnt = 0)
End Function

Function GB_JoinGroup(groupID, userID)
    ' 原子操作：仅在团未满时加入
    Dim sql, rowsAffected
    sql = "UPDATE GroupBuyOrders SET CurrentSize = CurrentSize + 1 " & _
          "WHERE GroupID = @GroupID AND Status = 0 AND CurrentSize < TargetSize"
    rowsAffected = DAL_Execute(sql, Array(Array("@GroupID", DAL_adInteger, 0, groupID)))
    
    If rowsAffected <= 0 Then
        GB_JoinGroup = "full"
        Exit Function
    End If
    
    ' 插入参团记录
    DAL_Execute "INSERT INTO GroupBuyParticipants (GroupID, UserID, IsInitiator, Status) VALUES (@GroupID, @UserID, 0, 0)", _
                Array(Array("@GroupID", DAL_adInteger, 0, groupID), Array("@UserID", DAL_adInteger, 0, userID))
    
    ' 检查是否已满团
    sql = "SELECT CurrentSize, TargetSize FROM GroupBuyOrders WHERE GroupID = @GroupID"
    Dim row : Set row = DAL_GetRow(sql, Array(Array("@GroupID", DAL_adInteger, 0, groupID)))
    If Not row Is Nothing Then
        If CLng(row("CurrentSize")) >= CLng(row("TargetSize")) Then
            ' 成团成功
            DAL_Execute "UPDATE GroupBuyOrders SET Status = 1, CompletedAt = GETDATE() WHERE GroupID = @GroupID", _
                        Array(Array("@GroupID", DAL_adInteger, 0, groupID))
            GB_JoinGroup = "success"
        Else
            GB_JoinGroup = "joined"
        End If
    Else
        GB_JoinGroup = "error"
    End If
End Function

Function GB_CreateGroup(planID, userID)
    ' 创建新团
    ' 先检查该计划是否还有效
    Dim sql, planRow
    sql = "SELECT gp.* FROM GroupBuyPlans gp WHERE gp.PlanID = @PlanID AND gp.IsActive = 1 " & _
          "AND GETDATE() >= gp.StartTime AND GETDATE() <= gp.EndTime"
    Set planRow = DAL_GetRow(sql, Array(Array("@PlanID", DAL_adInteger, 0, planID)))
    If planRow Is Nothing Then
        GB_CreateGroup = "invalid_plan"
        Exit Function
    End If
    
    ' 检查该用户是否已在同一计划的其他进行中的团中
    sql = "SELECT COUNT(*) FROM GroupBuyParticipants gp2 " & _
          "INNER JOIN GroupBuyOrders g ON gp2.GroupID = g.GroupID " & _
          "WHERE g.PlanID = @PlanID AND gp2.UserID = @UserID AND g.Status = 0"
    Dim existing : existing = CLng(DAL_GetScalar(sql, Array(Array("@PlanID", DAL_adInteger, 0, planID), Array("@UserID", DAL_adInteger, 0, userID)), 0))
    If existing > 0 Then
        GB_CreateGroup = "already_in"
        Exit Function
    End If
    
    ' 生成团编号
    Dim nowVal : nowVal = Now()
    Dim groupSN : groupSN = "GB" & Year(nowVal) & Right("0" & Month(nowVal), 2) & Right("0" & Day(nowVal), 2) & _
                              Right("0" & Hour(nowVal), 2) & Right("0" & Minute(nowVal), 2) & Right("0" & Second(nowVal), 2) & _
                              Right("000" & CStr(Int(Rnd * 1000)), 3)
    
    Dim teamSize : teamSize = CLng(planRow("TeamSize"))
    Dim groupID
    
    sql = "INSERT INTO GroupBuyOrders (PlanID, GroupSN, InitiatorID, Status, CurrentSize, TargetSize) " & _
          "VALUES (@PlanID, @GroupSN, @UserID, 0, 1, @TargetSize); SELECT SCOPE_IDENTITY();"
    groupID = CLng(DAL_GetScalar(sql, Array( _
        Array("@PlanID", DAL_adInteger, 0, planID), _
        Array("@GroupSN", DAL_adVarChar, 20, groupSN), _
        Array("@UserID", DAL_adInteger, 0, userID), _
        Array("@TargetSize", DAL_adInteger, 0, teamSize) _
    ), 0))
    
    If groupID <= 0 Then
        GB_CreateGroup = "error"
        Exit Function
    End If
    
    ' 插入团长记录
    DAL_Execute "INSERT INTO GroupBuyParticipants (GroupID, UserID, IsInitiator, Status) VALUES (@GroupID, @UserID, 1, 0)", _
                Array(Array("@GroupID", DAL_adInteger, 0, groupID), Array("@UserID", DAL_adInteger, 0, userID))
    
    ' 如果是2人团(一人成团)直接自动成功
    If teamSize <= 1 Then
        DAL_Execute "UPDATE GroupBuyOrders SET Status = 1, CompletedAt = GETDATE() WHERE GroupID = @GroupID", _
                    Array(Array("@GroupID", DAL_adInteger, 0, groupID))
        GB_CreateGroup = "success"
    Else
        GB_CreateGroup = "created"
    End If
End Function

Function GetQParam(name, defaultVal)
    Dim v : v = Request.QueryString(name)
    If v = "" Or IsNull(v) Then v = defaultVal
    GetQParam = v
End Function

' ============================================
' 处理 POST 操作
' ============================================
Dim gbAction : gbAction = Request.Form("action")
Dim gbResult, gbResultMsg

If gbAction <> "" And Session("UserID") <> "" Then
    Dim gbUserID : gbUserID = CLng(Session("UserID"))
    
    Select Case gbAction
        Case "create_group"
            Dim gbPlanID : gbPlanID = CLng(Request.Form("plan_id"))
            gbResult = GB_CreateGroup(gbPlanID, gbUserID)
            Select Case gbResult
                Case "created": gbResultMsg = "开团成功！快去邀请好友加入吧"
                Case "success": gbResultMsg = "恭喜，拼团成功！请前往结算"
                Case "already_in": gbResultMsg = "您已在同一拼团活动中，不能重复开团"
                Case "invalid_plan": gbResultMsg = "该拼团活动已失效"
                Case Else: gbResultMsg = "开团失败，请重试"
            End Select
            
        Case "join_group"
            Dim gbGroupID : gbGroupID = CLng(Request.Form("group_id"))
            ' 检查是否已在团中
            If Not GB_CanUserJoinGroup(gbGroupID, gbUserID) Then
                gbResult = "already_in"
                gbResultMsg = "您已在该拼团中"
            Else
                gbResult = GB_JoinGroup(gbGroupID, gbUserID)
                Select Case gbResult
                    Case "joined": gbResultMsg = "参团成功！等待更多伙伴加入"
                    Case "success": gbResultMsg = "恭喜，拼团成功！请前往结算"
                    Case "full": gbResultMsg = "该团已满员，试试其他团或自己开团"
                    Case Else: gbResultMsg = "参团失败，请重试"
                End Select
            End If
    End Select
End If

' 获取活动计划列表
Dim rsPlans
Set rsPlans = GB_GetActivePlans()
%>
<!--#include file="includes/header.asp"-->

<section class="page-hero group-buy-hero">
    <div class="container">
        <div class="hero-content text-center">
            <div class="gb-icon"><i class="fas fa-users"></i></div>
            <h1>拼团惠购</h1>
            <p>邀请好友一起拼，享受超低团购价</p>
        </div>
    </div>
</section>

<% If gbResultMsg <> "" Then %>
<div class="container" style="margin-top:20px;">
    <div class="alert <%= IIf(gbResult = "success" Or gbResult = "created" Or gbResult = "joined", "alert-success", "alert-warning") %>">
        <i class="fas fa-<%= IIf(gbResult = "success" Or gbResult = "created" Or gbResult = "joined", "check-circle", "info-circle") %>"></i>
        <%= gbResultMsg %>
        <% If gbResult = "created" Then %>
        <br><small>分享链接给好友，邀请他们一起拼团</small>
        <% End If %>
    </div>
</div>
<% End If %>

<section class="group-buy-section">
    <div class="container">
        <% If rsPlans Is Nothing Or rsPlans.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-users"></i>
            <p>当前没有正在进行的拼团活动</p>
        </div>
        <% Else %>
        <div class="gb-plan-list">
            <%
            Do While Not rsPlans.EOF
                Dim pID, pProductID, pTeamSize, pPrice, pStart, pEnd, pDuration
                Dim pProductName, pImage, pBasePrice, pDesc
                pID = rsPlans("PlanID")
                pProductID = rsPlans("ProductID")
                pTeamSize = rsPlans("TeamSize")
                pPrice = rsPlans("GroupPrice")
                pStart = rsPlans("StartTime")
                pEnd = rsPlans("EndTime")
                pDuration = rsPlans("DurationHours")
                pProductName = rsPlans("ProductName")
                pImage = rsPlans("ImageURL")
                pBasePrice = rsPlans("BasePrice")
                pDesc = rsPlans("Description")
                
                If IsNull(pImage) Or pImage = "" Then pImage = DEFAULT_PRODUCT_IMAGE
                
                Dim pDiscount : pDiscount = 0
                If CDbl(pBasePrice) > 0 Then pDiscount = Int((1 - CDbl(pPrice) / CDbl(pBasePrice)) * 100)
                
                Dim planStats : Set planStats = GB_GetPlanStats(pID)
                Dim openGroups : openGroups = CLng(planStats("openGroups"))
                Dim successGroups : successGroups = CLng(planStats("successGroups"))
                Dim totalParticipants : totalParticipants = CLng(planStats("totalParticipants"))
            %>
            <div class="gb-plan-card">
                <div class="gb-plan-image">
                    <img src="<%= pImage %>" alt="<%= Server.HTMLEncode(pProductName) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                    <div class="gb-team-badge"><%= pTeamSize %>人团</div>
                    <% If pDiscount >= 20 Then %>
                    <div class="gb-discount-tag">-<%= pDiscount %>%</div>
                    <% End If %>
                </div>
                <div class="gb-plan-body">
                    <h3><%= Server.HTMLEncode(pProductName) %></h3>
                    <p class="gb-plan-desc"><%= Server.HTMLEncode(Left(pDesc & "", 60)) %></p>
                    
                    <div class="gb-price-row">
                        <div class="gb-price-block">
                            <span class="gb-label">拼团价</span>
                            <span class="gb-price">&yen;<%= FormatNumber(pPrice, 2) %></span>
                        </div>
                        <div class="gb-price-block">
                            <span class="gb-label">原价</span>
                            <span class="gb-original">&yen;<%= FormatNumber(pBasePrice, 2) %></span>
                        </div>
                        <div class="gb-price-block">
                            <span class="gb-label">省</span>
                            <span class="gb-save">&yen;<%= FormatNumber(CDbl(pBasePrice) - CDbl(pPrice), 2) %></span>
                        </div>
                    </div>
                    
                    <div class="gb-stats">
                        <div class="gb-stat">
                            <span class="gb-stat-num"><%= openGroups %></span>
                            <span class="gb-stat-label">进行中</span>
                        </div>
                        <div class="gb-stat">
                            <span class="gb-stat-num"><%= totalParticipants %></span>
                            <span class="gb-stat-label">已参团</span>
                        </div>
                        <div class="gb-stat">
                            <span class="gb-stat-num"><%= successGroups %></span>
                            <span class="gb-stat-label">已成团</span>
                        </div>
                    </div>
                    
                    <div class="gb-time-info">
                        <i class="fas fa-clock"></i>
                        <span><%= FormatDateTime(pEnd, 0) %> 截止</span>
                        <span class="gb-duration">成团有效 <%= pDuration %> 小时</span>
                    </div>
                    
                    <div class="gb-actions">
                        <% If Session("UserID") <> "" Then %>
                        <form method="post" action="group_buy.asp" class="gb-form" onsubmit="return confirm('确定要发起 <%= pTeamSize %> 人拼团吗？')">
                            <input type="hidden" name="action" value="create_group">
                            <input type="hidden" name="plan_id" value="<%= pID %>">
                            <button type="submit" class="btn btn-primary gb-btn-start">
                                <i class="fas fa-user-plus"></i> 发起拼团
                            </button>
                        </form>
                        <% Else %>
                        <a href="/user/login.asp" class="btn btn-primary gb-btn-start">
                            <i class="fas fa-user-plus"></i> 登录后开团
                        </a>
                        <% End If %>
                        <button class="btn btn-outline gb-btn-open" onclick="toggleGroups(<%= pID %>)">
                            <i class="fas fa-list"></i> 查看可参团 (<%= openGroups %>)
                        </button>
                    </div>
                    
                    <!-- 可参团列表（默认隐藏） -->
                    <div class="gb-open-groups" id="groups_<%= pID %>" style="display:none;">
                        <%
                        Dim rsGroups : Set rsGroups = GB_GetOpenGroups(pID)
                        If Not rsGroups Is Nothing And Not rsGroups.EOF Then
                        %>
                        <table class="gb-table">
                            <thead>
                                <tr>
                                    <th>团编号</th>
                                    <th>团长</th>
                                    <th>进度</th>
                                    <th>剩余时间</th>
                                    <th>操作</th>
                                </tr>
                            </thead>
                            <tbody>
                            <%
                            Do While Not rsGroups.EOF
                                Dim gID, gSN, gCur, gTar, gInitiator, gCreated, gHours
                                gID = rsGroups("GroupID")
                                gSN = rsGroups("GroupSN")
                                gCur = rsGroups("CurrentSize")
                                gTar = rsGroups("TargetSize")
                                gInitiator = rsGroups("Username")
                                gCreated = rsGroups("CreatedAt")
                                gHours = rsGroups("HoursPassed")
                                
                                Dim remainSlots : remainSlots = gTar - gCur
                                Dim gProgress : gProgress = Int((gCur / gTar) * 100)
                            %>
                                <tr>
                                    <td><code><%= gSN %></code></td>
                                    <td><%= Server.HTMLEncode(gInitiator) %></td>
                                    <td>
                                        <div class="gb-mini-progress">
                                            <div class="gb-mini-bar" style="width:<%= gProgress %>%"></div>
                                        </div>
                                        <small><%= gCur %>/<%= gTar %> 人</small>
                                    </td>
                                    <td><small>差 <%= remainSlots %> 人</small></td>
                                    <td>
                                        <% If Session("UserID") <> "" Then %>
                                        <form method="post" action="group_buy.asp" class="gb-form-inline" onsubmit="return confirm('确定要加入该拼团吗？')">
                                            <input type="hidden" name="action" value="join_group">
                                            <input type="hidden" name="group_id" value="<%= gID %>">
                                            <button type="submit" class="btn btn-sm btn-success">加入</button>
                                        </form>
                                        <% Else %>
                                        <a href="/user/login.asp" class="btn btn-sm btn-success">登录加入</a>
                                        <% End If %>
                                    </td>
                                </tr>
                            <%
                                rsGroups.MoveNext
                            Loop
                            %>
                            </tbody>
                        </table>
                        <%
                        Else
                        %>
                        <p class="gb-no-groups">暂无进行中的团，快去发起一个吧！</p>
                        <%
                        End If
                        If Not rsGroups Is Nothing Then rsGroups.Close: Set rsGroups = Nothing
                        %>
                    </div>
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

<style nonce="<%= Session("csp_nonce") %>">
.group-buy-hero {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff; padding: 60px 0 40px; text-align: center;
}
.group-buy-hero .gb-icon { font-size: 3rem; margin-bottom: 10px; }
.group-buy-hero h1 { font-size: 2.5rem; margin: 10px 0; }
.group-buy-hero p { font-size: 1.1rem; opacity: 0.9; }

.gb-plan-list { display: flex; flex-direction: column; gap: 24px; }

.gb-plan-card {
    background: #fff; border-radius: 12px; overflow: hidden;
    display: flex; box-shadow: 0 2px 12px rgba(0,0,0,0.08);
}
.gb-plan-image {
    width: 280px; flex-shrink: 0; position: relative; overflow: hidden; background: #f9f9f9;
}
.gb-plan-image img { width: 100%; height: 100%; object-fit: cover; }
.gb-team-badge {
    position: absolute; top: 10px; left: 10px;
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: #fff; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 700;
}
.gb-discount-tag {
    position: absolute; top: 10px; right: 10px;
    background: #e74c3c; color: #fff; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 700;
}
.gb-plan-body { flex: 1; padding: 20px; display: flex; flex-direction: column; }
.gb-plan-body h3 { margin: 0 0 8px; font-size: 18px; }
.gb-plan-desc { color: #666; font-size: 13px; margin: 0 0 15px; }

.gb-price-row { display: flex; gap: 24px; margin-bottom: 15px; }
.gb-price-block { text-align: center; }
.gb-label { display: block; font-size: 12px; color: #999; margin-bottom: 4px; }
.gb-price { font-size: 24px; font-weight: 700; color: #764ba2; }
.gb-original { font-size: 15px; color: #999; text-decoration: line-through; }
.gb-save { font-size: 15px; font-weight: 700; color: #e74c3c; }

.gb-stats { display: flex; gap: 24px; margin-bottom: 12px; padding: 10px; background: #f8f9ff; border-radius: 8px; }
.gb-stat { text-align: center; flex: 1; }
.gb-stat-num { display: block; font-size: 20px; font-weight: 700; color: #667eea; }
.gb-stat-label { font-size: 12px; color: #999; }

.gb-time-info { font-size: 13px; color: #999; margin-bottom: 15px; }
.gb-time-info i { margin-right: 4px; }
.gb-duration { margin-left: 16px; }

.gb-actions { display: flex; gap: 10px; margin-top: auto; }
.gb-btn-start { flex: 1; }
.gb-btn-open { flex: 1; }

.gb-open-groups { margin-top: 15px; padding-top: 15px; border-top: 1px solid #eee; }
.gb-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.gb-table th { background: #f5f5f5; padding: 8px 10px; text-align: left; font-weight: 600; color: #666; }
.gb-table td { padding: 8px 10px; border-bottom: 1px solid #f0f0f0; }
.gb-table code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-size: 11px; }
.gb-mini-progress {
    height: 6px; background: #eee; border-radius: 3px; overflow: hidden; margin-bottom: 2px; width: 80px;
}
.gb-mini-bar {
    height: 100%; background: linear-gradient(90deg, #667eea, #764ba2); border-radius: 3px;
}
.gb-no-groups { text-align: center; color: #999; padding: 10px 0; }

.gb-form-inline { display: inline; }

.alert-success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; padding: 12px 16px; border-radius: 6px; }
.alert-warning { background: #fff3cd; color: #856404; border: 1px solid #ffeaa7; padding: 12px 16px; border-radius: 6px; }
.alert i { margin-right: 6px; }

@media (max-width: 768px) {
    .gb-plan-card { flex-direction: column; }
    .gb-plan-image { width: 100%; height: 200px; }
    .gb-actions { flex-direction: column; }
    .gb-stats { flex-wrap: wrap; }
}
</style>

<script nonce="<%= Session("csp_nonce") %>">
function toggleGroups(planId) {
    var el = document.getElementById('groups_' + planId);
    if (el.style.display === 'none') {
        el.style.display = 'block';
    } else {
        el.style.display = 'none';
    }
}
</script>

<!--#include file="includes/footer.asp"-->
<%
If Not rsPlans Is Nothing Then
    If rsPlans.State = 1 Then rsPlans.Close
    Set rsPlans = Nothing
End If
Call CloseConnection()
%>
