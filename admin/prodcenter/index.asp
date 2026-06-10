<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

Dim dashIsFullAccess
dashIsFullAccess = (Session("AdminRoleCode") = "SUPER_ADMIN" Or Session("AdminRoleCode") = "PROD_MANAGER")

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

' 生产统计
Dim poPending, poInProgress, poCompleted, poToday
poPending = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Pending'"))
poInProgress = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='InProgress'"))
poCompleted = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed'"))
poToday = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed' AND CompletedAt >= CAST(GETDATE() AS DATE)"))

' 订单统计
Dim orderPending, orderProcessing
orderPending = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE Status='Paid'"))
orderProcessing = SafeNum(GetScalar("SELECT COUNT(*) FROM Orders WHERE Status='Processing'"))

' 成品库存统计
Dim prodTotal, prodActive
prodTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM Products"))
prodActive = SafeNum(GetScalar("SELECT COUNT(*) FROM Products WHERE IsActive=1"))

' 瓶子库存
Dim bottleTotal, bottleActive
bottleTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleStyles"))
bottleActive = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleStyles WHERE IsActive=1"))

' 近期生产工单
Dim rsRecentPO
On Error Resume Next
Set rsRecentPO = conn.Execute("SELECT TOP 5 ProductionOrderID, OrderID, ISNULL(RecipeName, '-') as ProductName, PlannedQty, Status, CreatedAt FROM ProductionOrders ORDER BY CreatedAt DESC")
If Err.Number <> 0 Then
    Err.Clear
    Set rsRecentPO = Nothing
End If
On Error GoTo 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>生产概览 - 产品生产管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 12px; margin-bottom: 25px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 15px; margin-bottom: 30px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 20px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 12px; color: #888; margin-bottom: 8px; }
        .stat-card .stat-value { font-size: 28px; font-weight: 700; }
        .stat-card .stat-sub { font-size: 12px; color: #888; margin-top: 5px; }
        .stat-prod .stat-value { color: #4CAF50; }
        .stat-order .stat-value { color: #2196F3; }
        .stat-warehouse .stat-value { color: #FF9800; }
        .stat-bottle .stat-value { color: #9C27B0; }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(76,175,80,0.15); color: #81c784; font-weight: 600; padding: 12px 15px; text-align: left; font-size: 13px; }
        .data-table td { padding: 12px 15px; color: #e0e0e0; font-size: 14px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .section-title { font-size: 18px; color: #e0e0e0; margin-bottom: 15px; display: flex; align-items: center; gap: 10px; }
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-pending { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .status-progress { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .status-completed { background: rgba(76,175,80,0.15); color: #81c784; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <h2 class="page-title"><i class="fas fa-industry" style="color:#4CAF50;"></i> 产品生产概览</h2>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card stat-prod">
                <div class="stat-label"><i class="fas fa-clipboard-list"></i> 生产工单</div>
                <div class="stat-value"><%=poInProgress%></div>
                <div class="stat-sub">待处理 <%=poPending%> | 已完成 <%=poCompleted%> | 今日完成 <%=poToday%></div>
            </div>
            <div class="stat-card stat-order">
                <div class="stat-label"><i class="fas fa-shopping-cart"></i> 待生产订单</div>
                <div class="stat-value"><%=orderPending + orderProcessing%></div>
                <div class="stat-sub">已付款 <%=orderPending%> | 处理中 <%=orderProcessing%></div>
            </div>
            <div class="stat-card stat-warehouse">
                <div class="stat-label"><i class="fas fa-box"></i> 成品库存</div>
                <div class="stat-value"><%=prodActive%></div>
                <div class="stat-sub">总计 <%=prodTotal%> 个产品</div>
            </div>
            <div class="stat-card stat-bottle">
                <div class="stat-label"><i class="fas fa-flask"></i> 瓶子款式</div>
                <div class="stat-value"><%=bottleActive%></div>
                <div class="stat-sub">总计 <%=bottleTotal%> 种</div>
            </div>
        </div>
        
        <!-- 近期生产工单 -->
        <div style="margin-top:25px;">
            <h3 class="section-title"><i class="fas fa-clock" style="color:#4CAF50;"></i> 近期生产工单</h3>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>工单ID</th>
                        <th>订单号</th>
                        <th>产品</th>
                        <th>计划量</th>
                        <th>状态</th>
                        <th>创建时间</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not rsRecentPO Is Nothing Then
                        Do While Not rsRecentPO.EOF %>
                    <tr>
                        <td><strong>#<%=rsRecentPO("ProductionOrderID")%></strong></td>
                        <td><%=IIF(IsNull(rsRecentPO("OrderID")) Or rsRecentPO("OrderID")=0,"-",rsRecentPO("OrderID"))%></td>
                        <td><%=Server.HTMLEncode(rsRecentPO("ProductName") & "")%></td>
                        <td><%=rsRecentPO("PlannedQty")%></td>
                        <td><%
                            Select Case CStr(rsRecentPO("Status") & "")
                                Case "Pending": Response.Write "<span class='status-badge status-pending'>待生产</span>"
                                Case "InProgress": Response.Write "<span class='status-badge status-progress'>生产中</span>"
                                Case "Completed": Response.Write "<span class='status-badge status-completed'>已完成</span>"
                                Case Else: Response.Write rsRecentPO("Status")
                            End Select
                        %></td>
                        <td style="color:#888;"><%=IIF(IsNull(rsRecentPO("CreatedAt")) Or rsRecentPO("CreatedAt")="","-",Left(rsRecentPO("CreatedAt"),10))%></td>
                    </tr>
                    <%      rsRecentPO.MoveNext
                        Loop
                        rsRecentPO.Close
                        Set rsRecentPO = Nothing
                    End If %>
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
