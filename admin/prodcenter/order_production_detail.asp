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

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then If Not rs.EOF Then val = rs(0) : rs.Close
    Else : Err.Clear
    End If
    Set rs = Nothing : GetScalar = val
End Function

Dim orderId
orderId = SafeNum(Request.QueryString("id"))
If orderId = 0 Then
    Response.Redirect "order_production.asp?msg=无效的订单ID"
    Response.End
End If

' ========== 订单基本信息 ==========
Dim rsOrder
Set rsOrder = conn.Execute("SELECT o.*, u.Username FROM Orders o LEFT JOIN Users u ON o.UserID=u.UserID WHERE o.OrderID=" & orderId)
If rsOrder Is Nothing Or rsOrder.EOF Then
    Response.Redirect "order_production.asp?msg=订单不存在"
    Response.End
End If
Dim oNo, oName, oAmt, oStatus, oCreated, oCust
oNo = rsOrder("OrderNo") & ""
oName = rsOrder("ShippingName") & ""
oAmt = SafeNum(rsOrder("TotalAmount"))
oStatus = rsOrder("Status") & ""
oCreated = rsOrder("CreatedAt") & ""
oCust = rsOrder("Username") & ""
rsOrder.Close : Set rsOrder = Nothing

' ========== 该订单的工单列表 ==========
Dim rsPOs
Set rsPOs = conn.Execute("SELECT * FROM ProductionOrders WHERE OrderID=" & orderId & " ORDER BY CreatedAt DESC")

' ========== 生产日志 ==========
Dim rsLogs
Set rsLogs = conn.Execute("SELECT pl.* FROM ProductionLogs pl INNER JOIN ProductionOrders po ON pl.ProductionID=po.ProductionID WHERE po.OrderID=" & orderId & " ORDER BY pl.CreatedAt DESC")

' ========== 确定当前生产阶段 ==========
Dim latestPOStatus
latestPOStatus = GetScalar("SELECT TOP 1 Status FROM ProductionOrders WHERE OrderID=" & orderId & " ORDER BY UpdatedAt DESC")
If latestPOStatus = "" Or latestPOStatus = "0" Then latestPOStatus = "NoPO"

' 阶段定义
Dim stages(4, 2) ' 0=名称, 1=图标, 2=状态
stages(0,0) = "排产" : stages(0,1) = "calendar-alt"
stages(1,0) = "生产" : stages(1,1) = "cogs"
stages(2,0) = "质检" : stages(2,1) = "check-circle"
stages(3,0) = "入库" : stages(3,1) = "warehouse"
stages(4,0) = "发货" : stages(4,1) = "truck"

' 阶段完成映射
Dim currentStage : currentStage = -1
Select Case latestPOStatus
    Case "Pending"   : currentStage = 0
    Case "InProgress": currentStage = 1
    Case "Completed" : currentStage = 2
    Case "QC_Passed" : currentStage = 3
    Case "WarehouseIn","ShippedOut" : currentStage = 4
    Case "QC_Fail"   : currentStage = -2
End Select

' 阶段时间
Dim stageTimes(4)
stageTimes(0) = GetScalar("SELECT TOP 1 CONVERT(NVARCHAR(10),MIN(CreatedAt),120) FROM ProductionOrders WHERE OrderID=" & orderId)
stageTimes(1) = GetScalar("SELECT TOP 1 CONVERT(NVARCHAR(10),MIN(StartedAt),120) FROM ProductionOrders WHERE OrderID=" & orderId & " AND StartedAt IS NOT NULL")
stageTimes(2) = GetScalar("SELECT TOP 1 CONVERT(NVARCHAR(10),MAX(QCPassedAt),120) FROM ProductionOrders WHERE OrderID=" & orderId & " AND QCPassedAt IS NOT NULL")
stageTimes(3) = GetScalar("SELECT TOP 1 CONVERT(NVARCHAR(10),MAX(WarehouseInAt),120) FROM ProductionOrders WHERE OrderID=" & orderId & " AND WarehouseInAt IS NOT NULL")
stageTimes(4) = GetScalar("SELECT TOP 1 CONVERT(NVARCHAR(10),MAX(ShippedOutAt),120) FROM ProductionOrders WHERE OrderID=" & orderId & " AND ShippedOutAt IS NOT NULL")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>订单生产详情 #<%=Server.HTMLEncode(oNo)%> - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #4CAF50; }

        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 14px 20px; font-weight: 600; font-size: 15px; color: #e0e0e0; border-bottom: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; gap: 8px; }
        .card-body { padding: 16px 20px; }

        .order-info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; }
        .info-item .label { font-size: 11px; color: #888; margin-bottom: 4px; }
        .info-item .value { font-size: 16px; font-weight: 600; color: #e0e0e0; }

        .progress-container { padding: 10px 0; }
        .progress-stages { display: flex; align-items: center; justify-content: center; gap: 0; padding: 20px 0; }
        .stage-item { display: flex; flex-direction: column; align-items: center; position: relative; min-width: 90px; }
        .stage-icon { width: 48px; height: 48px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; transition: all 0.3s; margin-bottom: 8px; }
        .stage-icon.done { background: rgba(76,175,80,0.2); color: #4CAF50; border: 2px solid #4CAF50; }
        .stage-icon.active { background: rgba(33,150,243,0.2); color: #2196F3; border: 2px solid #2196F3; animation: pulse 2s infinite; }
        .stage-icon.pending { background: rgba(255,255,255,0.05); color: #555; border: 2px solid rgba(255,255,255,0.08); }
        .stage-icon.fail { background: rgba(244,67,54,0.2); color: #f44336; border: 2px solid #f44336; }
        .stage-label { font-size: 12px; font-weight: 600; color: #e0e0e0; }
        .stage-time { font-size: 11px; color: #666; margin-top: 3px; }
        .stage-connector { flex: 1; height: 2px; min-width: 30px; }
        .stage-connector.done { background: #4CAF50; }
        .stage-connector.pending { background: rgba(255,255,255,0.08); }
        @keyframes pulse { 0%, 100% { box-shadow: 0 0 0 0 rgba(33,150,243,0.4); } 50% { box-shadow: 0 0 0 8px rgba(33,150,243,0); } }

        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th { background: rgba(76,175,80,0.12); color: #81c784; padding: 12px 14px; text-align: left; font-weight: 600; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table td { padding: 12px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }

        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-paid { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .badge-processing { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .badge-shipped { background: rgba(156,39,176,0.15); color: #ba68c8; }

        .timeline { position: relative; padding-left: 30px; }
        .timeline::before { content: ''; position: absolute; left: 14px; top: 0; bottom: 0; width: 2px; background: rgba(255,255,255,0.06); }
        .timeline-item { position: relative; padding: 8px 0 20px 10px; }
        .timeline-item::before { content: ''; position: absolute; left: -20px; top: 12px; width: 10px; height: 10px; border-radius: 50%; background: #4CAF50; }
        .timeline-time { font-size: 11px; color: #666; }
        .timeline-content { margin-top: 4px; color: #e0e0e0; font-size: 13px; }

        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .mt-20 { margin-top: 20px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->

    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-search-location"></i> 订单生产详情</h2>
            <a href="order_production.asp" class="btn btn--outline btn--sm"><i class="fas fa-arrow-left"></i> 返回追踪列表</a>
        </div>

        <!-- 订单基本信息卡片 -->
        <div class="card">
            <div class="card-header" style="background:rgba(76,175,80,0.06);">
                <i class="fas fa-shopping-cart" style="color:#4CAF50;"></i> 订单 #<%=Server.HTMLEncode(oNo)%>
            </div>
            <div class="card-body">
                <div class="order-info-grid">
                    <div class="info-item">
                        <div class="label"><i class="fas fa-hashtag"></i> 订单号</div>
                        <div class="value">#<%=Server.HTMLEncode(oNo)%></div>
                    </div>
                    <div class="info-item">
                        <div class="label"><i class="fas fa-user"></i> 客户</div>
                        <div class="value"><%=Server.HTMLEncode(oName)%></div>
                    </div>
                    <div class="info-item">
                        <div class="label"><i class="fas fa-user-tag"></i> 账户</div>
                        <div class="value"><%=Server.HTMLEncode(oCust)%></div>
                    </div>
                    <div class="info-item">
                        <div class="label"><i class="fas fa-dollar-sign"></i> 金额</div>
                        <div class="value" style="color:#4CAF50;">¥<%=FormatNumber(oAmt,2)%></div>
                    </div>
                    <div class="info-item">
                        <div class="label"><i class="fas fa-flag"></i> 订单状态</div>
                        <div class="value"><span class="badge badge-<%=LCase(oStatus)%>"><%=oStatus%></span></div>
                    </div>
                    <div class="info-item">
                        <div class="label"><i class="fas fa-calendar"></i> 下单时间</div>
                        <div class="value"><%=IIf(oCreated="","-",Left(oCreated,10))%></div>
                    </div>
                    <div class="info-item">
                        <div class="label"><i class="fas fa-industry"></i> 生产阶段</div>
                        <div class="value" style="color:<%=IIf(latestPOStatus="NoPO","#888","#2196F3")%>;"><%=IIf(latestPOStatus="NoPO","未排产","第 "&(currentStage+1)&"/5 阶段")%></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- 5阶段可视化进度条 -->
        <div class="card">
            <div class="card-header" style="background:rgba(33,150,243,0.06);">
                <i class="fas fa-tasks" style="color:#2196F3;"></i> 生产进度追踪
            </div>
            <div class="card-body">
                <div class="progress-container">
                    <div class="progress-stages">
                        <% Dim s
                        For s = 0 To 4
                            Dim sClass, sConnectorClass
                            If s < currentStage Then
                                sClass = "done"
                            ElseIf s = currentStage Then
                                sClass = "active"
                            ElseIf currentStage = -2 And s = 2 Then
                                sClass = "fail"
                            Else
                                sClass = "pending"
                            End If
                            sConnectorClass = "pending"
                            If s < currentStage Then sConnectorClass = "done"
                        %>
                        <div class="stage-item">
                            <div class="stage-icon <%=sClass%>">
                                <i class="fas fa-<%=stages(s,1)%>"></i>
                            </div>
                            <div class="stage-label"><%=stages(s,0)%></div>
                            <div class="stage-time"><%=IIf(stageTimes(s)="0" Or stageTimes(s)="","-",stageTimes(s))%></div>
                        </div>
                        <% If s < 4 Then %>
                        <div class="stage-connector <%=sConnectorClass%>"></div>
                        <% End If
                        Next %>
                    </div>
                </div>
            </div>
        </div>

        <!-- 关联工单列表 -->
        <div class="card">
            <div class="card-header" style="background:rgba(255,152,0,0.06);">
                <i class="fas fa-clipboard-list" style="color:#FF9800;"></i> 关联生产工单
            </div>
            <div class="card-body" style="overflow-x:auto;">
                <table class="data-table">
                    <thead>
                        <tr><th>工单号</th><th>配方名</th><th>计划量</th><th>状态</th><th>优先级</th><th>负责人</th><th>创建时间</th><th>开始时间</th><th>完成时间</th><th>操作</th></tr>
                    </thead>
                    <tbody>
                        <%
                        Dim poRow : poRow = 0
                        If Not rsPOs Is Nothing Then
                            Do While Not rsPOs.EOF
                                poRow = poRow + 1
                                Dim poStatus : poStatus = rsPOs("Status") & ""
                                Dim poBadgeClass
                                Select Case poStatus
                                    Case "Pending"    : poBadgeClass = "badge-processing"
                                    Case "InProgress" : poBadgeClass = "badge-processing"
                                    Case "Completed"  : poBadgeClass = "badge-paid"
                                    Case "QC_Passed"  : poBadgeClass = "badge-paid"
                                    Case "QC_Fail"    : poBadgeClass = "badge-processing"
                                    Case "WarehouseIn","ShippedOut" : poBadgeClass = "badge-shipped"
                                    Case Else         : poBadgeClass = "badge-paid"
                                End Select
                        %>
                        <tr>
                            <td><strong><%=rsPOs("WorkOrderNo") & ""%></strong></td>
                            <td><%=rsPOs("RecipeName") & ""%></td>
                            <td><%=rsPOs("PlannedQty") & ""%></td>
                            <td><span class="badge <%=poBadgeClass%>"><%=poStatus%></span></td>
                            <td><%=IIf(IsNull(rsPOs("Priority")),"-",rsPOs("Priority"))%></td>
                            <td><%=rsPOs("AssignedTo") & ""%></td>
                            <td style="color:#888;"><%=IIf(IsNull(rsPOs("CreatedAt")),"-",Left(rsPOs("CreatedAt"),10))%></td>
                            <td style="color:#888;"><%=IIf(IsNull(rsPOs("StartedAt")),"-",Left(rsPOs("StartedAt"),10))%></td>
                            <td style="color:#888;"><%=IIf(IsNull(rsPOs("CompletedAt")),"-",Left(rsPOs("CompletedAt"),10))%></td>
                            <td><a href="production_management.asp?id=<%=rsPOs("ProductionID")%>" class="btn btn--outline btn--sm" style="font-size:11px;padding:4px 10px;"><i class="fas fa-eye"></i></a></td>
                        </tr>
                        <%
                                rsPOs.MoveNext
                            Loop
                            rsPOs.Close
                        End If
                        Set rsPOs = Nothing
                        If poRow = 0 Then %>
                        <tr><td colspan="10" class="text-center text-muted" style="padding:30px;">暂无关联工单</td></tr>
                        <% End If %>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- 生产日志时间线 -->
        <div class="card">
            <div class="card-header" style="background:rgba(156,39,176,0.06);">
                <i class="fas fa-history" style="color:#9C27B0;"></i> 生产日志
            </div>
            <div class="card-body">
                <%
                Dim logRow : logRow = 0
                If Not rsLogs Is Nothing And Not rsLogs.EOF Then
                %>
                <div class="timeline">
                    <%
                    Do While Not rsLogs.EOF
                        logRow = logRow + 1
                    %>
                    <div class="timeline-item">
                        <div class="timeline-time"><%=IIf(IsNull(rsLogs("CreatedAt")),"-",Left(rsLogs("CreatedAt"),19))%></div>
                        <div class="timeline-content">
                            <strong><%=rsLogs("CreatedBy") & ""%></strong> — <%=rsLogs("Notes") & ""%>
                            <span style="color:#888;margin-left:8px;font-size:11px;">[<%=rsLogs("Status") & ""%>]</span>
                        </div>
                    </div>
                    <%
                        rsLogs.MoveNext
                    Loop
                    rsLogs.Close
                    %>
                </div>
                <%
                Else
                    If IsObject(rsLogs) And Not rsLogs Is Nothing Then rsLogs.Close
                %>
                <div class="text-center text-muted" style="padding:30px;">暂无生产日志</div>
                <% End If
                Set rsLogs = Nothing
                %>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
