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

' 获取营销活动列表
Dim rsCampaigns
Set rsCampaigns = ExecuteQuery("SELECT * FROM MarketingCampaigns ORDER BY StartDate DESC")

' 获取统计
Dim activeCampaigns, totalCoupons
totalCoupons = GetScalar("SELECT COUNT(*) FROM Coupons")
activeCampaigns = GetScalar("SELECT COUNT(*) FROM MarketingCampaigns WHERE EndDate >= CAST(GETDATE() AS DATE) AND IsActive = 1")

Call LogAdminAction("查看营销活动", "operation", "MarketingCampaigns", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>营销活动 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .stats-cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); text-align: center; }
        .stat-card i { font-size: 36px; color: #667eea; margin-bottom: 10px; }
        .stat-card h3 { font-size: 32px; margin: 10px 0; color: #333; }
        .stat-card p { color: #666; margin: 0; }
        .campaigns-table { width: 100%; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .campaigns-table th { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; text-align: left; }
        .campaigns-table td { padding: 15px; border-bottom: 1px solid #f0f0f0; }
        .campaigns-table tr:hover { background: #f8f9fa; }
        .campaign-type { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .type-discount { background: #e3f2fd; color: #1976d2; }
        .type-coupon { background: #e8f5e9; color: #2e7d32; }
        .type-gift { background: #fff3e0; color: #f57c00; }
        .campaign-status { padding: 4px 12px; border-radius: 12px; font-size: 12px; }
        .status-active { background: #e8f5e9; color: #2e7d32; }
        .status-ended { background: #ffebee; color: #c62828; }
        .status-pending { background: #fff3e0; color: #f57c00; }
        .date-range { color: #666; font-size: 13px; }
        .btn-edit { padding: 6px 15px; background: #667eea; color: white; border-radius: 6px; text-decoration: none; font-size: 13px; }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; }
        .tab { padding: 10px 25px; background: white; border-radius: 8px; cursor: pointer; }
        .tab.active { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-bullhorn"></i> 营销活动</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>营销活动</span>
            </div>
        </div>
        
        <div class="stats-cards">
            <div class="stat-card">
                <i class="fas fa-bullhorn"></i>
                <h3><%= GetScalar("SELECT COUNT(*) FROM MarketingCampaigns") %></h3>
                <p>总活动数</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-play-circle"></i>
                <h3><%= activeCampaigns %></h3>
                <p>进行中</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-ticket-alt"></i>
                <h3><%= totalCoupons %></h3>
                <p>优惠券总数</p>
            </div>
        </div>
        
        <div class="tabs">
            <div class="tab active" onclick="showTab('campaigns')">活动列表</div>
            <div class="tab" onclick="showTab('coupons')">优惠券管理</div>
        </div>
        
        <div style="margin-bottom: 20px;">
            <a href="campaign_edit.asp" class="admin-btn admin-btn-primary"><i class="fas fa-plus"></i> 创建活动</a>
            <a href="coupons.asp" class="admin-btn admin-btn-secondary"><i class="fas fa-ticket-alt"></i> 优惠券管理</a>
        </div>
        
        <table class="campaigns-table">
            <thead>
                <tr>
                    <th>活动名称</th>
                    <th>类型</th>
                    <th>时间范围</th>
                    <th>参与人数</th>
                    <th>销售额</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsCampaigns Is Nothing Then %>
                <% Do While Not rsCampaigns.EOF %>
                <tr>
                    <td>
                        <strong><%= rsCampaigns("CampaignName") %></strong>
                        <div style="font-size: 12px; color: #999;"><%= rsCampaigns("Description") %></div>
                    </td>
                    <td>
                        <% 
                        Dim typeClass, typeName
                        Select Case rsCampaigns("CampaignType")
                            Case "discount": typeClass = "type-discount": typeName = "折扣"
                            Case "coupon": typeClass = "type-coupon": typeName = "优惠券"
                            Case "gift": typeClass = "type-gift": typeName = "赠品"
                            Case Else: typeClass = "": typeName = rsCampaigns("CampaignType")
                        End Select
                        %>
                        <span class="campaign-type <%= typeClass %>"><%= typeName %></span>
                    </td>
                    <td class="date-range">
                        <%= SafeFormatDateTime(rsCampaigns("StartDate"), 2) %> ~ <%= SafeFormatDateTime(rsCampaigns("EndDate"), 2) %>
                    </td>
                    <td><%= rsCampaigns("ParticipantCount") %></td>
                    <td>¥<%= FormatNumber(CDbl("0" & rsCampaigns("TotalSales")), 2) %></td>
                    <td>
                        <% 
                        Dim statusClass, statusName
                        If Date() < rsCampaigns("StartDate") Then
                            statusClass = "status-pending": statusName = "未开始"
                        ElseIf Date() > rsCampaigns("EndDate") Then
                            statusClass = "status-ended": statusName = "已结束"
                        ElseIf rsCampaigns("IsActive") = True Then
                            statusClass = "status-active": statusName = "进行中"
                        Else
                            statusClass = "status-ended": statusName = "已停用"
                        End If
                        %>
                        <span class="campaign-status <%= statusClass %>"><%= statusName %></span>
                    </td>
                    <td>
                        <a href="campaign_edit.asp?id=<%= rsCampaigns("CampaignID") %>" class="btn-edit"><i class="fas fa-edit"></i> 编辑</a>
                    </td>
                </tr>
                <% rsCampaigns.MoveNext %>
                <% Loop %>
                <% rsCampaigns.Close %>
                <% End If %>
            </tbody>
        </table>
    </div>
    
    <script>
        function showTab(tab) {
            if (tab === 'coupons') {
                window.location.href = 'coupons.asp';
            }
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
