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

' 原料库存统计
Dim rawMaterialCount, rawLowStock
rawMaterialCount = GetScalar("SELECT COUNT(*) FROM RawMaterialInventory")
rawLowStock = GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= MinStockLevel AND MinStockLevel > 0")

' 基香库存统计
Dim baseNoteCount, baseNoteActive
baseNoteCount = GetScalar("SELECT COUNT(*) FROM BaseNotes")
baseNoteActive = GetScalar("SELECT COUNT(*) FROM BaseNotes WHERE IsActive=1")

' 香调库存统计
Dim noteCount, noteLowStock, noteTotalStock
noteCount = GetScalar("SELECT COUNT(*) FROM NoteInventory")
noteLowStock = GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel AND MinStockLevel > 0")
noteTotalStock = GetScalar("SELECT ISNULL(SUM(StockQuantity),0) FROM NoteInventory")

' Accord生产统计
Dim accordPending, accordProcessing, accordCompleted
accordPending = GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status='Pending'")
accordProcessing = GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status='InProgress'")
accordCompleted = GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status='Completed'")

' 近期Accord生产
Dim rsRecentAccords
Set rsRecentAccords = ExecuteQuery("SELECT TOP 5 ProductionID, NoteName, BatchNo, PlannedQty, Status, CreatedAt FROM AccordProductions ORDER BY CreatedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>生产概览 - 半成品生产中心</title>
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
        .stat-raw .stat-value { color: #4CAF50; }
        .stat-base .stat-value { color: #2196F3; }
        .stat-note .stat-value { color: #FF9800; }
        .stat-accord .stat-value { color: #9C27B0; }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(33,150,243,0.15); color: #64b5f6; font-weight: 600; padding: 12px 15px; text-align: left; font-size: 13px; }
        .data-table td { padding: 12px 15px; color: #e0e0e0; font-size: 14px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .section-title { font-size: 18px; color: #e0e0e0; margin-bottom: 15px; display: flex; align-items: center; gap: 10px; }
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-pending { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .status-progress { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .status-completed { background: rgba(76,175,80,0.15); color: #81c784; }
        .alert-badge { background: rgba(244,67,54,0.15); color: #e57373; padding: 2px 8px; border-radius: 10px; font-size: 11px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <h2 class="page-title"><i class="fas fa-flask" style="color:#2196F3;"></i> 半成品生产概览</h2>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card stat-raw">
                <div class="stat-label"><i class="fas fa-boxes"></i> 原料种类</div>
                <div class="stat-value"><%= rawMaterialCount %></div>
                <div class="stat-sub"><% If rawLowStock > 0 Then %><span class="alert-badge"><%= rawLowStock %> 低库存</span><% Else %>库存正常<% End If %></div>
            </div>
            <div class="stat-card stat-base">
                <div class="stat-label"><i class="fas fa-database"></i> 基香</div>
                <div class="stat-value"><%= baseNoteActive %></div>
                <div class="stat-sub">总计 <%= baseNoteCount %> 种</div>
            </div>
            <div class="stat-card stat-note">
                <div class="stat-label"><i class="fas fa-layer-group"></i> 香调库存</div>
                <div class="stat-value"><%= noteTotalStock %></div>
                <div class="stat-sub"><%= noteCount %> 种香调<% If noteLowStock > 0 Then %> | <span class="alert-badge"><%= noteLowStock %> 低库存</span><% End If %></div>
            </div>
            <div class="stat-card stat-accord">
                <div class="stat-label"><i class="fas fa-cogs"></i> Accord生产</div>
                <div class="stat-value"><%= accordProcessing %></div>
                <div class="stat-sub">待处理 <%= accordPending %> | 已完成 <%= accordCompleted %></div>
            </div>
        </div>
        
        <!-- 近期Accord生产 -->
        <div style="margin-top:25px;">
            <h3 class="section-title"><i class="fas fa-clock" style="color:#2196F3;"></i> 近期Accord生产</h3>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>生产批号</th>
                        <th>香调名称</th>
                        <th>计划产量</th>
                        <th>状态</th>
                        <th>创建时间</th>
                    </tr>
                </thead>
                <tbody>
                    <% If Not rsRecentAccords Is Nothing Then
                        Do While Not rsRecentAccords.EOF %>
                    <tr>
                        <td><%= rsRecentAccords("BatchNo") %></td>
                        <td><strong><%= Server.HTMLEncode(rsRecentAccords("NoteName") & "") %></strong></td>
                        <td><%= rsRecentAccords("PlannedQty") %></td>
                        <td><%
                            Select Case rsRecentAccords("Status")
                                Case "Pending": Response.Write "<span class='status-badge status-pending'>待生产</span>"
                                Case "InProgress": Response.Write "<span class='status-badge status-progress'>生产中</span>"
                                Case "Completed": Response.Write "<span class='status-badge status-completed'>已完成</span>"
                                Case Else: Response.Write rsRecentAccords("Status")
                            End Select
                        %></td>
                        <td style="color:#888;"><%= FormatDateField(rsRecentAccords("CreatedAt")) %></td>
                    </tr>
                    <%      rsRecentAccords.MoveNext
                        Loop
                        rsRecentAccords.Close
                        Set rsRecentAccords = Nothing
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
