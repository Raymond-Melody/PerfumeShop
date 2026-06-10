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

' 获取系统统计
Dim totalAdmins, totalRoles, todayLogs
totalAdmins = GetScalar("SELECT COUNT(*) FROM AdminUsers")
totalRoles = GetScalar("SELECT COUNT(*) FROM AdminRoles")
todayLogs = GetScalar("SELECT COUNT(*) FROM AdminLogs WHERE CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>系统概览 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); text-align: center; }
        .stat-icon { width: 60px; height: 60px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 15px; font-size: 24px; color: white; background: linear-gradient(135deg, #fa709a 0%, #fee140 100%); }
        .stat-value { font-size: 28px; font-weight: bold; color: #fff; }
        .stat-label { color: #888; font-size: 14px; margin-top: 5px; }
        .quick-actions { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; }
        .action-btn { display: flex; flex-direction: column; align-items: center; padding: 25px; background: #2d2d44; border-radius: 10px; text-decoration: none; color: #e0e0e0; transition: all 0.3s; border: 1px solid rgba(255,255,255,0.06); }
        .action-btn:hover { background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; transform: translateY(-3px); }
        .action-btn i { font-size: 28px; margin-bottom: 10px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-cogs"></i> 系统概览</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>概览</span>
            </div>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon"><i class="fas fa-users-cog"></i></div>
                <div class="stat-value"><%= totalAdmins %></div>
                <div class="stat-label">管理员数</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon"><i class="fas fa-user-tag"></i></div>
                <div class="stat-value"><%= totalRoles %></div>
                <div class="stat-label">角色数</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon"><i class="fas fa-history"></i></div>
                <div class="stat-value"><%= todayLogs %></div>
                <div class="stat-label">今日日志</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon"><i class="fas fa-shield-alt"></i></div>
                <div class="stat-value">ON</div>
                <div class="stat-label">系统状态</div>
            </div>
        </div>
        
        <div class="dashboard-card">
            <h3><i class="fas fa-bolt"></i> 快捷操作</h3>
            <div class="quick-actions">
                <a href="roles.asp" class="action-btn"><i class="fas fa-user-tag"></i><span>角色管理</span></a>
                <a href="admins.asp" class="action-btn"><i class="fas fa-users-cog"></i><span>管理员管理</span></a>
                <a href="logs.asp" class="action-btn"><i class="fas fa-history"></i><span>操作日志</span></a>
                <a href="site_settings.asp" class="action-btn"><i class="fas fa-sliders-h"></i><span>站点设置</span></a>
                <a href="settings.asp" class="action-btn"><i class="fas fa-cog"></i><span>管理配置</span></a>
                <a href="statistics.asp" class="action-btn"><i class="fas fa-chart-line"></i><span>数据统计</span></a>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
