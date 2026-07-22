<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>香调配比设置诊断 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI','Microsoft YaHei',sans-serif; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .diag-container { max-width: 900px; margin: 0 auto; }
        .diag-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.06); box-shadow: 0 4px 20px rgba(0,0,0,0.3); }
        h1 { color: #fff; font-size: 20px; margin-bottom: 5px; display: flex; align-items: center; gap: 10px; }
        h1 i { color: #00bcd4; }
        h2 { color: #80deea; font-size: 16px; margin: 20px 0 10px; }
        .result-item { padding: 15px; margin: 10px 0; border-radius: 8px; border-left: 4px solid; }
        .result-item.found { border-left-color: #4CAF50; background: rgba(76,175,80,0.08); }
        .result-item.notfound { border-left-color: #f44336; background: rgba(244,67,54,0.08); }
        .result-item ul { margin: 8px 0 0 0; padding-left: 20px; }
        .result-item li { font-size: 13px; color: #d0d6e0; margin: 3px 0; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 10px 14px; border: 1px solid rgba(255,255,255,0.06); text-align: left; font-size: 13px; }
        th { background: rgba(0,188,212,0.12); color: #80deea; font-weight: 600; }
        td { color: #d0d6e0; }
        .btn-row { margin-top: 20px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.06); display: flex; gap: 12px; flex-wrap: wrap; }
        .btn { display: inline-flex; align-items: center; gap: 6px; padding: 10px 18px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: 500; cursor: pointer; border: none; transition: all 0.2s; }
        .btn-primary { background: linear-gradient(135deg, #00bcd4, #00838f); color: #fff; }
        .btn-primary:hover { background: linear-gradient(135deg, #00acc1, #006064); transform: translateY(-1px); }
        .btn-outline { background: transparent; color: #00bcd4; border: 1px solid rgba(0,188,212,0.4); }
        .btn-outline:hover { background: rgba(0,188,212,0.1); }
        @media (max-width: 768px) {
            .main-content { margin-left: 0; padding: 15px; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-stethoscope"></i> 香调配比设置诊断</h2>
            <div class="breadcrumb">
                <a href="index.asp">技术中心</a> / <a href="product_settings.asp?tab=ratio">香调配比参数</a> / <span>诊断</span>
            </div>
        </div>
        
    <div class="diag-container">
    <div class="diag-card">
    <h1><i class="fas fa-stethoscope"></i> SiteSettings 表诊断</h1>
    
    <h2>1. 检查 MinTopPercent</h2>
    <% 
    Dim rs1
    Set rs1 = ExecuteQuery("SELECT * FROM SiteSettings WHERE SettingKey = 'MinTopPercent'")
    If Not rs1 Is Nothing Then
        If Not rs1.EOF Then
            Response.Write "<div class='result-item found'><strong>✓ 找到记录:</strong></div>"
            Response.Write "<ul>"
            Response.Write "<li>SettingKey: " & rs1("SettingKey") & "</li>"
            Response.Write "<li>SettingValue: " & rs1("SettingValue") & "</li>"
            If Not IsNull(rs1("UpdatedAt")) Then
                Response.Write "<li>UpdatedAt: " & rs1("UpdatedAt") & "</li>"
            End If
            Response.Write "</ul>"
        Else
            Response.Write "<div class='result-item notfound'><strong>✗ 未找到 MinTopPercent 记录</strong></div>"
        End If
        rs1.Close
    Else
        Response.Write "<div class='result-item notfound'><strong>✗ 查询失败</strong></div>"
    End If
    %>
    
    <h2>2. 检查 MinMiddlePercent</h2>
    <% 
    Dim rs2
    Set rs2 = ExecuteQuery("SELECT * FROM SiteSettings WHERE SettingKey = 'MinMiddlePercent'")
    If Not rs2 Is Nothing Then
        If Not rs2.EOF Then
            Response.Write "<div class='result-item found'><strong>✓ 找到记录:</strong></div>"
            Response.Write "<ul>"
            Response.Write "<li>SettingKey: " & rs2("SettingKey") & "</li>"
            Response.Write "<li>SettingValue: " & rs2("SettingValue") & "</li>"
            If Not IsNull(rs2("UpdatedAt")) Then
                Response.Write "<li>UpdatedAt: " & rs2("UpdatedAt") & "</li>"
            End If
            Response.Write "</ul>"
        Else
            Response.Write "<div class='result-item notfound'><strong>✗ 未找到 MinMiddlePercent 记录</strong></div>"
        End If
        rs2.Close
    Else
        Response.Write "<div class='result-item notfound'><strong>✗ 查询失败</strong></div>"
    End If
    %>
    
    <h2>3. 检查 MinBasePercent</h2>
    <% 
    Dim rs3
    Set rs3 = ExecuteQuery("SELECT * FROM SiteSettings WHERE SettingKey = 'MinBasePercent'")
    If Not rs3 Is Nothing Then
        If Not rs3.EOF Then
            Response.Write "<div class='result-item found'><strong>✓ 找到记录:</strong></div>"
            Response.Write "<ul>"
            Response.Write "<li>SettingKey: " & rs3("SettingKey") & "</li>"
            Response.Write "<li>SettingValue: " & rs3("SettingValue") & "</li>"
            If Not IsNull(rs3("UpdatedAt")) Then
                Response.Write "<li>UpdatedAt: " & rs3("UpdatedAt") & "</li>"
            End If
            Response.Write "</ul>"
        Else
            Response.Write "<div class='result-item notfound'><strong>✗ 未找到 MinBasePercent 记录</strong></div>"
        End If
        rs3.Close
    Else
        Response.Write "<div class='result-item notfound'><strong>✗ 查询失败</strong></div>"
    End If
    %>
    
    <h2>4. SiteSettings 表所有记录</h2>
    <table border="1" cellpadding="5">
        <tr>
            <th>SettingKey</th>
            <th>SettingValue</th>
            <th>UpdatedAt</th>
        </tr>
    <% 
    Dim rsAll
    Set rsAll = ExecuteQuery("SELECT SettingKey, SettingValue, UpdatedAt FROM SiteSettings ORDER BY SettingKey")
    If Not rsAll Is Nothing Then
        Do While Not rsAll.EOF
    %>
        <tr>
            <td><%= rsAll("SettingKey") %></td>
            <td><%= rsAll("SettingValue") %></td>
            <td><%= rsAll("UpdatedAt") %></td>
        </tr>
    <% 
            rsAll.MoveNext
        Loop
        rsAll.Close
    End If
    %>
    </table>
    </div>
    
    <div class="btn-row">
        <a href="product_settings.asp?tab=ratio" class="btn btn-primary"><i class="fas fa-arrow-left"></i> 返回设置页面</a>
        <a href="index.asp" class="btn btn-outline"><i class="fas fa-home"></i> 返回技术中心</a>
    </div>
    
    </div>
    </div>
</body>
</html>
