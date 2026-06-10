<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection

' 处理表单提交
Dim action
action = Request.Form("action")

If action = "update" Then
    Dim showFixed, showCustom, showKOL
    Dim minTopPercent, minMiddlePercent, minBasePercent
    
    ' 调试：输出接收到的值
    ' Response.Write "<pre>DEBUG: showFixed=" & Request.Form("showFixed") & ", showCustom=" & Request.Form("showCustom") & ", showKOL=" & Request.Form("showKOL") & "</pre>"
    
    ' 使用 Count 属性判断字段是否存在（比判断空字符串更可靠）
    If Request.Form("showFixed").Count = 0 Then
        showFixed = "0"
    Else
        showFixed = "1"
    End If
    
    If Request.Form("showCustom").Count = 0 Then
        showCustom = "0"
    Else
        showCustom = "1"
    End If
    
    If Request.Form("showKOL").Count = 0 Then
        showKOL = "0"
    Else
        showKOL = "1"
    End If
    
    ' 获取最小比例设置
    minTopPercent = Request.Form("minTopPercent")
    minMiddlePercent = Request.Form("minMiddlePercent")
    minBasePercent = Request.Form("minBasePercent")
    
    ' 验证输入值
    If Not IsNumeric(minTopPercent) Or CInt(minTopPercent) < 0 Or CInt(minTopPercent) > 100 Then minTopPercent = "10"
    If Not IsNumeric(minMiddlePercent) Or CInt(minMiddlePercent) < 0 Or CInt(minMiddlePercent) > 100 Then minMiddlePercent = "10"
    If Not IsNumeric(minBasePercent) Or CInt(minBasePercent) < 0 Or CInt(minBasePercent) > 100 Then minBasePercent = "10"
    
    ' 更新配置
    Dim sql1, sql2, sql3, sql4, sql5, sql6
    sql1 = "UPDATE SiteSettings SET SettingValue = '" & showFixed & "', UpdatedAt = GETDATE() WHERE SettingKey = 'ShowFixedSection'"
    sql2 = "UPDATE SiteSettings SET SettingValue = '" & showCustom & "', UpdatedAt = GETDATE() WHERE SettingKey = 'ShowCustomSection'"
    sql3 = "UPDATE SiteSettings SET SettingValue = '" & showKOL & "', UpdatedAt = GETDATE() WHERE SettingKey = 'ShowKOLSection'"
    sql4 = "UPDATE SiteSettings SET SettingValue = '" & minTopPercent & "', UpdatedAt = GETDATE() WHERE SettingKey = 'MinTopPercent'"
    sql5 = "UPDATE SiteSettings SET SettingValue = '" & minMiddlePercent & "', UpdatedAt = GETDATE() WHERE SettingKey = 'MinMiddlePercent'"
    sql6 = "UPDATE SiteSettings SET SettingValue = '" & minBasePercent & "', UpdatedAt = GETDATE() WHERE SettingKey = 'MinBasePercent'"
    
    Dim result1, result2, result3, result4, result5, result6
    result1 = ExecuteNonQuery(sql1)
    result2 = ExecuteNonQuery(sql2)
    result3 = ExecuteNonQuery(sql3)
    result4 = ExecuteNonQuery(sql4)
    result5 = ExecuteNonQuery(sql5)
    result6 = ExecuteNonQuery(sql6)
    
    ' 检查执行结果
    If Not result1 Or Not result2 Or Not result3 Or Not result4 Or Not result5 Or Not result6 Then
        Response.Write "<script>alert('更新失败: " & Replace(Session("LastDBError"), "'", "\'") & "');</script>"
    Else
        Response.Redirect "site_settings.asp?msg=更新成功"
    End If
End If

' 获取当前配置
Dim rsSettings, dictSettings
Set dictSettings = Server.CreateObject("Scripting.Dictionary")
Set rsSettings = ExecuteQuery("SELECT SettingKey, SettingValue FROM SiteSettings")

If Not rsSettings Is Nothing Then
    Do While Not rsSettings.EOF
        Dim key, value
        key = rsSettings("SettingKey")
        value = rsSettings("SettingValue")
        dictSettings.Item(key) = value
        rsSettings.MoveNext
    Loop
    rsSettings.Close
End If
Set rsSettings = Nothing

' 确保所有必需的键都存在
dim requiredKeys(5)
requiredKeys(0) = "ShowFixedSection"
requiredKeys(1) = "ShowCustomSection"
requiredKeys(2) = "ShowKOLSection"
requiredKeys(3) = "MinTopPercent"
requiredKeys(4) = "MinMiddlePercent"
requiredKeys(5) = "MinBasePercent"

For i = 0 To UBound(requiredKeys)
    If Not dictSettings.Exists(requiredKeys(i)) Then
        dictSettings.Add requiredKeys(i), IIF(InStr(requiredKeys(i), "Percent") > 0, "10", "0")
    End If
Next

' 获取各类型商品数量统计
Dim fixedCount, customCount, kolCount
fixedCount = GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType='Fixed' AND IsActive<>0")
customCount = GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType='Custom' AND IsActive<>0")
kolCount = GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType='KOL' AND IsActive<>0 AND ReviewStatus='Approved'")
If Not IsNumeric(fixedCount) Then fixedCount = 0
If Not IsNumeric(customCount) Then customCount = 0
If Not IsNumeric(kolCount) Then kolCount = 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>站点设置 - 香氛定制电商网站</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .settings-grid {
            display: grid;
            gap: 20px;
            margin-top: 20px;
        }
        .setting-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border: 2px solid rgba(255,255,255,0.08);
            border-radius: 8px;
            padding: 25px;
            transition: all 0.3s;
        }
        .setting-card:hover {
            border-color: #4CAF50;
            box-shadow: 0 2px 8px rgba(76, 175, 80, 0.2);
        }
        .setting-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }
        .setting-title {
            font-size: 18px;
            font-weight: bold;
            color: #e0e0e0;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .setting-badge {
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: bold;
        }
        .badge-fixed {
            background: rgba(25, 118, 210, 0.2);
            color: #64b5f6;
            border: 1px solid rgba(25, 118, 210, 0.3);
        }
        .badge-custom {
            background: rgba(245, 124, 0, 0.2);
            color: #ffb74d;
            border: 1px solid rgba(245, 124, 0, 0.3);
        }
        .badge-kol {
            background: rgba(123, 31, 162, 0.2);
            color: #ce93d8;
            border: 1px solid rgba(123, 31, 162, 0.3);
        }
        .setting-desc {
            color: #888;
            font-size: 14px;
            margin-bottom: 15px;
            line-height: 1.6;
        }
        .setting-stats {
            display: flex;
            gap: 20px;
            margin-bottom: 15px;
            padding: 10px;
            background: rgba(255,255,255,0.05);
            border-radius: 4px;
        }
        .stat-item {
            display: flex;
            align-items: center;
            gap: 8px;
            color: #b0b0b0;
        }
        .stat-item i {
            color: #4CAF50;
        }
        .toggle-switch {
            position: relative;
            display: inline-block;
            width: 60px;
            height: 30px;
        }
        .toggle-switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }
        .slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #555;
            transition: .4s;
            border-radius: 30px;
        }
        .slider:before {
            position: absolute;
            content: "";
            height: 22px;
            width: 22px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }
        input:checked + .slider {
            background-color: #4CAF50;
        }
        input:checked + .slider:before {
            transform: translateX(30px);
        }
        .status-indicator {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 6px 12px;
            border-radius: 4px;
            font-size: 13px;
            font-weight: bold;
        }
        .status-on {
            background: rgba(46, 125, 50, 0.2);
            color: #81c784;
        }
        .status-off {
            background: rgba(198, 40, 40, 0.2);
            color: #ef9a9a;
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success">
            <i class="fas fa-check-circle"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
            <br><small>当前状态 - 品牌定香: <%= IIF(dictSettings.Item("ShowFixedSection") = "1", "开启", "关闭") %> | 定制香水: <%= IIF(dictSettings.Item("ShowCustomSection") = "1", "开启", "关闭") %> | KOL推荐: <%= IIF(dictSettings.Item("ShowKOLSection") = "1", "开启", "关闭") %></small>
        </div>
        <% End If %>
        
        <div class="admin-card">
            <div class="admin-card-header">
                <h2 class="admin-card-title"><i class="fas fa-cog"></i> 站点设置</h2>
                <p style="color: #666; font-size: 14px; margin-top: 10px;">
                    配置网站基础参数和香调比例限制
                </p>
            </div>
            
            <div class="admin-card-body">
                <form method="post" action="site_settings.asp">
                    <input type="hidden" name="action" value="update">
                    
                    <!-- 保留隐藏字段以维持向后兼容 -->
                    <input type="hidden" name="showFixed" value="<%= dictSettings.Item("ShowFixedSection") %>">
                    <input type="hidden" name="showCustom" value="<%= dictSettings.Item("ShowCustomSection") %>">
                    <input type="hidden" name="showKOL" value="<%= dictSettings.Item("ShowKOLSection") %>">
                    
                    <!-- 商品类型管理迁移提示 -->
                    <div style="background: #e3f2fd; border: 1px solid #90caf9; border-radius: 8px; padding: 15px 20px; margin: 15px 0;">
                        <i class="fas fa-info-circle" style="color: #1976d2; margin-right: 8px;"></i>
                        <strong>提示：</strong>商品类型管理（包括栏目显示名称、开关控制等）已迁移至运营管理中心。
                        <br><br>
                        <a href="/admin/operation/product_types.asp" style="color: #1976d2; text-decoration: underline;">
                            <i class="fas fa-external-link-alt"></i> 前往运营管理中心 &gt; 商品类型 进行管理
                        </a>
                    </div>
                    
                    <!-- 香调配比参数迁移提示 -->
                    <div style="background: #fff3e0; border: 1px solid #ff9800; border-radius: 8px; padding: 15px 20px; margin: 15px 0;">
                        <i class="fas fa-exclamation-triangle" style="color: #f57c00; margin-right: 8px;"></i>
                        <strong>重要提示：</strong>香调最小比例设置已迁移至<strong>产品技术管理中心</strong>。
                        <br><br>
                        <a href="/admin/techcenter/product_settings.asp?tab=ratio" style="color: #f57c00; text-decoration: underline;">
                            <i class="fas fa-external-link-alt"></i> 前往产品技术管理中心 &gt; 香调配比参数 进行管理
                        </a>
                        <br><br>
                        <span style="color: #666; font-size: 13px;">
                            当前设置：前调 <%= dictSettings.Item("MinTopPercent") %>% | 中调 <%= dictSettings.Item("MinMiddlePercent") %>% | 后调 <%= dictSettings.Item("MinBasePercent") %>
                        </span>
                    </div>
                    
                    <!-- 保留隐藏字段以维持向后兼容 -->
                    <input type="hidden" name="minTopPercent" value="<%= dictSettings.Item("MinTopPercent") %>">
                    <input type="hidden" name="minMiddlePercent" value="<%= dictSettings.Item("MinMiddlePercent") %>">
                    <input type="hidden" name="minBasePercent" value="<%= dictSettings.Item("MinBasePercent") %>">
                    
                    <div style="margin-top: 20px; text-align: center;">
                        <button type="submit" class="admin-btn admin-btn-primary" style="padding: 12px 40px; font-size: 16px;">
                            <i class="fas fa-save"></i> 保存设置
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection
%>
