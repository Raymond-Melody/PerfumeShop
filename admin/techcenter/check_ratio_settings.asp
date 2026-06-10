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
<html>
<head>
    <meta charset="UTF-8">
    <title>香调配比设置诊断</title>
</head>
<body>
    <h1>SiteSettings 表诊断</h1>
    
    <h2>1. 检查 MinTopPercent</h2>
    <% 
    Dim rs1
    Set rs1 = ExecuteQuery("SELECT * FROM SiteSettings WHERE SettingKey = 'MinTopPercent'")
    If Not rs1 Is Nothing Then
        If Not rs1.EOF Then
            Response.Write "<p style='color:green'>✓ 找到记录:</p>"
            Response.Write "<ul>"
            Response.Write "<li>SettingKey: " & rs1("SettingKey") & "</li>"
            Response.Write "<li>SettingValue: " & rs1("SettingValue") & "</li>"
            If Not IsNull(rs1("UpdatedAt")) Then
                Response.Write "<li>UpdatedAt: " & rs1("UpdatedAt") & "</li>"
            End If
            Response.Write "</ul>"
        Else
            Response.Write "<p style='color:red'>✗ 未找到 MinTopPercent 记录</p>"
        End If
        rs1.Close
    Else
        Response.Write "<p style='color:red'>✗ 查询失败</p>"
    End If
    %>
    
    <h2>2. 检查 MinMiddlePercent</h2>
    <% 
    Dim rs2
    Set rs2 = ExecuteQuery("SELECT * FROM SiteSettings WHERE SettingKey = 'MinMiddlePercent'")
    If Not rs2 Is Nothing Then
        If Not rs2.EOF Then
            Response.Write "<p style='color:green'>✓ 找到记录:</p>"
            Response.Write "<ul>"
            Response.Write "<li>SettingKey: " & rs2("SettingKey") & "</li>"
            Response.Write "<li>SettingValue: " & rs2("SettingValue") & "</li>"
            If Not IsNull(rs2("UpdatedAt")) Then
                Response.Write "<li>UpdatedAt: " & rs2("UpdatedAt") & "</li>"
            End If
            Response.Write "</ul>"
        Else
            Response.Write "<p style='color:red'>✗ 未找到 MinMiddlePercent 记录</p>"
        End If
        rs2.Close
    Else
        Response.Write "<p style='color:red'>✗ 查询失败</p>"
    End If
    %>
    
    <h2>3. 检查 MinBasePercent</h2>
    <% 
    Dim rs3
    Set rs3 = ExecuteQuery("SELECT * FROM SiteSettings WHERE SettingKey = 'MinBasePercent'")
    If Not rs3 Is Nothing Then
        If Not rs3.EOF Then
            Response.Write "<p style='color:green'>✓ 找到记录:</p>"
            Response.Write "<ul>"
            Response.Write "<li>SettingKey: " & rs3("SettingKey") & "</li>"
            Response.Write "<li>SettingValue: " & rs3("SettingValue") & "</li>"
            If Not IsNull(rs3("UpdatedAt")) Then
                Response.Write "<li>UpdatedAt: " & rs3("UpdatedAt") & "</li>"
            End If
            Response.Write "</ul>"
        Else
            Response.Write "<p style='color:red'>✗ 未找到 MinBasePercent 记录</p>"
        End If
        rs3.Close
    Else
        Response.Write "<p style='color:red'>✗ 查询失败</p>"
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
    
    <p><a href="product_settings.asp?tab=ratio">返回设置页面</a></p>
</body>
</html>
