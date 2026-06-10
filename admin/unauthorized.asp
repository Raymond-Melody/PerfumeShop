<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>无访问权限 - 香氛电商系统</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container { 
            background: white;
            border-radius: 20px;
            padding: 60px 50px;
            text-align: center;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 500px;
            width: 100%;
        }
        .icon {
            width: 120px;
            height: 120px;
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 30px;
            font-size: 60px;
            color: white;
        }
        h1 {
            font-size: 28px;
            color: #333;
            margin-bottom: 15px;
        }
        p {
            color: #666;
            font-size: 16px;
            line-height: 1.6;
            margin-bottom: 30px;
        }
        .module-name {
            background: #f5f5f5;
            padding: 10px 20px;
            border-radius: 8px;
            display: inline-block;
            margin: 10px 0;
            color: #f5576c;
            font-weight: bold;
        }
        .contact-info {
            margin-top: 30px;
            padding-top: 30px;
            border-top: 1px solid #eee;
            color: #999;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="icon">
            <i class="fas fa-lock"></i>
        </div>
        <h1>无访问权限</h1>
        <p>
            抱歉，您没有权限访问此功能模块。<br>
            <% If Request.QueryString("module") <> "" Then %>
            尝试访问的模块：
            <div class="module-name"><%= Server.HTMLEncode(Request.QueryString("module")) %></div>
            <% End If %>
        </p>
        <div>
            <a href="portal.asp" class="btn"><i class="fas fa-home"></i> 返回统一入口</a>
            <a href="logout.asp" class="btn btn-secondary"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
        </div>
        <div class="contact-info">
            <p><i class="fas fa-info-circle"></i> 如需访问此功能，请联系系统管理员</p>
        </div>
    </div>
</body>
</html>
