<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.Status = "500 Internal Server Error"
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>系统错误 - 香氛定制</title>
<style>
body{font-family:Arial,sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:#f5f5f5}
.error-container{max-width:500px;padding:40px;background:#fff;border-radius:8px;box-shadow:0 5px 20px rgba(0,0,0,0.1);text-align:center}
.error-container h1{color:#e74c3c;font-size:48px;margin:0 0 10px 0}
.error-container h2{color:#333;margin:0 0 20px 0}
.error-container p{color:#666;line-height:1.6;margin:0 0 20px 0}
.error-container .btn{display:inline-block;padding:10px 30px;background:#3498db;color:#fff;text-decoration:none;border-radius:4px;font-size:14px}
.error-container .btn:hover{background:#2980b9}
</style>
</head>
<body>
<div class="error-container">
    <h1>500</h1>
    <h2>服务器内部错误</h2>
    <p>很抱歉，系统遇到了一个意外错误。请稍后重试，或联系管理员。</p>
    <a href="/" class="btn">返回首页</a>
</div>
</body>
</html>
