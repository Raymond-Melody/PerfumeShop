<%@ Language="VBScript" CodePage=65001 %>
<% Option Explicit %>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>PerfumeShop - 环境诊断 & 连接测试</title>
<style>
body{font-family:Arial,sans-serif;max-width:900px;margin:20px auto;padding:20px;background:#f5f5f5}
h1{color:#333;border-bottom:2px solid #2196F3;padding-bottom:10px}
h2{color:#555;margin-top:25px}
.pass{background:#d4edda;color:#155724;padding:4px 10px;border-radius:3px;font-weight:bold}
.fail{background:#f8d7da;color:#721c24;padding:4px 10px;border-radius:3px;font-weight:bold}
.warn{background:#fff3cd;color:#856404;padding:4px 10px;border-radius:3px;font-weight:bold}
table{width:100%;border-collapse:collapse;margin:10px 0;font-size:13px}
th,td{border:1px solid #ddd;padding:8px 12px;text-align:left}
th{background:#f2f2f2;width:35%}
pre{background:#fff;padding:10px;border:1px solid #ddd;border-radius:3px;overflow-x:auto;font-size:12px}
</style>
</head>
<body>
<h1>PerfumeShop 环境诊断报告</h1>
<p>生成时间: <%= Now() %></p>

<%
Dim testResult, diagInfo

' ==================== 诊断函数 ====================

Function TestClass(progID, label)
    On Error Resume Next
    Dim obj : Set obj = Server.CreateObject(progID)
    If Err.Number = 0 And Not obj Is Nothing Then
        TestClass = "<span class='pass'>已安装</span>"
        Set obj = Nothing
    Else
        TestClass = "<span class='fail'>未安装</span> (" & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function GetServerVar(name)
    GetServerVar = Request.ServerVariables(name)
End Function
%>

<h2>1. IIS 服务器信息</h2>
<table>
<tr><th>IIS 版本</th><td><%= GetServerVar("SERVER_SOFTWARE") %></td></tr>
<tr><th>服务器名称</th><td><%= GetServerVar("SERVER_NAME") %></td></tr>
<tr><th>服务器端口</th><td><%= GetServerVar("SERVER_PORT") %></td></tr>
<tr><th>脚本引擎</th><td><%= ScriptEngine & " v" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & " (Build " & ScriptEngineBuildVersion & ")" %></td></tr>
<tr><th>当前物理路径</th><td><%= Server.MapPath("./") %></td></tr>
<tr><th>应用程序池身份</th><td><%= GetServerVar("LOGON_USER") %></td></tr>
<tr><th>匿名用户</th><td><%= GetServerVar("AUTH_USER") %></td></tr>
<tr><th>请求方法</th><td><%= GetServerVar("REQUEST_METHOD") %></td></tr>
<tr><th>HTTPS</th><td><%= GetServerVar("HTTPS") %></td></tr>
</table>

<h2>2. ASP 关键配置检测</h2>
<table>
<tr><th>父路径启用</th>
    <td><%
        On Error Resume Next
        Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If Err.Number = 0 Then
            Response.Write "<span class='pass'>FSO 可用</span>"
        Else
            Response.Write "<span class='warn'>FSO 不可用</span>"
            Err.Clear
        End If
    %></td></tr>
<tr><th>Session 状态</th><td><%= IIf(Session.SessionID <> "", "<span class='pass'>正常 (ID: " & Session.SessionID & ")</span>", "<span class='fail'>不可用</span>") %></td></tr>
<tr><th>缓冲区</th><td><span class='pass'><%= IIf(Response.Buffer, "已启用", "未启用") %></span></td></tr>
<tr><th>脚本超时</th><td><%= Server.ScriptTimeout %> 秒</td></tr>
</table>

<h2>3. 数据库驱动检测</h2>
<table>
<tr><th>SQLOLEDB (SQL Server)</th><td><%= TestClass("ADODB.Connection", "ADO") %></td></tr>
<tr><th>ACE.OLEDB.12.0 (Access)</th><td><%
    On Error Resume Next
    Dim testConn : Set testConn = Server.CreateObject("ADODB.Connection")
    testConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("../database/PerfumeShop.mdb") & ";"
    If Err.Number = 0 Then
        Response.Write "<span class='pass'>已安装并可连接</span>"
        testConn.Close
    Else
        Response.Write "<span class='warn'>不可用 (" & Err.Description & ")</span>"
        Err.Clear
        ' 尝试 Jet
        testConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../database/PerfumeShop.mdb") & ";"
        If Err.Number = 0 Then
            Response.Write " | <span class='pass'>Jet.OLEDB.4.0 可用</span>"
            testConn.Close
        Else
            Response.Write " | <span class='fail'>Jet 也不可用</span>"
            Err.Clear
        End If
    End If
    Set testConn = Nothing
    On Error GoTo 0
%></td></tr>
</table>

<h2>4. SQL Server 数据库连接测试</h2>
<table>
<tr><th>SQL Server 连接</th>
    <td><%
        On Error Resume Next
        Dim connSQL
        Set connSQL = Server.CreateObject("ADODB.Connection")
        connSQL.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=master;Integrated Security=SSPI;"
        If Err.Number = 0 Then
            Response.Write "<span class='pass'>成功连接到 localhost</span>"
            ' 检查 PerfumeShop 数据库
            Dim rsDB : Set rsDB = connSQL.Execute("SELECT COUNT(*) FROM sys.databases WHERE name='PerfumeShop'")
            If Not rsDB.EOF Then
                If rsDB.Fields(0).Value > 0 Then
                    Response.Write "<br><span class='pass'>PerfumeShop 数据库: 存在</span>"
                    ' 检查表数量
                    connSQL.Execute "USE [PerfumeShop]"
                    Set rsDB = connSQL.Execute("SELECT COUNT(*) FROM sys.tables")
                    Response.Write "<br>表数量: " & rsDB.Fields(0).Value
                    rsDB.Close
                Else
                    Response.Write "<br><span class='fail'>PerfumeShop 数据库: 不存在</span>"
                End If
            End If
            rsDB.Close : Set rsDB = Nothing
            connSQL.Close
        Else
            Response.Write "<span class='fail'>连接失败: " & Err.Description & "</span>"
            Err.Clear
        End If
        Set connSQL = Nothing
        On Error GoTo 0
    %></td></tr>
<tr><th>SQL Server 版本</th>
    <td><%
        On Error Resume Next
        Set connSQL = Server.CreateObject("ADODB.Connection")
        connSQL.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=master;Integrated Security=SSPI;"
        If Err.Number = 0 Then
            Set rsDB = connSQL.Execute("SELECT @@VERSION")
            If Not rsDB.EOF Then
                Response.Write Server.HTMLEncode(rsDB.Fields(0).Value)
            End If
            rsDB.Close : Set rsDB = Nothing
            connSQL.Close
        End If
        Set connSQL = Nothing
        On Error GoTo 0
    %></td></tr>
<tr><th>当前连接用户</th>
    <td><%
        On Error Resume Next
        Set connSQL = Server.CreateObject("ADODB.Connection")
        connSQL.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=master;Integrated Security=SSPI;"
        If Err.Number = 0 Then
            Set rsDB = connSQL.Execute("SELECT SUSER_SNAME()")
            If Not rsDB.EOF Then
                Response.Write rsDB.Fields(0).Value
            End If
            rsDB.Close : Set rsDB = Nothing
            connSQL.Close
        End If
        Set connSQL = Nothing
        On Error GoTo 0
    %></td></tr>
</table>

<h2>5. web.config 安全头检测</h2>
<table>
<tr><th>X-Content-Type-Options</th><td><%= GetServerVar("HTTP_X_CONTENT_TYPE_OPTIONS") %></td></tr>
<tr><th>X-Frame-Options</th><td><%= GetServerVar("HTTP_X_FRAME_OPTIONS") %></td></tr>
<tr><th>X-XSS-Protection</th><td><%= GetServerVar("HTTP_X_XSS_PROTECTION") %></td></tr>
</table>

<h2>6. 核心文件检查</h2>
<%
Dim coreFiles, filePath, fileExists
coreFiles = Array( _
    "includes/connection.asp", _
    "includes/config.asp", _
    "includes/auth.asp", _
    "index.asp", _
    "admin/login.asp", _
    "admin/index.html", _
    "user/login.asp", _
    "user/register.asp", _
    "products.asp", _
    "product.asp", _
    "cart.asp", _
    "checkout.asp" _
)
%>
<table>
<%
For Each filePath In coreFiles
    Dim fullPath : fullPath = Server.MapPath("../" & filePath)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    fileExists = fso.FileExists(fullPath)
    Set fso = Nothing
%>
<tr><th><%= filePath %></th>
    <td><%= IIf(fileExists, "<span class='pass'>存在</span>", "<span class='fail'>缺失!</span>") %></td></tr>
<%
Next
%>
</table>

<h2>7. 部署建议</h2>
<div style="background:#e3f2fd;padding:15px;border-radius:5px;border:1px solid #90caf9">
<%
If GetServerVar("SERVER_PORT") = "80" Then
    Response.Write "<p><span class='pass'>站点似乎已在 80 端口运行</span></p>"
Else
    Response.Write "<p><span class='warn'>当前端口: " & GetServerVar("SERVER_PORT") & " (非标准80端口)</span></p>"
End If
%>
<p><strong>需要手动配置的 IIS 设置:</strong></p>
<ol>
    <li>打开 IIS 管理器 (inetmgr)</li>
    <li>确保网站指向: <code>f:\网站制作\网站\网站二</code></li>
    <li>应用程序池设置:
        <ul>
            <li>.NET CLR 版本: 无托管代码</li>
            <li>托管管道模式: Classic (经典)</li>
            <li>标识: ApplicationPoolIdentity 或 NetworkService</li>
        </ul>
    </li>
    <li>ASP 设置:
        <ul>
            <li>启用父路径: True</li>
            <li>调试属性 > 将错误发送到浏览器: True (仅开发环境)</li>
        </ul>
    </li>
    <li>运行 <a href="deploy.asp">deploy.asp</a> 一键部署数据库（建库+建表+种子数据+权限）</li>
    <li>运行 <a href="data_migrate.asp">data_migrate.asp</a> 从Access迁移现有数据</li>
</ol>
</div>

<%
Function IIf(cond, tv, fv)
    If cond Then IIf = tv Else IIf = fv
End Function
%>
</body>
</html>
