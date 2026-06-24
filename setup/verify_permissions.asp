<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PerfumeShop - 数据库权限验证工具</title>
<style>
body{font-family:'Segoe UI',Arial,sans-serif;max-width:1200px;margin:0 auto;padding:20px;background:#f0f2f5}
h1{color:#1a237e;border-bottom:3px solid #1a237e;padding-bottom:10px}
.card{background:#fff;border-radius:8px;padding:20px;margin:15px 0;box-shadow:0 2px 4px rgba(0,0,0,0.1)}
.card h2{color:#283593;margin-top:0;border-bottom:2px solid #e3f2fd;padding-bottom:10px}
table{width:100%;border-collapse:collapse;margin:10px 0}
th,td{padding:10px;text-align:left;border-bottom:1px solid #e0e0e0}
th{background:#e3f2fd;color:#1a237e;font-weight:600}
tr:hover{background:#f5f5f5}
.badge{display:inline-block;padding:4px 12px;border-radius:12px;font-size:12px;font-weight:bold}
.badge-success{background:#e8f5e9;color:#2e7d32}
.badge-warning{background:#fff3e0;color:#e65100}
.badge-error{background:#ffebee;color:#c62828}
.badge-info{background:#e3f2fd;color:#1565c0}
.permission-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:15px;margin:15px 0}
.permission-item{padding:15px;border-radius:6px;border-left:4px solid}
.permission-granted{background:#e8f5e9;border-color:#4caf50}
.permission-denied{background:#ffebee;border-color:#f44336}
.permission-partial{background:#fff3e0;border-color:#ff9800}
.code{background:#f5f5f5;padding:2px 6px;border-radius:3px;font-family:'Courier New',monospace;font-size:12px}
.alert{padding:15px;border-radius:6px;margin:15px 0}
.alert-warning{background:#fff3e0;border-left:4px solid #ff9800;color:#e65100}
.alert-success{background:#e8f5e9;border-left:4px solid #4caf50;color:#2e7d32}
.alert-info{background:#e3f2fd;border-left:4px solid #2196f3;color:#1565c0}
.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:bold;margin:5px}
.btn-primary{background:#1a237e;color:#fff}
.btn-success{background:#2e7d32;color:#fff}
.btn-warning{background:#ff9800;color:#fff}
</style>
</head>
<body>

<h1>🔐 PerfumeShop 数据库权限验证工具</h1>
<p style="color:#666">检查当前 IIS 应用池身份的 SQL Server 数据库权限配置</p>

<%
Dim conn, rs
Dim currentUser, hasError
hasError = False

' 连接数据库
On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"

If Err.Number <> 0 Then
    hasError = True
    Response.Write "<div class='alert alert-warning'>"
    Response.Write "<strong>❌ 数据库连接失败</strong><br>"
    Response.Write "错误信息: " & Server.HTMLEncode(Err.Description) & "<br>"
    Response.Write "请确认 SQL Server 服务已启动，然后<a href='deploy.asp'>运行部署工具</a>"
    Response.Write "</div>"
    Response.End
End If
On Error GoTo 0

' 获取当前用户
Set rs = conn.Execute("SELECT SUSER_NAME() AS CurrentUser, USER_NAME() AS DatabaseUser")
If Not rs.EOF Then
    currentUser = rs.Fields("CurrentUser").Value
    Dim dbUser : dbUser = rs.Fields("DatabaseUser").Value
End If
rs.Close : Set rs = Nothing
%>

<div class="card">
    <h2>👤 当前身份信息</h2>
    <table>
        <tr>
            <th width="200">属性</th>
            <th>值</th>
        </tr>
        <tr>
            <td>SQL Server 登录名</td>
            <td><code class="code"><%= Server.HTMLEncode(currentUser) %></code></td>
        </tr>
        <tr>
            <td>数据库用户名</td>
            <td><code class="code"><%= Server.HTMLEncode(dbUser) %></code></td>
        </tr>
        <tr>
            <td>数据库名称</td>
            <td><code class="code">PerfumeShop</code></td>
        </tr>
        <tr>
            <td>SQL Server 版本</td>
            <td>
            <%
            Set rs = conn.Execute("SELECT @@VERSION AS VersionInfo")
            If Not rs.EOF Then
                Dim versionInfo : versionInfo = rs.Fields("VersionInfo").Value
                Response.Write "<code class='code'>" & Server.HTMLEncode(Left(versionInfo, 80)) & "...</code>"
            End If
            rs.Close : Set rs = Nothing
            %>
            </td>
        </tr>
    </table>
</div>

<div class="card">
    <h2>🔑 数据库角色权限</h2>
    <div class="permission-grid">
    <%
    ' 检查 db_ddladmin
    Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_ddladmin') AS HasRole")
    Dim hasDDL : hasDDL = False
    If Not rs.EOF And rs.Fields("HasRole").Value = 1 Then hasDDL = True
    rs.Close : Set rs = Nothing
    
    If hasDDL Then
        Response.Write "<div class='permission-item permission-granted'>"
        Response.Write "<h3>✅ db_ddladmin</h3>"
        Response.Write "<p><strong>权限:</strong> CREATE/ALTER/DROP 表、索引、视图、存储过程等</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-success'>已授予</span></p>"
        Response.Write "</div>"
    Else
        Response.Write "<div class='permission-item permission-denied'>"
        Response.Write "<h3>❌ db_ddladmin</h3>"
        Response.Write "<p><strong>权限:</strong> CREATE/ALTER/DROP 表、索引、视图、存储过程等</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-error'>未授予</span></p>"
        Response.Write "</div>"
    End If
    
    ' 检查 db_datareader
    Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_datareader') AS HasRole")
    Dim hasReader : hasReader = False
    If Not rs.EOF And rs.Fields("HasRole").Value = 1 Then hasReader = True
    rs.Close : Set rs = Nothing
    
    If hasReader Then
        Response.Write "<div class='permission-item permission-granted'>"
        Response.Write "<h3>✅ db_datareader</h3>"
        Response.Write "<p><strong>权限:</strong> SELECT 所有用户表数据</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-success'>已授予</span></p>"
        Response.Write "</div>"
    Else
        Response.Write "<div class='permission-item permission-denied'>"
        Response.Write "<h3>❌ db_datareader</h3>"
        Response.Write "<p><strong>权限:</strong> SELECT 所有用户表数据</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-error'>未授予</span></p>"
        Response.Write "</div>"
    End If
    
    ' 检查 db_datawriter
    Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_datawriter') AS HasRole")
    Dim hasWriter : hasWriter = False
    If Not rs.EOF And rs.Fields("HasRole").Value = 1 Then hasWriter = True
    rs.Close : Set rs = Nothing
    
    If hasWriter Then
        Response.Write "<div class='permission-item permission-granted'>"
        Response.Write "<h3>✅ db_datawriter</h3>"
        Response.Write "<p><strong>权限:</strong> INSERT/UPDATE/DELETE 所有用户表数据</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-success'>已授予</span></p>"
        Response.Write "</div>"
    Else
        Response.Write "<div class='permission-item permission-denied'>"
        Response.Write "<h3>❌ db_datawriter</h3>"
        Response.Write "<p><strong>权限:</strong> INSERT/UPDATE/DELETE 所有用户表数据</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-error'>未授予</span></p>"
        Response.Write "</div>"
    End If
    
    ' 检查 db_backupoperator 或 BACKUP DATABASE 权限
    Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_backupoperator') AS HasRole")
    Dim hasBackupRole : hasBackupRole = False
    If Not rs.EOF And rs.Fields("HasRole").Value = 1 Then hasBackupRole = True
    rs.Close : Set rs = Nothing
    
    Set rs = conn.Execute("SELECT COUNT(*) AS HasPerm FROM sys.database_permissions p JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id WHERE dp.name = USER_NAME() AND p.permission_name = 'BACKUP DATABASE'")
    Dim hasBackupPerm : hasBackupPerm = False
    If Not rs.EOF And rs.Fields("HasPerm").Value > 0 Then hasBackupPerm = True
    rs.Close : Set rs = Nothing
    
    If hasBackupRole Or hasBackupPerm Then
        Response.Write "<div class='permission-item permission-granted'>"
        Response.Write "<h3>✅ 备份权限</h3>"
        Response.Write "<p><strong>权限:</strong> BACKUP DATABASE / BACKUP LOG</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-success'>已授予</span></p>"
        Response.Write "</div>"
    Else
        Response.Write "<div class='permission-item permission-partial'>"
        Response.Write "<h3>⚠️ 备份权限</h3>"
        Response.Write "<p><strong>权限:</strong> BACKUP DATABASE / BACKUP LOG</p>"
        Response.Write "<p><strong>状态:</strong> <span class='badge badge-warning'>未授予（可选）</span></p>"
        Response.Write "</div>"
    End If
    %>
    </div>
</div>

<div class="card">
    <h2>📊 权限功能测试</h2>
    <table>
        <tr>
            <th width="200">测试项目</th>
            <th>测试操作</th>
            <th width="100">结果</th>
        </tr>
        <%
        ' 测试 1: 读取数据
        Dim testRead : testRead = False
        On Error Resume Next
        Set rs = conn.Execute("SELECT TOP 1 * FROM sys.tables")
        If Err.Number = 0 Then testRead = True
        If Not rs Is Nothing Then rs.Close : Set rs = Nothing
        Err.Clear
        
        Response.Write "<tr>"
        Response.Write "<td>读取权限测试</td>"
        Response.Write "<td><code class='code'>SELECT TOP 1 * FROM sys.tables</code></td>"
        If testRead Then
            Response.Write "<td><span class='badge badge-success'>✓ 通过</span></td>"
        Else
            Response.Write "<td><span class='badge badge-error'>✗ 失败</span></td>"
        End If
        Response.Write "</tr>"
        
        ' 测试 2: 创建测试表（DDL）
        Dim testDDL : testDDL = False
        conn.Execute "IF OBJECT_ID('PermissionTest_Temp') IS NOT NULL DROP TABLE PermissionTest_Temp"
        conn.Execute "CREATE TABLE PermissionTest_Temp (ID INT IDENTITY(1,1) PRIMARY KEY, TestName NVARCHAR(50))"
        If Err.Number = 0 Then testDDL = True
        Err.Clear
        
        Response.Write "<tr>"
        Response.Write "<td>DDL 权限测试</td>"
        Response.Write "<td><code class='code'>CREATE TABLE PermissionTest_Temp</code></td>"
        If testDDL Then
            Response.Write "<td><span class='badge badge-success'>✓ 通过</span></td>"
        Else
            Response.Write "<td><span class='badge badge-error'>✗ 失败</span></td>"
        End If
        Response.Write "</tr>"
        
        ' 测试 3: 插入数据
        Dim testInsert : testInsert = False
        If testDDL Then
            conn.Execute "INSERT INTO PermissionTest_Temp (TestName) VALUES ('Permission Test')"
            If Err.Number = 0 Then testInsert = True
            Err.Clear
        End If
        
        Response.Write "<tr>"
        Response.Write "<td>写入权限测试</td>"
        Response.Write "<td><code class='code'>INSERT INTO PermissionTest_Temp</code></td>"
        If testInsert Then
            Response.Write "<td><span class='badge badge-success'>✓ 通过</span></td>"
        Else
            Response.Write "<td><span class='badge badge-warning'>⊘ 跳过</span></td>"
        End If
        Response.Write "</tr>"
        
        ' 测试 4: 删除测试表
        If testDDL Then
            conn.Execute "DROP TABLE PermissionTest_Temp"
            Err.Clear
        End If
        %>
    </table>
</div>

<div class="card">
    <h2>📋 所有 IIS 身份权限概览</h2>
    <table>
        <tr>
            <th>IIS 身份</th>
            <th>db_ddladmin</th>
            <th>db_datareader</th>
            <th>db_datawriter</th>
            <th>BACKUP</th>
        </tr>
        <%
        Dim iisUsers(2)
        iisUsers(0) = "NT AUTHORITY\IUSR"
        iisUsers(1) = "IIS APPPOOL\DefaultAppPool"
        iisUsers(2) = "NT AUTHORITY\NETWORK SERVICE"
        
        Dim i, userName
        For i = 0 To 2
            userName = iisUsers(i)
            Response.Write "<tr>"
            Response.Write "<td><code class='code'>" & Server.HTMLEncode(userName) & "</code></td>"
            
            ' 检查用户是否存在
            Set rs = conn.Execute("SELECT COUNT(*) AS UserExists FROM sys.database_principals WHERE name='" & Replace(userName, "'", "''") & "'")
            Dim userExists : userExists = False
            If Not rs.EOF And rs.Fields("UserExists").Value > 0 Then userExists = True
            rs.Close : Set rs = Nothing
            
            If Not userExists Then
                Response.Write "<td colspan='4' style='color:#999'>数据库用户不存在</td>"
            Else
                ' 检查 db_ddladmin
                Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_ddladmin', '" & Replace(userName, "'", "''") & "') AS HasRole")
                If Not rs.EOF And rs.Fields("HasRole").Value = 1 Then
                    Response.Write "<td><span class='badge badge-success'>✓</span></td>"
                Else
                    Response.Write "<td><span class='badge badge-error'>✗</span></td>"
                End If
                rs.Close : Set rs = Nothing
                
                ' 检查 db_datareader
                Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_datareader', '" & Replace(userName, "'", "''") & "') AS HasRole")
                If Not rs.EOF And rs.Fields("HasRole").Value = 1 Then
                    Response.Write "<td><span class='badge badge-success'>✓</span></td>"
                Else
                    Response.Write "<td><span class='badge badge-error'>✗</span></td>"
                End If
                rs.Close : Set rs = Nothing
                
                ' 检查 db_datawriter
                Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_datawriter', '" & Replace(userName, "'", "''") & "') AS HasRole")
                If Not rs.EOF And rs.Fields("HasRole").Value = 1 Then
                    Response.Write "<td><span class='badge badge-success'>✓</span></td>"
                Else
                    Response.Write "<td><span class='badge badge-error'>✗</span></td>"
                End If
                rs.Close : Set rs = Nothing
                
                ' 检查 BACKUP 权限
                Set rs = conn.Execute("SELECT COUNT(*) AS HasPerm FROM sys.database_permissions p JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id WHERE dp.name='" & Replace(userName, "'", "''") & "' AND p.permission_name = 'BACKUP DATABASE'")
                If Not rs.EOF And rs.Fields("HasPerm").Value > 0 Then
                    Response.Write "<td><span class='badge badge-success'>✓</span></td>"
                Else
                    Response.Write "<td><span class='badge badge-warning'>-</span></td>"
                End If
                rs.Close : Set rs = Nothing
            End If
            
            Response.Write "</tr>"
        Next
        %>
    </table>
</div>

<%
' 显示总结和建议
Dim allGranted : allGranted = hasDDL And hasReader And hasWriter

If allGranted Then
    Response.Write "<div class='alert alert-success'>"
    Response.Write "<h3>✅ 权限配置完美！</h3>"
    Response.Write "<p>当前用户拥有所有必需的数据库权限，可以正常使用所有功能。</p>"
    Response.Write "</div>"
Else
    Response.Write "<div class='alert alert-warning'>"
    Response.Write "<h3>⚠️ 权限配置不完整</h3>"
    Response.Write "<p>当前用户缺少部分数据库权限，可能影响某些功能。请使用以下方法之一修复：</p>"
    Response.Write "<ol>"
    Response.Write "<li><strong>一键修复（推荐）：</strong><br>"
    Response.Write "  在文件资源管理器中打开 <code>" & Server.MapPath(".") & "\grant_ddl_permission.ps1</code><br>"
    Response.Write "  → 右键'以管理员身份运行'</li>"
    Response.Write "<li><strong>完整权限 SQL 脚本：</strong><br>"
    Response.Write "  在 SSMS 中打开并执行 <a href='grant_full_permissions.sql'><code>setup/grant_full_permissions.sql</code></a></li>"
    Response.Write "<li><strong>重新运行部署工具：</strong><br>"
    Response.Write "  访问 <a href='deploy.asp?action=run'><code>setup/deploy.asp</code></a> 并执行完整部署</li>"
    Response.Write "</ol>"
    Response.Write "</div>"
End If
%>

<div class="card">
    <h2> 操作按钮</h2>
    <a href="deploy.asp"><button class="btn btn-primary">📦 返回部署工具</button></a>
    <a href="deploy.asp?action=run"><button class="btn btn-success">🚀 重新部署</button></a>
    <button class="btn btn-warning" onclick="alert('请在文件资源管理器中打开：\n\nf:\\\\网站制作\\\\网站\\\\网站二\\\\setup\\\\grant_full_permissions.sql\n\n然后在 SSMS 中执行此脚本')">📄 查看 SQL 脚本路径</button>
    <button class="btn btn-primary" onclick="window.location.reload()">🔄 刷新验证</button>
</div>

<%
conn.Close : Set conn = Nothing
%>

<div style="text-align:center;margin-top:30px;color:#666">
    <p>PerfumeShop 数据库权限验证工具 v2.0 | 最后更新: <%= Now() %></p>
</div>

</body>
</html>
