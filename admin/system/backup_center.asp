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

' =====================================
' 下载自动修复脚本
' =====================================
If Request.QueryString("action") = "download_bat" Then
    Dim fsBat : Set fsBat = Server.CreateObject("Scripting.FileSystemObject")
    Dim batFilePath : batFilePath = Server.MapPath("/database/auto_fix_backup.bat")
    If fsBat.FileExists(batFilePath) Then
        Response.Clear
        Response.ContentType = "application/octet-stream"
        Response.AddHeader "Content-Disposition", "attachment; filename=auto_fix_backup.bat"
        Dim batStream : Set batStream = fsBat.OpenTextFile(batFilePath, 1)
        Response.Write batStream.ReadAll
        batStream.Close
        Set batStream = Nothing
    Else
        Response.ContentType = "text/plain; charset=utf-8"
        Response.Write "错误：auto_fix_backup.bat 文件未生成，请先在页面中点击「检测修复」生成该文件。"
    End If
    Set fsBat = Nothing
    Response.End
End If

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                val = rs(0)
                rs.Close
            End If
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing : GetScalar = val
End Function

' =====================================
' 自助诊断函数
' =====================================

' 获取当前SQL登录名
Function GetCurrentLogin()
    Dim rsL, v : v = "未知"
    On Error Resume Next
    Set rsL = conn.Execute("SELECT SUSER_SNAME()")
    If Err.Number = 0 And Not rsL Is Nothing And Not rsL.EOF Then v = CStr(rsL(0)) : rsL.Close
    Set rsL = Nothing : Err.Clear : GetCurrentLogin = v
End Function

' 检查 usp_BackupDatabase 存储过程是否存在
Function SPExists()
    Dim rsS, v : v = False
    On Error Resume Next
    Set rsS = conn.Execute("SELECT COUNT(*) FROM sys.procedures WHERE name='usp_BackupDatabase'")
    If Err.Number = 0 And Not rsS Is Nothing And Not rsS.EOF Then v = (CInt(rsS(0)) > 0) : rsS.Close
    Set rsS = Nothing : Err.Clear : SPExists = v
End Function

' 尝试创建 usp_BackupDatabase（多层方法，自动降级）
Function TryCreateSP()
    Dim savedErr : savedErr = ""
    On Error Resume Next
    
    ' == 方法1: EXECUTE AS OWNER（以 dbo 身份运行绕过权限限制）==
    Err.Clear
    conn.Execute "IF OBJECT_ID('usp_BackupDatabase','P') IS NOT NULL DROP PROCEDURE usp_BackupDatabase"
    If Err.Number <> 0 Then Err.Clear
    
    conn.Execute "CREATE PROCEDURE usp_BackupDatabase @backupPath NVARCHAR(500), @backupFile NVARCHAR(200), @dbName NVARCHAR(100)='PerfumeShop' WITH EXECUTE AS OWNER AS BEGIN SET NOCOUNT ON; DECLARE @sql NVARCHAR(2000); DECLARE @fullPath NVARCHAR(700); SET @fullPath = @backupPath + '\' + @backupFile; SET @sql = 'BACKUP DATABASE [' + @dbName + '] TO DISK = N''' + @fullPath + ''' WITH INIT, NAME = N''自动备份-'' + @backupFile; EXEC sp_executesql @sql END"
    If Err.Number = 0 Then
        Err.Clear : conn.Execute "GRANT EXECUTE ON usp_BackupDatabase TO public"
        If Err.Number = 0 Then TryCreateSP = True : Exit Function
    Else
        savedErr = Err.Description
    End If
    
    ' == 方法2: 先授予 BACKUP DATABASE + CREATE PROCEDURE 权限再创建 ==
    Err.Clear
    Dim loginName : loginName = GetCurrentLogin()
    If loginName <> "" And loginName <> "未知" Then
        conn.Execute "USE [master]; GRANT BACKUP DATABASE TO [" & Replace(loginName, "]", "]]") & "]"
        Err.Clear : conn.Execute "USE [PerfumeShop]; GRANT BACKUP DATABASE TO [" & Replace(loginName, "]", "]]") & "]"
        Err.Clear : conn.Execute "USE [PerfumeShop]; GRANT CREATE PROCEDURE TO [" & Replace(loginName, "]", "]]") & "]"
        Err.Clear
        conn.Execute "IF OBJECT_ID('usp_BackupDatabase','P') IS NOT NULL DROP PROCEDURE usp_BackupDatabase"
        Err.Clear
        conn.Execute "CREATE PROCEDURE usp_BackupDatabase @backupPath NVARCHAR(500), @backupFile NVARCHAR(200), @dbName NVARCHAR(100)='PerfumeShop' WITH EXECUTE AS OWNER AS BEGIN SET NOCOUNT ON; DECLARE @sql NVARCHAR(2000); DECLARE @fullPath NVARCHAR(700); SET @fullPath = @backupPath + '\' + @backupFile; SET @sql = 'BACKUP DATABASE [' + @dbName + '] TO DISK = N''' + @fullPath + ''' WITH INIT, NAME = N''自动备份-'' + @backupFile; EXEC sp_executesql @sql END"
        If Err.Number = 0 Then
            Err.Clear : conn.Execute "GRANT EXECUTE ON usp_BackupDatabase TO public"
            If Err.Number = 0 Then TryCreateSP = True : Exit Function
        End If
        Err.Clear
    End If
    
    ' 所有方法均失败
    TryCreateSP = False
End Function

' 备份配置
Dim backupPath : backupPath = Server.MapPath("/database/backups/")
Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(backupPath) Then fso.CreateFolder backupPath

' 执行备份
Dim msg : msg = ""
If Request.Form("action") = "backup_now" Then
    On Error Resume Next
    Dim bakFile : bakFile = "backup_" & Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & "_" & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2) & ".bak"
    Dim bakFullPath : bakFullPath = backupPath & "\" & bakFile
    
    ' 获取数据库名称
    Dim dbName : dbName = ""
    Set rs = conn.Execute("SELECT DB_NAME() AS DBName")
    If Not rs Is Nothing And Not rs.EOF Then dbName = rs("DBName") : rs.Close
    Set rs = Nothing
    
    If dbName <> "" Then
        ' 优先使用存储过程（EXECUTE AS OWNER 旁路权限限制）
        If SPExists() Then
            conn.Execute "EXEC usp_BackupDatabase N'" & Replace(backupPath, "'", "''") & "', N'" & Replace(bakFile, "'", "''") & "', N'" & Replace(dbName, "'", "''") & "'"
            If Err.Number = 0 Then
                msg = "备份成功：" & bakFile
            Else
                Err.Clear
                conn.Execute "BACKUP DATABASE [" & dbName & "] TO DISK = N'" & bakFullPath & "' WITH INIT, NAME = N'V8手动备份-" & bakFile & "'"
                If Err.Number = 0 Then msg = "备份成功：" & bakFile Else msg = "备份失败：" & Err.Description
            End If
        Else
            conn.Execute "BACKUP DATABASE [" & dbName & "] TO DISK = N'" & bakFullPath & "' WITH INIT, NAME = N'V8手动备份-" & bakFile & "'"
            If Err.Number = 0 Then msg = "备份成功：" & bakFile Else msg = "备份失败：" & Err.Description
        End If
    Else
        msg = "无法获取数据库名称"
    End If
    On Error GoTo 0
End If

' 处理自助修复请求（创建存储过程）
If Request.Form("fix_action") = "create_sp" Then
    On Error Resume Next
    Dim spOK : spOK = TryCreateSP()
    If spOK Then
        msg = "存储过程创建成功！请再次点击「立即备份」"
    Else
        Dim loginInfo : loginInfo = GetCurrentLogin()
        Dim spFailMsg : spFailMsg = "创建失败"
        If Err.Number <> 0 Then spFailMsg = spFailMsg & "（" & Trim(Err.Description) & "）"
        Err.Clear
        spFailMsg = spFailMsg & "。当前SQL登录【" & loginInfo & "】没有 CREATE PROCEDURE 权限，"
        spFailMsg = spFailMsg & "请下载自动修复脚本并以管理员身份运行"
        msg = spFailMsg
    End If
    On Error GoTo 0
End If

' 列出备份文件
Dim bakFiles(), bakCount : bakCount = 0
If fso.FolderExists(backupPath) Then
    Dim f
    For Each f In fso.GetFolder(backupPath).Files
        If LCase(fso.GetExtensionName(f.Name)) = "bak" Then
            ReDim Preserve bakFiles(bakCount)
            bakFiles(bakCount) = Array(f.Name, FormatNumber(f.Size / 1048576, 2) & " MB", f.DateLastModified, f.Path)
            bakCount = bakCount + 1
        End If
    Next
End If

' 数据库大小
Dim dbSize : dbSize = GetScalar("SELECT SUM(size)*8/1024 FROM sys.database_files WHERE type=0")
Dim dbLogSize : dbLogSize = GetScalar("SELECT SUM(size)*8/1024 FROM sys.database_files WHERE type=1")

' =====================================
' V10.4: PowerShell 手动备份触发
' =====================================
Dim psMsg : psMsg = ""
If Request.Form("action") = "ps_backup" Then
    On Error Resume Next
    Dim psScript : psScript = Server.MapPath("/database/backup_database.ps1")
    Dim psBackupDir : psBackupDir = Server.MapPath("/database/backups/")
    Dim wsShell : Set wsShell = Server.CreateObject("WScript.Shell")
    If Not wsShell Is Nothing Then
        Dim psCmd : psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & psScript & """ -BackupDir """ & psBackupDir & """ -RetentionDays " & BACKUP_RETENTION_DAYS
        Dim psExec : psExec = wsShell.Run(psCmd, 0, True)
        If psExec = 0 Then
            psMsg = "PowerShell 备份任务已执行完成。"
        Else
            psMsg = "PowerShell 备份返回代码: " & psExec & "，请检查服务器日志。"
        End If
        Set wsShell = Nothing
    Else
        psMsg = "无法创建 WScript.Shell 对象（可能被安全策略禁用）。请通过任务计划程序手动运行。"
    End If
    On Error GoTo 0
End If

' =====================================
' V10.4: 最近7天备份时间线
' =====================================
Dim timeline(6), tlIdx, tlDate, tlFound, f2
tlIdx = 0
Do While tlIdx < 7
    tlDate = DateAdd("d", -tlIdx, Date())
    tlFound = False
    If bakCount > 0 Then
        For Each f2 In fso.GetFolder(backupPath).Files
            If LCase(fso.GetExtensionName(f2.Name)) = "bak" Then
                If DateValue(f2.DateLastModified) = DateValue(tlDate) Then
                    tlFound = True
                    Exit For
                End If
            End If
        Next
    End If
    timeline(tlIdx) = Array(tlDate, tlFound)
    tlIdx = tlIdx + 1
Loop

' =====================================
' V10.4: 备份成功率统计 (最近30天)
' =====================================
Dim successCount, totalAttempts, successRate
successCount = 0 : totalAttempts = 0
If bakCount > 0 Then
    For Each f2 In fso.GetFolder(backupPath).Files
        If LCase(fso.GetExtensionName(f2.Name)) = "bak" Then
            If DateDiff("d", f2.DateLastModified, Now()) <= 30 Then
                totalAttempts = totalAttempts + 1
                If f2.Size > 0 Then successCount = successCount + 1
            End If
        End If
    Next
End If
If totalAttempts > 0 Then
    successRate = Round((successCount / totalAttempts) * 100, 1)
Else
    successRate = 0
End If

' V10.4: 下次计划备份时间
Dim nextBackupTime
nextBackupTime = DateAdd("d", 1, Date()) & " 02:00 AM"
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>备份中心 - 站点技术管理</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI',Arial,sans-serif; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .msg { padding: 12px 20px; border-radius: 6px; margin-bottom: 16px; }
        .msg-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        .msg-error { background: rgba(244,67,54,0.15); color: #EF9A9A; border: 1px solid rgba(244,67,54,0.3); }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-card .num { font-size: 28px; font-weight: 700; color: #00bcd4; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 5px; }
        .backup-actions { display: flex; gap: 15px; margin-bottom: 25px; }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .info-item label { display: block; color: #888; font-size: 12px; margin-bottom: 3px; }
        .info-item .value { font-size: 16px; font-weight: 600; color: #e0e0e0; }
        @media (max-width: 768px) { .main-content { margin-left: 0; } .stats-grid { grid-template-columns: 1fr 1fr; } .info-grid { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-database"></i> 备份中心</h2>
        <div class="breadcrumb"><a href="index.asp">系统中心</a> / 备份中心</div>
    </div>
    <% If msg <> "" Then %>
    <div class="msg <%= IIf(InStr(msg,"失败")>0,"msg-error","msg-success") %>"><i class="fas fa-info-circle"></i> <%= msg %></div>
    <% End If %>
    <% If psMsg <> "" Then %>
    <div class="msg <%= IIf(InStr(psMsg,"失败")>0 Or InStr(psMsg,"无法")>0,"msg-error","msg-success") %>"><i class="fas fa-terminal"></i> <%= psMsg %></div>
    <% End If %>
    
    <div class="stats-grid">
        <div class="stat-card"><div class="num"><%= dbSize %></div><div class="label">数据库大小(MB)</div></div>
        <div class="stat-card"><div class="num"><%= dbLogSize %></div><div class="label">日志大小(MB)</div></div>
        <div class="stat-card"><div class="num"><%= bakCount %></div><div class="label">备份文件数</div></div>
        <div class="stat-card"><div class="num" style="color:<%= IIf(successRate>=80,"#4CAF50",IIf(successRate>=50,"#FF9800","#F44336")) %>;"><%= successRate %>%</div><div class="label">30天成功率</div></div>
        <div class="stat-card"><div class="num" style="font-size:18px;color:#FF9800;"><%= nextBackupTime %></div><div class="label">下次计划备份</div></div>
    </div>
    
    <div class="backup-actions">
        <form method="post" style="display:inline;">
            <input type="hidden" name="action" value="backup_now">
            <button type="submit" class="btn btn-success" onclick="return confirm('确认立即执行数据库备份？')"><i class="fas fa-save"></i> 立即备份 (SQL)</button>
        </form>
        <form method="post" style="display:inline;" onsubmit="return confirm('将通过 PowerShell 执行完整备份（含验证），确认继续？');">
            <input type="hidden" name="action" value="ps_backup">
            <button type="submit" class="btn btn-warning" style="background:#FF9800;border-color:#FF9800;"><i class="fas fa-terminal"></i> PowerShell 备份</button>
        </form>
        <button class="btn btn-primary" onclick="document.getElementById('diagPanel').style.display=document.getElementById('diagPanel').style.display=='none'?'':'none'"><i class="fas fa-stethoscope"></i> 检测修复</button>
    </div>
    
    <!-- 自助诊断面板 -->
    <div id="diagPanel" class="card" style="border:1px solid rgba(0,188,212,0.2);display:none;">
        <div class="card-header"><i class="fas fa-stethoscope"></i> 系统诊断</div>
        <div class="card-body">
            <div class="info-grid">
                <div class="info-item"><label>当前SQL登录</label><div class="value" style="font-size:13px;"><%= GetCurrentLogin() %></div></div>
                <div class="info-item"><label>存储过程状态</label><div class="value" style="font-size:13px;"><%= IIf(SPExists(), "<span style='color:#4CAF50'>✓ 已存在（备份可用）</span>", "<span style='color:#EF9A9A'>✗ 未创建</span>") %></div></div>
                <div class="info-item"><label>备份路径</label><div class="value" style="font-size:13px;word-break:break-all;"><%= backupPath %></div></div>
            </div>
            <% If Not SPExists() Then %>
            <div style="margin-top:15px;padding:15px;background:rgba(0,188,212,0.08);border-radius:8px;border:1px solid rgba(0,188,212,0.2);">
                <p style="margin:0 0 8px;color:#00bcd4;"><i class="fas fa-lightbulb"></i> <strong>自助修复</strong></p>
                <p style="margin:0 0 12px;font-size:13px;color:#ccc;">缺少 usp_BackupDatabase 存储过程，该过程以数据库所有者(dbo)身份运行，可直接绕过 BACKUP DATABASE 权限限制。</p>
                <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="fix_action" value="create_sp">
                        <button type="submit" class="btn btn--primary"><i class="fas fa-wrench"></i> 一键创建存储过程</button>
                    </form>
                    <a href="?action=download_bat" class="btn btn--warning"><i class="fas fa-download"></i> 下载自动修复脚本</a>
                </div>
                <% If Not SPExists() Then %>
                <div style="margin-top:8px;padding:8px 12px;background:rgba(255,152,0,0.1);border-radius:4px;border:1px solid rgba(255,152,0,0.2);font-size:12px;color:#FFB74D;">
                    <i class="fas fa-info-circle"></i> 如果一键创建失败，请<strong>下载脚本 → 右键 → 以管理员身份运行</strong>，即可自动授予权限并创建存储过程
                </div>
                <% End If %>
            </div>
            <% End If %>
            <% If msg <> "" And InStr(msg, "失败") > 0 Then %>
            <div style="margin-top:12px;padding:12px;background:rgba(244,67,54,0.1);border-radius:6px;border:1px solid rgba(244,67,54,0.2);">
                <p style="margin:0;font-size:13px;color:#EF9A9A;"><i class="fas fa-exclamation-triangle"></i> 上次备份失败原因：<strong><%= msg %></strong></p>
            </div>
            <% End If %>
        </div>
    </div>
    
    <!-- V10.4: 最近7天备份时间线 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-calendar-check"></i> 最近7天备份时间线</div>
        <div class="card-body">
            <div class="timeline-bar">
                <% Dim tlI
                For tlI = 0 To 6
                    Dim tlItem : tlItem = timeline(tlI)
                    Dim tlDayName : tlDayName = WeekdayName(Weekday(tlItem(0)), True)
                    Dim tlDayNum : tlDayNum = Day(tlItem(0))
                    Dim tlHas : tlHas = tlItem(1)
                %>
                <div class="timeline-day <%= IIf(tlHas, "tl-success", "tl-miss") %>">
                    <div class="tl-icon"><i class="fas <%= IIf(tlHas, "fa-check-circle", "fa-times-circle") %>"></i></div>
                    <div class="tl-label"><%= tlDayName %></div>
                    <div class="tl-date"><%= tlDayNum %></div>
                </div>
                <% Next %>
            </div>
            <div style="display:flex;justify-content:center;gap:24px;margin-top:10px;font-size:12px;color:#888;">
                <span><span style="color:#4CAF50;">●</span> 已备份</span>
                <span><span style="color:#666;">●</span> 未备份</span>
            </div>
        </div>
    </div>
    <style>
        .timeline-bar { display: flex; justify-content: space-between; gap: 8px; }
        .timeline-day { flex: 1; text-align: center; padding: 12px 8px; border-radius: 8px; transition: all 0.3s; }
        .tl-success { background: rgba(76,175,80,0.12); border: 1px solid rgba(76,175,80,0.3); }
        .tl-miss { background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); }
        .tl-icon { font-size: 20px; margin-bottom: 6px; }
        .tl-success .tl-icon { color: #4CAF50; }
        .tl-miss .tl-icon { color: #555; }
        .tl-label { font-size: 12px; color: #999; font-weight: 600; }
        .tl-date { font-size: 16px; font-weight: 700; color: #e0e0e0; margin-top: 2px; }
        @media (max-width: 768px) { .timeline-bar { flex-wrap: wrap; } .timeline-day { min-width: 40px; } }
    </style>
    
    <!-- 备份配置信息 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-info-circle"></i> 备份配置</div>
        <div class="card-body">
            <div class="info-grid">
                <div class="info-item"><label>备份路径</label><div class="value" style="font-size:13px;word-break:break-all;"><%= backupPath %></div></div>
                <div class="info-item"><label>备份格式</label><div class="value">SQL Server .bak</div></div>
                <div class="info-item"><label>备份方式</label><div class="value">完整备份 (Full)</div></div>
                <div class="info-item"><label>建议频率</label><div class="value">每日自动备份</div></div>
            </div>
        </div>
    </div>
    
    <!-- 备份文件列表 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-archive"></i> 备份文件列表</div>
        <div class="card-body">
            <table>
                <thead><tr><th>文件名</th><th>大小</th><th>修改时间</th><th>操作</th></tr></thead>
                <tbody>
                    <% If bakCount > 0 Then
                        Dim bi, rowData
                        On Error Resume Next
                        For bi = 0 To bakCount - 1
                            Err.Clear
                            rowData = bakFiles(bi)
                            If IsArray(rowData) Then
                    %>
                    <tr>
                        <td><strong><%= rowData(0) %></strong></td>
                        <td><%= rowData(1) %></td>
                        <td><%= rowData(2) %></td>
                        <td>
                            <a href="/database/backups/<%= rowData(0) %>" class="btn" style="padding:5px 12px;font-size:12px;">下载</a>
                        </td>
                    </tr>
                    <%      End If
                        Next
                        On Error GoTo 0
                    Else %>
                    <tr><td colspan="4" style="text-align:center;padding:40px;color:#666;">暂无备份文件</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
    
    <div style="font-size:12px;color:#666;text-align:center;margin-top:20px;">
        <i class="fas fa-shield-alt"></i> 定时备份建议：在SQL Server Agent中创建作业，每日凌晨2:00执行完整备份
    </div>
</div>
</body>
</html>
<%
Set fso = Nothing
Call CloseConnection()
%>
