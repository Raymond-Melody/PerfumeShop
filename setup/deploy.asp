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
<title>PerfumeShop - 系统重新部署工具</title>
<style>
body{font-family:'Segoe UI',Arial,sans-serif;max-width:1000px;margin:0 auto;padding:20px;background:#f0f2f5}
h1{color:#1a237e;border-bottom:3px solid #1a237e;padding-bottom:10px;margin-bottom:5px}
.subtitle{color:#666;margin-bottom:20px}
.stage{border-radius:8px;margin:10px 0;padding:0;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1)}
.stage-header{padding:12px 16px;font-weight:bold;font-size:15px;display:flex;align-items:center;gap:10px}
.stage-body{padding:12px 16px;background:#fff}
.stage-pending .stage-header{background:#e0e0e0;color:#555}
.stage-running .stage-header{background:#2196F3;color:#fff}
.stage-success .stage-header{background:#4CAF50;color:#fff}
.stage-error .stage-header{background:#f44336;color:#fff}
.stage-warning .stage-header{background:#FF9800;color:#fff}
.step{padding:6px 0;font-size:13px;border-bottom:1px solid #f5f5f5}
.step:last-child{border-bottom:none}
.step-success{color:#2e7d32}
.step-error{color:#c62828}
.step-info{color:#1565C0}
.step-warning{color:#e65100}
.status-icon{font-size:20px}
.spinner{display:inline-block;width:16px;height:16px;border:2px solid #fff;border-top-color:transparent;border-radius:50%;animation:spin 0.8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.btn{padding:12px 40px;font-size:16px;border:none;border-radius:6px;cursor:pointer;font-weight:bold;margin:10px 5px}
.btn-primary{background:#1a237e;color:#fff}
.btn-primary:hover{background:#283593}
.btn-danger{background:#c62828;color:#fff}
.btn-success{background:#2e7d32;color:#fff}
.btn:disabled{opacity:0.5;cursor:not-allowed}
.summary-box{border-radius:8px;padding:16px;margin:10px 0}
.summary-success{background:#e8f5e9;border:2px solid #4CAF50;color:#1b5e20}
.summary-warning{background:#fff3e0;border:2px solid #FF9800;color:#e65100}
pre{background:#263238;color:#aed581;padding:6px 10px;border-radius:4px;font-size:11px;overflow-x:auto;margin:4px 0}
.actions{text-align:center;padding:20px 0}
</style>
</head>
<body>

<h1>PerfumeShop 系统重新部署工具</h1>
<p class="subtitle">一站式部署：数据库 → 表结构 → 种子数据 → 权限 → 验证</p>

<div id="stages">
<%
' ============================================
' 部署主控制器
' ============================================
Dim deployMode
deployMode = Request.QueryString("action")

If deployMode = "run" Then
    Response.Write "<div id=""live-output"">"
    Call RunFullDeployment()
    Response.Write "</div>"
Else
    ' 显示预检信息和启动按钮
%>
<div class="summary-box summary-warning">
    <strong>预检提示：</strong>执行部署前，请确保：
    <ul style="margin:10px 0 0 20px">
        <li>SQL Server (MSSQLSERVER) 服务已启动</li>
        <li>IIS 网站已正确指向项目根目录：<code>f:\网站制作\网站\网站二</code></li>
        <li>应用程序池使用 <b>Classic</b> 管道模式</li>
        <li>ASP 父路径已启用</li>
    </ul>
</div>

<div class="summary-box" style="background:#e3f2fd;border:2px solid #2196F3;color:#0d47a1">
    <strong>部署流程：</strong>
    <ol style="margin:10px 0 0 20px;line-height:1.8">
        <li>连接 SQL Server（默认实例 localhost）</li>
        <li>创建 PerfumeShop 数据库</li>
        <li>创建全部 55+ 张数据表及索引</li>
        <li>插入种子数据（管理员账号、角色、站点配置等）</li>
        <li>配置数据库权限</li>
        <li>验证所有表结构完整性</li>
    </ol>
</div>

<div class="actions">
    <button class="btn btn-primary" onclick="startDeploy()">🚀 开始完整部署</button>
    <a href="data_migrate.asp" style="display:inline-block;margin-left:10px">
        <button class="btn btn-success" type="button">📦 数据迁移 (Access→SQL)</button>
    </a>
    <a href="env_check.asp" style="display:inline-block;margin-left:10px">
        <button class="btn btn-primary" type="button" style="background:#555">🔍 环境诊断</button>
    </a>
</div>

<script>
function startDeploy(){
    if(!confirm('确认要重新部署整个系统吗？\n\n这将创建/重建 PerfumeShop 数据库及所有表结构。\n如果数据库已存在，已有数据可能会被保留。')) return;
    window.location.href = 'deploy.asp?action=run';
}
</script>
<%
End If

' ============================================
' 日志输出函数
' ============================================
Sub Log(level, msg)
    Dim icon, cssClass
    Select Case level
        Case "success": icon = ChrW(&H2713) : cssClass = "step-success"
        Case "error":   icon = ChrW(&H2717) : cssClass = "step-error"
        Case "info":    icon = ChrW(&H25BA) : cssClass = "step-info"
        Case "warning": icon = ChrW(&H26A0) : cssClass = "step-warning"
        Case "stage":   icon = ""
    End Select
    Response.Write "<div class=""step " & cssClass & """>" & icon & " " & Server.HTMLEncode(msg) & "</div>"
    Response.Flush
End Sub

Sub StageStart(name)
    Response.Write "<div class=""stage stage-running""><div class=""stage-header""><span class=""spinner""></span> " & Server.HTMLEncode(name) & "</div><div class=""stage-body"">"
    Response.Flush
End Sub

Sub StageEnd(status)
    Dim headerClass, icon
    If status = "success" Then headerClass = "stage-success": icon = ChrW(&H2713)
    If status = "error" Then headerClass = "stage-error": icon = ChrW(&H2717)
    If status = "warning" Then headerClass = "stage-warning": icon = ChrW(&H26A0)
    Response.Write "</div></div>"
    Response.Write "<script>var s=document.querySelector('.stage-running');if(s){s.className='stage " & headerClass & "';s.querySelector('.stage-header').innerHTML='" & icon & " ' + s.querySelector('.stage-header').textContent.trim()}</script>"
    Response.Flush
End Sub

' ============================================
' 创建表（如果不存在）
' ============================================
Function CreateTableIfNotExists(conn, tableName, createSQL)
    On Error Resume Next
    Dim rs
    Set rs = conn.Execute("SELECT COUNT(*) FROM sys.tables WHERE name='" & tableName & "'")
    If Not rs.EOF Then
        If rs.Fields(0).Value > 0 Then
            rs.Close : Set rs = Nothing
            Call Log("info", tableName & " - 已存在，跳过")
            CreateTableIfNotExists = True
            Exit Function
        End If
    End If
    rs.Close : Set rs = Nothing
    
    conn.Execute createSQL
    If Err.Number <> 0 Then
        Call Log("error", tableName & " - 创建失败: " & Err.Description)
        Err.Clear
        CreateTableIfNotExists = False
    Else
        Call Log("success", tableName & " - 创建成功")
        CreateTableIfNotExists = True
    End If
End Function

' ============================================
' 安全执行 SQL
' ============================================
Function SafeExec(conn, sql, desc)
    On Error Resume Next
    conn.Execute sql
    If Err.Number <> 0 Then
        Call Log("warning", desc & " - " & Err.Description)
        Err.Clear
        SafeExec = False
    Else
        SafeExec = True
    End If
End Function

' ============================================
' 完整部署主流程
' ============================================
Sub RunFullDeployment()
    Dim conn, rs
    Dim totalTables : totalTables = 0
    Dim successTables : successTables = 0
    
    ' ==== Stage 1: 连接 SQL Server ====
    Call StageStart("Stage 1: 连接 SQL Server")
    
    On Error Resume Next
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=master;Integrated Security=SSPI;"
    
    If Err.Number <> 0 Then
        Call Log("error", "无法连接到 SQL Server: " & Err.Description)
        Call Log("error", "请确认 SQL Server (MSSQLSERVER) 服务已启动")
        Call StageEnd("error")
        Response.Write "<div class=""summary-box summary-warning""><strong>连接失败！</strong>请先启动 SQL Server 服务后重试。<br>可以通过 服务管理器 (services.msc) 查找 SQL Server (MSSQLSERVER) 并启动。</div>"
        Exit Sub
    End If
    Call Log("success", "成功连接到 SQL Server 默认实例 (localhost)")
    Call StageEnd("success")
    
    ' ==== Stage 2: 创建 PerfumeShop 数据库 ====
    Call StageStart("Stage 2: 创建 PerfumeShop 数据库")
    
    Dim dbExists : dbExists = False
    Set rs = conn.Execute("SELECT COUNT(*) FROM sys.databases WHERE name='PerfumeShop'")
    If Not rs.EOF And rs.Fields(0).Value > 0 Then dbExists = True
    rs.Close : Set rs = Nothing
    
    If dbExists Then
        Call Log("warning", "数据库 PerfumeShop 已存在")
        ' 删除重建选项 - 仅在明确需要时
        Dim forceRecreate
        forceRecreate = Request.QueryString("force")
        If forceRecreate = "1" Then
            conn.Execute "ALTER DATABASE [PerfumeShop] SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
            conn.Execute "DROP DATABASE [PerfumeShop]"
            Call Log("warning", "已删除旧数据库 PerfumeShop（强制重建模式）")
            dbExists = False
        Else
            Call Log("info", "将使用现有数据库（如需重建请添加 &force=1 参数）")
        End If
    End If
    
    If Not dbExists Then
        conn.Execute "CREATE DATABASE [PerfumeShop]"
        If Err.Number <> 0 Then
            Call Log("error", "创建数据库失败: " & Err.Description)
            Call StageEnd("error")
            Exit Sub
        End If
        Call Log("success", "已创建数据库 PerfumeShop")
        
        conn.Execute "ALTER DATABASE [PerfumeShop] SET RECOVERY SIMPLE"
        Call Log("success", "恢复模式已设置为 SIMPLE")
    End If
    Call StageEnd("success")
    
    ' 切换到 PerfumeShop
    conn.Execute "USE [PerfumeShop]"
    
    ' ==== Stage 2b: 授予当前用户 DDL 权限 ====
    Call StageStart("Stage 2b: 授予数据库权限（DDL + 数据读写 + 备份）")
    
    Dim currentUser
    Set rs = conn.Execute("SELECT SUSER_NAME()")
    If Not rs.EOF Then currentUser = rs.Fields(0).Value
    rs.Close : Set rs = Nothing
    
    Call Log("info", "当前连接用户: " & currentUser)
    
    ' 创建数据库用户（如果不存在）
    Call SafeExec(conn, "IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name='" & Replace(currentUser, "'", "''") & "') CREATE USER [" & currentUser & "] FOR LOGIN [" & currentUser & "]", "创建数据库用户")
    
    ' 授予 db_ddladmin（CREATE TABLE/INDEX/ALTER 等 DDL 权限）
    Dim ddlGranted, readerGranted, writerGranted
    ddlGranted = SafeExec(conn, "ALTER ROLE db_ddladmin ADD MEMBER [" & Replace(currentUser, "'", "''") & "]", "授予 db_ddladmin（DDL 权限）")
    readerGranted = SafeExec(conn, "ALTER ROLE db_datareader ADD MEMBER [" & Replace(currentUser, "'", "''") & "]", "授予 db_datareader（数据读取）")
    writerGranted = SafeExec(conn, "ALTER ROLE db_datawriter ADD MEMBER [" & Replace(currentUser, "'", "''") & "]", "授予 db_datawriter（数据写入）")
    
    ' 授予备份权限（可选）
    Call SafeExec(conn, "GRANT BACKUP DATABASE TO [" & Replace(currentUser, "'", "''") & "]", "授予 BACKUP DATABASE")
    Call SafeExec(conn, "GRANT BACKUP LOG TO [" & Replace(currentUser, "'", "''") & "]", "授予 BACKUP LOG")
    
    If ddlGranted Then
        Call Log("success", "DDL 权限授予完成（CREATE/ALTER/DROP 表、索引等）")
    Else
        Call Log("warning", "DDL 权限授予失败（可能需要手动授权）")
    End If
    
    If readerGranted Then
        Call Log("success", "数据读取权限授予完成（SELECT）")
    Else
        Call Log("warning", "数据读取权限授予失败")
    End If
    
    If writerGranted Then
        Call Log("success", "数据写入权限授予完成（INSERT/UPDATE/DELETE）")
    Else
        Call Log("warning", "数据写入权限授予失败")
    End If
    
    ' 如果 DDL 权限授予失败，显示详细的解决指南
    If Not ddlGranted Then
        Response.Write "<div style='background:#fff3e0;border:2px solid #FF9800;border-radius:6px;padding:16px;margin:10px 0;color:#e65100'>"
        Response.Write "<strong>⚠ 权限不足警告：</strong>当前 IIS 身份 (<code>" & Server.HTMLEncode(currentUser) & "</code>) 没有安全管理员权限，无法自动授予 DDL 权限。<br><br>"
        Response.Write "<strong>🔧 解决方案（三选一，按推荐顺序）：</strong><br>"
        Response.Write "<ol style='margin:10px 0 0 20px;line-height:2'>"
        Response.Write "<li><strong>一键修复（推荐）：</strong><br>"
        Response.Write "  在文件资源管理器中打开 <code>" & Server.MapPath("..") & "\setup\grant_ddl_permission.ps1</code><br>"
        Response.Write "  → 右键'以管理员身份运行' → 自动配置所有权限后刷新此页面</li>"
        Response.Write "<li><strong>完整权限 SQL 脚本：</strong><br>"
        Response.Write "  在文件资源管理器中打开 <code>" & Server.MapPath("..") & "\setup\grant_full_permissions.sql</code><br>"
        Response.Write "  → 在 SSMS 中打开并执行，自动为所有 IIS 身份配置完整权限（DDL + 数据读写 + 备份）</li>"
        Response.Write "<li><strong>SSMS 手动授权：</strong><br>"
        Response.Write "  在 SSMS 中执行 <code>ALTER ROLE db_ddladmin ADD MEMBER [" & Server.HTMLEncode(currentUser) & "]</code><br>"
        Response.Write "  → 然后再运行此部署工具</li>"
        Response.Write "</ol>"
        Response.Write "<hr style='margin:12px 0;border-color:#FF9800'>"
        Response.Write "<strong>📋 需要授予的权限清单：</strong><br>"
        Response.Write "<ul style='margin:8px 0 0 20px;line-height:1.8'>"
        Response.Write "<li><code>db_ddladmin</code> - 创建/修改/删除表、索引、视图等数据库对象</li>"
        Response.Write "<li><code>db_datareader</code> - 读取所有用户表中的数据</li>"
        Response.Write "<li><code>db_datawriter</code> - 插入/更新/删除所有用户表中的数据</li>"
        Response.Write "<li><code>BACKUP DATABASE</code> - 执行数据库完整备份（可选）</li>"
        Response.Write "<li><code>BACKUP LOG</code> - 执行事务日志备份（可选）</li>"
        Response.Write "</ul></div>"
    End If
    
    Call StageEnd(IIf(ddlGranted, "success", "warning"))
    
    ' ==== Stage 3: 创建所有表结构 ====
    Call StageStart("Stage 3: 创建数据表结构（共62+张表）")
    
    ' --- 核心用户与权限表 ---
    If CreateTableIfNotExists(conn, "Users", "CREATE TABLE [Users] ([UserID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [Username] NVARCHAR(50) NOT NULL, [Password] NVARCHAR(255) NOT NULL, [Email] NVARCHAR(100) NOT NULL, [FullName] NVARCHAR(100) NULL, [Phone] NVARCHAR(20) NULL, [Address] NVARCHAR(200) NULL, [City] NVARCHAR(50) NULL, [PostalCode] NVARCHAR(20) NULL, [Points] INT NULL DEFAULT 0, [IsActive] BIT NULL DEFAULT 1, [IsVIP] BIT NULL DEFAULT 0, [UserRole] NVARCHAR(20) NULL DEFAULT 'user', [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "AdminRoles", "CREATE TABLE [AdminRoles] ([RoleID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RoleCode] NVARCHAR(20) NOT NULL, [RoleName] NVARCHAR(50) NOT NULL, [Description] NVARCHAR(MAX) NULL, [Permissions] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "AdminUsers", "CREATE TABLE [AdminUsers] ([AdminID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [Username] NVARCHAR(50) NOT NULL, [PasswordHash] NVARCHAR(255) NOT NULL, [Email] NVARCHAR(100) NOT NULL, [FullName] NVARCHAR(100) NULL, [Department] NVARCHAR(50) NULL, [RoleID] INT NULL, [IsActive] BIT NULL DEFAULT 1, [IsLocked] BIT NULL DEFAULT 0, [LastLogin] DATETIME2(7) NULL, [ResetToken] NVARCHAR(255) NULL, [ResetTokenExpiry] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "AdminLogs", "CREATE TABLE [AdminLogs] ([LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [AdminID] INT NULL, [ActionType] NVARCHAR(100) NULL, [TableName] NVARCHAR(50) NULL, [RecordID] NVARCHAR(50) NULL, [ModuleCode] NVARCHAR(50) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [IPAddress] NVARCHAR(50) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "LoginAlerts", "CREATE TABLE [LoginAlerts] ([AlertID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [AlertType] NVARCHAR(50) NULL, [AlertLevel] NVARCHAR(20) NULL DEFAULT 'info', [AlertMessage] NVARCHAR(500) NULL, [IPAddress] NVARCHAR(50) NULL, [AdminID] INT NULL, [IsRead] BIT NULL DEFAULT 0, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "IPBlacklist", "CREATE TABLE [IPBlacklist] ([IPID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [IPAddress] NVARCHAR(50) NOT NULL, [Reason] NVARCHAR(255) NULL, [BlockedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [BlockedBy] INT NULL, [IsActive] BIT NULL DEFAULT 1, [ExpiresAt] DATETIME2(7) NULL, [HitCount] INT NULL DEFAULT 0, [LastHitAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ModulePermissions", "CREATE TABLE [ModulePermissions] ([PermissionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ModuleCode] NVARCHAR(50) NOT NULL, [ModuleName] NVARCHAR(100) NOT NULL, [ParentModule] NVARCHAR(50) NULL, [RequiredRole] NVARCHAR(20) NULL, [PermissionLevel] INT NULL, [URLPattern] NVARCHAR(200) NULL, [IsActive] BIT NULL DEFAULT 1)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 产品相关表 ---
    If CreateTableIfNotExists(conn, "Products", "CREATE TABLE [Products] ([ProductID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductName] NVARCHAR(100) NOT NULL, [ProductType] NVARCHAR(50) NULL, [Category] NVARCHAR(50) NULL, [BasePrice] DECIMAL(19,4) NOT NULL, [Description] NVARCHAR(MAX) NULL, [ImageURL] NVARCHAR(200) NULL, [IsActive] BIT NULL DEFAULT 1, [RecipeID] INT NULL, [BOMCost] DECIMAL(19,4) NULL, [UnitCost] DECIMAL(19,4) NULL, [BaseIngredients] NVARCHAR(MAX) NULL, [Engravable] BIT NULL DEFAULT 0, [EngravingPrice] DECIMAL(19,4) NULL, [KOLID] INT NULL, [ReviewStatus] NVARCHAR(20) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "Categories", "CREATE TABLE [Categories] ([CategoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [CategoryName] NVARCHAR(100) NOT NULL, [SortOrder] INT NULL, [IsActive] BIT NULL DEFAULT 1)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "FragranceNotes", "CREATE TABLE [FragranceNotes] ([NoteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [NoteName] NVARCHAR(50) NOT NULL, [NoteType] NVARCHAR(20) NOT NULL, [Description] NVARCHAR(MAX) NULL, [ImageURL] NVARCHAR(200) NULL, [Ingredients] NVARCHAR(MAX) NULL, [PriceAddition] DECIMAL(19,4) NULL, [RecommendedPercentage] INT NULL, [IsActive] BIT NULL DEFAULT 1, [IsBaseNote] INT NULL, [BaseNoteID] INT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "Ingredients", "CREATE TABLE [Ingredients] ([IngredientID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [IngredientName] NVARCHAR(100) NOT NULL, [CASNumber] NVARCHAR(50) NULL, [Description] NVARCHAR(255) NULL, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "FragranceIngredients", "CREATE TABLE [FragranceIngredients] ([FragranceIngredientID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [NoteID] INT NOT NULL, [IngredientID] INT NOT NULL, [Percentage] REAL NOT NULL, [SortOrder] INT NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductNotes", "CREATE TABLE [ProductNotes] ([ProductNoteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NOT NULL, [NoteID] INT NOT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductNoteRatios", "CREATE TABLE [ProductNoteRatios] ([RatioID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NOT NULL, [NoteID] INT NOT NULL, [NoteType] NVARCHAR(20) NULL, [Percentage] INT NOT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductVolumePrices", "CREATE TABLE [ProductVolumePrices] ([PVPriceID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NOT NULL, [VolumeID] INT NOT NULL, [Price] DECIMAL(19,4) NOT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductBottleStyles", "CREATE TABLE [ProductBottleStyles] ([ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NOT NULL, [BottleID] INT NOT NULL, [CustomPrice] DECIMAL(19,4) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductCosts", "CREATE TABLE [ProductCosts] ([CostID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NOT NULL, [CostName] NVARCHAR(100) NULL, [CostType] NVARCHAR(20) NULL, [UnitCost] DECIMAL(19,4) NULL, [TotalCost] DECIMAL(19,4) NULL, [Quantity] FLOAT NULL, [EffectiveDate] DATETIME2(7) NULL, [ExpiryDate] DATETIME2(7) NULL, [CreatedBy] NVARCHAR(50) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductInventory", "CREATE TABLE [ProductInventory] ([InventoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NULL, [NoteID] INT NULL, [StockType] NVARCHAR(20) NULL, [StockQty] INT NULL, [SafetyStock] INT NULL, [UnitCost] DECIMAL(19,4) NULL, [UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductTypeConfig", "CREATE TABLE [ProductTypeConfig] ([ConfigID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [TypeCode] NVARCHAR(20) NOT NULL, [DisplayName] NVARCHAR(50) NULL, [NavName] NVARCHAR(50) NULL, [Description] NVARCHAR(MAX) NULL, [Icon] NVARCHAR(100) NULL, [DisplayOrder] INT NULL, [RequiresRatio] BIT NULL DEFAULT 0, [RequiresReview] BIT NULL DEFAULT 0, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 配方相关表 ---
    If CreateTableIfNotExists(conn, "Recipes", "CREATE TABLE [Recipes] ([RecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RecipeCode] NVARCHAR(50) NULL, [RecipeName] NVARCHAR(100) NULL, [Description] NVARCHAR(MAX) NULL, [ProductType] NVARCHAR(20) NULL, [CreatedBy] NVARCHAR(100) NULL, [ReviewStatus] NVARCHAR(20) NULL, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipeNotes", "CREATE TABLE [RecipeNotes] ([ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RecipeID] INT NULL, [NoteID] INT NULL, [NoteType] NVARCHAR(20) NULL, [Percentage] INT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipeIngredients", "CREATE TABLE [RecipeIngredients] ([ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RecipeID] INT NULL, [IngredientName] NVARCHAR(100) NULL, [NoteID] INT NULL, [Percentage] FLOAT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipeProducts", "CREATE TABLE [RecipeProducts] ([ProductRecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RecipeID] INT NULL, [ProductID] INT NULL, [BatchSize] FLOAT NULL, [Status] NVARCHAR(20) NULL, [PublishedBy] NVARCHAR(50) NULL, [PublishedAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipeProductNotes", "CREATE TABLE [RecipeProductNotes] ([DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductRecipeID] INT NULL, [NoteID] INT NULL, [NoteName] NVARCHAR(100) NULL, [Percentage] FLOAT NULL, [PlannedQty] FLOAT NULL, [Notes] NVARCHAR(MAX) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipeAccords", "CREATE TABLE [RecipeAccords] ([AccordRecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RecipeID] INT NULL, [RecipeName] NVARCHAR(100) NULL, [NoteID] INT NULL, [BatchSize] FLOAT NULL, [Status] NVARCHAR(20) NULL, [PublishedBy] NVARCHAR(50) NULL, [PublishedAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipeAccordMaterials", "CREATE TABLE [RecipeAccordMaterials] ([DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [AccordRecipeID] INT NULL, [MaterialID] INT NULL, [MaterialName] NVARCHAR(100) NULL, [Percentage] FLOAT NULL, [PlannedQty] FLOAT NULL, [Notes] NVARCHAR(MAX) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipePublishLog", "CREATE TABLE [RecipePublishLog] ([LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RecipeID] INT NULL, [TargetRecipeID] INT NULL, [PublishType] NVARCHAR(20) NULL, [PublishedBy] NVARCHAR(50) NULL, [PublishedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [IPAddress] NVARCHAR(50) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecipePopularity", "CREATE TABLE [RecipePopularity] ([PopularityID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NOT NULL, [ViewCount] INT NULL DEFAULT 0, [FavoriteCount] INT NULL DEFAULT 0, [PurchaseCount] INT NULL DEFAULT 0, [LastCalculatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RecommendedRecipes", "CREATE TABLE [RecommendedRecipes] ([RecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RecipeName] NVARCHAR(200) NULL, [Description] NVARCHAR(MAX) NULL, [ProductID] INT NULL, [SortOrder] INT NULL, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "Formulas", "CREATE TABLE [Formulas] ([FormulaID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [FormulaName] NVARCHAR(100) NOT NULL, [Description] NVARCHAR(MAX) NULL, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "FormulaNotes", "CREATE TABLE [FormulaNotes] ([ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [FormulaID] INT NOT NULL, [NoteID] INT NOT NULL, [Percentage] INT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 订单相关表 ---
    If CreateTableIfNotExists(conn, "Orders", "CREATE TABLE [Orders] ([OrderID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OrderNo] NVARCHAR(50) NOT NULL, [UserID] INT NOT NULL, [TotalAmount] DECIMAL(19,4) NOT NULL, [Status] NVARCHAR(20) NULL DEFAULT 'pending', [ShippingStatus] NVARCHAR(20) NULL, [ShippingName] NVARCHAR(100) NULL, [ShippingPhone] NVARCHAR(20) NULL, [ShippingAddress] NVARCHAR(200) NULL, [ShippingCity] NVARCHAR(50) NULL, [ShippingPostalCode] NVARCHAR(20) NULL, [ShippingFee] DECIMAL(19,4) NULL DEFAULT 0, [ShippingCompany] NVARCHAR(50) NULL, [ShippingNotes] NVARCHAR(MAX) NULL, [TrackingNumber] NVARCHAR(100) NULL, [PaymentMethod] NVARCHAR(50) NULL, [ChannelSource] NVARCHAR(50) NULL, [CostAmount] DECIMAL(19,4) NULL, [ProfitAmount] DECIMAL(19,4) NULL, [RefundAmount] DECIMAL(19,4) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL, [ShippedAt] DATETIME2(7) NULL, [DeliveredAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "OrderDetails", "CREATE TABLE [OrderDetails] ([DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OrderID] INT NOT NULL, [ProductID] INT NOT NULL, [ProductName] NVARCHAR(200) NULL, [Quantity] INT NOT NULL, [UnitPrice] DECIMAL(19,4) NOT NULL, [Subtotal] DECIMAL(19,4) NOT NULL, [CustomLabel] NVARCHAR(200) NULL, [VolumeML] INT NULL, [VolumeName] NVARCHAR(50) NULL, [BottleName] NVARCHAR(100) NULL, [TopNoteName] NVARCHAR(100) NULL, [MiddleNoteName] NVARCHAR(100) NULL, [BaseNoteName] NVARCHAR(100) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "OrderDetailNoteSelections", "CREATE TABLE [OrderDetailNoteSelections] ([SelectionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [DetailID] INT NOT NULL, [NoteID] INT NOT NULL, [NoteType] NVARCHAR(20) NULL, [Percentage] INT NOT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "OrderIngredients", "CREATE TABLE [OrderIngredients] ([IngredientID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OrderID] INT NOT NULL, [DetailID] INT NULL, [IngredientName] NVARCHAR(100) NOT NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 购物车表 ---
    If CreateTableIfNotExists(conn, "Cart", "CREATE TABLE [Cart] ([CartID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [UserID] INT NULL, [SessionID] NVARCHAR(100) NULL, [ProductID] INT NOT NULL, [Quantity] INT NULL DEFAULT 1, [UnitPrice] DECIMAL(19,4) NOT NULL, [BottleID] INT NULL, [VolumeID] INT NULL, [TopNoteID] INT NULL, [MiddleNoteID] INT NULL, [BaseNoteID] INT NULL, [CustomLabel] NVARCHAR(200) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "CartNoteSelections", "CREATE TABLE [CartNoteSelections] ([SelectionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [CartID] INT NOT NULL, [NoteID] INT NOT NULL, [NoteType] NVARCHAR(20) NULL, [Percentage] INT NOT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 评价收藏表 ---
    If CreateTableIfNotExists(conn, "ProductReviews", "CREATE TABLE [ProductReviews] ([ReviewID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [UserID] INT NOT NULL, [ProductID] INT NULL, [OrderID] INT NOT NULL, [Rating] INT NULL, [Comment] NVARCHAR(MAX) NULL, [Status] NVARCHAR(20) NULL DEFAULT 'pending', [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "UserFavorites", "CREATE TABLE [UserFavorites] ([FavoriteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [UserID] INT NOT NULL, [ProductID] INT NOT NULL, [CreatedTime] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 用户扩展表 ---
    If CreateTableIfNotExists(conn, "UserAddresses", "CREATE TABLE [UserAddresses] ([AddressID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [UserID] INT NOT NULL, [Consignee] NVARCHAR(50) NOT NULL, [Phone] NVARCHAR(20) NOT NULL, [Address] NVARCHAR(200) NOT NULL, [Province] NVARCHAR(50) NULL, [City] NVARCHAR(50) NULL, [District] NVARCHAR(50) NULL, [IsDefault] BIT NULL DEFAULT 0, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "UserPoints", "CREATE TABLE [UserPoints] ([PointID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [UserID] INT NOT NULL, [TotalPoints] INT NULL DEFAULT 0, [AvailablePoints] INT NULL DEFAULT 0, [UsedPoints] INT NULL DEFAULT 0, [ExpiredPoints] INT NULL DEFAULT 0, [LastUpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "PointTransactions", "CREATE TABLE [PointTransactions] ([TransactionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [UserID] INT NOT NULL, [Points] INT NOT NULL, [PointsChange] INT NULL, [TransactionType] NVARCHAR(20) NULL, [Reason] NVARCHAR(200) NULL, [Description] NVARCHAR(255) NULL, [OrderID] INT NULL, [CreatedBy] NVARCHAR(50) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "UserPreferences", "CREATE TABLE [UserPreferences] ([PreferenceID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [UserID] INT NOT NULL, [PreferredTopNotes] NVARCHAR(255) NULL, [PreferredMiddleNotes] NVARCHAR(255) NULL, [PreferredBaseNotes] NVARCHAR(255) NULL, [PreferredCategories] NVARCHAR(255) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 库存相关表 ---
    If CreateTableIfNotExists(conn, "NoteInventory", "CREATE TABLE [NoteInventory] ([InventoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [NoteID] INT NOT NULL, [StockQuantity] INT NULL DEFAULT 0, [MinStockLevel] INT NULL, [LastRestockDate] DATETIME2(7) NULL, [UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "InventoryTransactions", "CREATE TABLE [InventoryTransactions] ([TransactionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [NoteID] INT NOT NULL, [ProductID] INT NULL, [MaterialID] INT NULL, [Quantity] INT NOT NULL, [TransactionType] NVARCHAR(20) NULL, [TransactionDirection] NVARCHAR(10) NULL, [ReferenceType] NVARCHAR(50) NULL, [ReferenceOrderID] INT NULL, [UnitCost] DECIMAL(19,4) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedBy] NVARCHAR(50) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 采购相关表 ---
    If CreateTableIfNotExists(conn, "Suppliers", "CREATE TABLE [Suppliers] ([SupplierID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [SupplierName] NVARCHAR(100) NOT NULL, [ContactPerson] NVARCHAR(50) NULL, [Phone] NVARCHAR(30) NULL, [Email] NVARCHAR(100) NULL, [Address] NVARCHAR(255) NULL, [Category] NVARCHAR(50) NULL, [Notes] NVARCHAR(MAX) NULL, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "PurchaseCategories", "CREATE TABLE [PurchaseCategories] ([CategoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [CategoryCode] NVARCHAR(20) NULL, [CategoryName] NVARCHAR(100) NULL, [Description] NVARCHAR(MAX) NULL, [DisplayOrder] INT NULL, [IsActive] BIT NULL DEFAULT 1)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "SupplierPrices", "CREATE TABLE [SupplierPrices] ([PriceID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [SupplierID] INT NULL, [ItemCode] NVARCHAR(50) NULL, [ItemName] NVARCHAR(200) NULL, [UnitPrice] DECIMAL(19,4) NULL, [MinOrderQty] FLOAT NULL, [EffectiveDate] DATETIME2(7) NULL, [ExpiryDate] DATETIME2(7) NULL, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "PurchaseOrders", "CREATE TABLE [PurchaseOrders] ([PurchaseID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [PurchaseNo] NVARCHAR(50) NULL, [SupplierID] INT NULL, [CategoryCode] NVARCHAR(20) NULL, [OrderDate] DATETIME2(7) NULL, [ExpectedDate] DATETIME2(7) NULL, [TotalAmount] DECIMAL(19,4) NULL, [Status] NVARCHAR(20) NULL DEFAULT 'draft', [Remarks] NVARCHAR(MAX) NULL, [CreatedBy] INT NULL, [ApprovedBy] INT NULL, [ApprovedAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "PurchaseOrderDetails", "CREATE TABLE [PurchaseOrderDetails] ([DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [PurchaseID] INT NULL, [ItemCode] NVARCHAR(50) NULL, [ItemName] NVARCHAR(200) NULL, [Specification] NVARCHAR(200) NULL, [Unit] NVARCHAR(20) NULL, [Quantity] FLOAT NULL, [UnitPrice] DECIMAL(19,4) NULL, [TotalPrice] DECIMAL(19,4) NULL, [ReceivedQty] FLOAT NULL DEFAULT 0, [Remarks] NVARCHAR(MAX) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "PurchaseReceipts", "CREATE TABLE [PurchaseReceipts] ([ReceiptID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ReceiptNo] NVARCHAR(50) NULL, [PurchaseID] INT NULL, [SupplierID] INT NULL, [ReceiptDate] DATETIME2(7) NULL, [TotalReceivedQty] FLOAT NULL, [ReceivedBy] NVARCHAR(50) NULL, [Status] NVARCHAR(20) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "PurchaseReceiptDetails", "CREATE TABLE [PurchaseReceiptDetails] ([ReceiptDetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ReceiptID] INT NULL, [PurchaseDetailID] INT NULL, [MaterialID] INT NULL, [ReceivedQty] FLOAT NULL, [AcceptedQty] FLOAT NULL, [RejectedQty] FLOAT NULL, [RejectReason] NVARCHAR(200) NULL, [UnitPrice] DECIMAL(19,4) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "PurchaseCostReview", "CREATE TABLE [PurchaseCostReview] ([ReviewID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [PurchaseID] INT NULL, [ReviewAmount] DECIMAL(19,4) NULL, [CostAllocation] NVARCHAR(20) NULL, [ReviewStatus] NVARCHAR(20) NULL, [ReviewComments] NVARCHAR(MAX) NULL, [ReviewerID] INT NULL, [ReviewedAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RawMaterialInventory", "CREATE TABLE [RawMaterialInventory] ([MaterialID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ItemCode] NVARCHAR(50) NULL, [ItemName] NVARCHAR(200) NULL, [CategoryCode] NVARCHAR(20) NULL, [Unit] NVARCHAR(20) NULL, [StockQty] FLOAT NULL DEFAULT 0, [SafetyStock] FLOAT NULL, [UnitPrice] DECIMAL(19,4) NULL, [SupplierID] INT NULL, [LastPurchaseDate] DATETIME2(7) NULL, [UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 财务相关表 ---
    If CreateTableIfNotExists(conn, "PaymentRecords", "CREATE TABLE [PaymentRecords] ([RecordID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OrderID] INT NULL, [OrderNo] NVARCHAR(50) NULL, [TransactionNo] NVARCHAR(100) NULL, [PaymentMethod] NVARCHAR(50) NULL, [Amount] DECIMAL(19,4) NULL, [Fee] DECIMAL(19,4) NULL, [NetAmount] DECIMAL(19,4) NULL, [Status] NVARCHAR(20) NULL, [TransactionType] NVARCHAR(20) NULL, [Category] NVARCHAR(50) NULL, [ReconcileStatus] NVARCHAR(20) NULL, [Remark] NVARCHAR(200) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ReconciliationLogs", "CREATE TABLE [ReconciliationLogs] ([LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OrderID] INT NULL, [OrderNo] NVARCHAR(50) NULL, [OrderAmount] DECIMAL(19,4) NULL, [PaymentAmount] DECIMAL(19,4) NULL, [Difference] DECIMAL(19,4) NULL, [Status] NVARCHAR(20) NULL, [Resolution] NVARCHAR(MAX) NULL, [ReconcileDate] DATETIME2(7) NULL, [ResolvedBy] NVARCHAR(50) NULL, [ResolvedAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "RefundRecords", "CREATE TABLE [RefundRecords] ([RefundID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [RefundNo] NVARCHAR(50) NULL, [OrderID] INT NOT NULL, [OrderNo] NVARCHAR(50) NULL, [RefundAmount] DECIMAL(19,4) NOT NULL, [RefundReason] NVARCHAR(MAX) NULL, [Status] NVARCHAR(20) NULL DEFAULT 'pending', [ApprovedBy] NVARCHAR(50) NULL, [ApprovedAt] DATETIME2(7) NULL, [CompletedAt] DATETIME2(7) NULL, [CostWriteBack] BIT NULL DEFAULT 0, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "FundAccounts", "CREATE TABLE [FundAccounts] ([AccountID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [AccountName] NVARCHAR(100) NULL, [AccountType] NVARCHAR(30) NULL, [TotalBalance] DECIMAL(19,4) NULL DEFAULT 0, [AvailableBalance] DECIMAL(19,4) NULL DEFAULT 0, [FrozenAmount] DECIMAL(19,4) NULL DEFAULT 0, [PendingSettlement] DECIMAL(19,4) NULL DEFAULT 0, [AlertThreshold] DECIMAL(19,4) NULL, [IsActive] BIT NULL DEFAULT 1, [LastSyncAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ExpenseRecords", "CREATE TABLE [ExpenseRecords] ([ExpenseID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ExpenseName] NVARCHAR(100) NULL, [ExpenseType] NVARCHAR(30) NULL, [Amount] DECIMAL(19,4) NULL, [AllocationMethod] NVARCHAR(20) NULL, [AllocationRatio] FLOAT NULL, [ProductID] INT NULL, [OrderID] INT NULL, [SourceOrderID] INT NULL, [Period] NVARCHAR(10) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "BudgetPlans", "CREATE TABLE [BudgetPlans] ([BudgetID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [BudgetName] NVARCHAR(100) NULL, [Category] NVARCHAR(50) NULL, [Period] NVARCHAR(10) NULL, [BudgetAmount] DECIMAL(19,4) NULL, [ActualAmount] DECIMAL(19,4) NULL DEFAULT 0, [GMVAmount] DECIMAL(19,4) NULL DEFAULT 0, [ROI] FLOAT NULL, [AlertPercent] FLOAT NULL, [AlertROI] FLOAT NULL, [Status] NVARCHAR(20) NULL, [CreatedBy] NVARCHAR(50) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- V10.2 新增: 财务扩展表 ---
    If CreateTableIfNotExists(conn, "AccountsReceivable", "CREATE TABLE [AccountsReceivable] ([ReceivableID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OrderID] INT NULL, [UserID] INT NOT NULL, [CustomerName] NVARCHAR(200) NULL, [ReceivableNo] NVARCHAR(50) NULL, [Amount] DECIMAL(19,4) NULL DEFAULT 0, [ReceivedAmount] DECIMAL(19,4) NULL DEFAULT 0, [Status] NVARCHAR(20) NULL DEFAULT 'Pending', [DueDate] DATE NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "AccountsPayable", "CREATE TABLE [AccountsPayable] ([PayableID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [PurchaseID] INT NULL, [SupplierID] INT NULL, [SupplierName] NVARCHAR(200) NULL, [PayableNo] NVARCHAR(50) NULL, [Amount] DECIMAL(19,4) NULL DEFAULT 0, [PaidAmount] DECIMAL(19,4) NULL DEFAULT 0, [Status] NVARCHAR(20) NULL DEFAULT 'Pending', [DueDate] DATE NULL, [InvoiceNo] NVARCHAR(100) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "CostCenters", "CREATE TABLE [CostCenters] ([CenterID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [CenterCode] NVARCHAR(50) NOT NULL, [CenterName] NVARCHAR(200) NOT NULL, [CenterType] NVARCHAR(50) NULL DEFAULT 'Department', [ParentID] INT NULL, [BudgetAmount] DECIMAL(19,4) NULL DEFAULT 0, [IsActive] BIT NULL DEFAULT 1, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "GLTransactions", "CREATE TABLE [GLTransactions] ([GLID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [GLNo] NVARCHAR(50) NULL, [TransactionDate] DATETIME2(7) NULL, [AccountCode] NVARCHAR(20) NULL, [AccountName] NVARCHAR(100) NULL, [DebitAmount] DECIMAL(19,4) NULL DEFAULT 0, [CreditAmount] DECIMAL(19,4) NULL DEFAULT 0, [CenterID] INT NULL, [RefType] NVARCHAR(50) NULL, [RefID] INT NULL, [RefNo] NVARCHAR(100) NULL, [Description] NVARCHAR(MAX) NULL, [CreatedBy] NVARCHAR(50) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 营销相关表 ---
    If CreateTableIfNotExists(conn, "Coupons", "CREATE TABLE [Coupons] ([CouponID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [CouponCode] NVARCHAR(50) NULL, [DiscountType] NVARCHAR(20) NULL, [DiscountValue] DECIMAL(19,4) NULL, [MinPurchase] DECIMAL(19,4) NULL, [StartDate] DATETIME2(7) NULL, [EndDate] DATETIME2(7) NULL, [UsageLimit] INT NULL, [UsedCount] INT NULL DEFAULT 0, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "MarketingCampaigns", "CREATE TABLE [MarketingCampaigns] ([CampaignID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [CampaignName] NVARCHAR(200) NULL, [CampaignType] NVARCHAR(50) NULL, [Description] NVARCHAR(MAX) NULL, [DiscountValue] DECIMAL(19,4) NULL, [MinPurchase] DECIMAL(19,4) NULL, [StartDate] DATETIME2(7) NULL, [EndDate] DATETIME2(7) NULL, [ParticipantCount] INT NULL DEFAULT 0, [TotalSales] DECIMAL(19,4) NULL DEFAULT 0, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "DailyStatistics", "CREATE TABLE [DailyStatistics] ([StatID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [StatDate] DATETIME2(7) NOT NULL, [NewUsers] INT NULL, [TotalUsers] INT NULL, [TotalOrders] INT NULL, [TotalRevenue] DECIMAL(19,4) NULL, [TopProductID] INT NULL, [TopNoteID] INT NULL, [DataJSON] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 生产相关表 ---
    If CreateTableIfNotExists(conn, "ProductionOrders", "CREATE TABLE [ProductionOrders] ([ProductionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [WorkOrderNo] NVARCHAR(50) NULL, [OrderID] INT NOT NULL, [DetailID] INT NULL, [RecipeID] INT NULL, [RecipeName] NVARCHAR(100) NULL, [TotalBottles] INT NULL, [BottleIndex] INT NULL, [Priority] INT NULL, [PriorityText] NVARCHAR(10) NULL, [Status] NVARCHAR(20) NULL DEFAULT 'pending', [AssignedTo] NVARCHAR(100) NULL, [EstimatedDate] DATETIME2(7) NULL, [Notes] NVARCHAR(MAX) NULL, [QCNotes] NVARCHAR(MAX) NULL, [StartedAt] DATETIME2(7) NULL, [CompletedAt] DATETIME2(7) NULL, [QCPassedAt] DATETIME2(7) NULL, [WarehouseInAt] DATETIME2(7) NULL, [ShippedOutAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductionLogs", "CREATE TABLE [ProductionLogs] ([LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductionID] INT NOT NULL, [Status] NVARCHAR(20) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedBy] NVARCHAR(100) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductManufacturing", "CREATE TABLE [ProductManufacturing] ([ManufacturingID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductID] INT NULL, [ProductName] NVARCHAR(100) NULL, [ProductRecipeID] INT NULL, [BatchNo] NVARCHAR(30) NULL, [PlannedQty] FLOAT NULL, [ActualQty] FLOAT NULL, [WorkCenter] NVARCHAR(20) NULL, [TransferRequestID] INT NULL, [Status] NVARCHAR(20) NULL, [Notes] NVARCHAR(MAX) NULL, [StartedAt] DATETIME2(7) NULL, [CompletedAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "ProductManufacturingDetails", "CREATE TABLE [ProductManufacturingDetails] ([DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ManufacturingID] INT NULL, [NoteID] INT NULL, [NoteName] NVARCHAR(100) NULL, [PlannedQty] FLOAT NULL, [ActualQty] FLOAT NULL, [UnitCost] DECIMAL(19,4) NULL, [TotalCost] DECIMAL(19,4) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "AccordProductions", "CREATE TABLE [AccordProductions] ([ProductionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [AccordRecipeID] INT NULL, [BatchNo] NVARCHAR(30) NULL, [NoteID] INT NULL, [NoteName] NVARCHAR(100) NULL, [PlannedQty] FLOAT NULL, [ActualQty] FLOAT NULL, [WorkCenter] NVARCHAR(20) NULL, [Status] NVARCHAR(20) NULL, [ApprovedBy] NVARCHAR(50) NULL, [Notes] NVARCHAR(MAX) NULL, [StartedAt] DATETIME2(7) NULL, [CompletedAt] DATETIME2(7) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), [UpdatedAt] DATETIME2(7) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "AccordProductionDetails", "CREATE TABLE [AccordProductionDetails] ([DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductionID] INT NULL, [MaterialID] INT NULL, [MaterialName] NVARCHAR(100) NULL, [PlannedQty] FLOAT NULL, [ActualQty] FLOAT NULL, [UnitCost] DECIMAL(19,4) NULL, [TotalCost] DECIMAL(19,4) NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "AccordQCReports", "CREATE TABLE [AccordQCReports] ([QCReportID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [ProductionID] INT NULL, [BatchNo] NVARCHAR(30) NULL, [QCResult] NVARCHAR(20) NULL, [TesterID] INT NULL, [TesterName] NVARCHAR(50) NULL, [TestDate] DATETIME2(7) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 物流与出库表 ---
    If CreateTableIfNotExists(conn, "MaterialOutbound", "CREATE TABLE [MaterialOutbound] ([OutboundID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OutboundNo] NVARCHAR(50) NULL, [OutboundType] NVARCHAR(20) NULL, [OutboundDate] DATETIME2(7) NULL, [ReferenceType] NVARCHAR(50) NULL, [ReferenceID] INT NULL, [RequestedBy] NVARCHAR(50) NULL, [ApprovedBy] NVARCHAR(50) NULL, [Status] NVARCHAR(20) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "MaterialOutboundDetails", "CREATE TABLE [MaterialOutboundDetails] ([OutboundDetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [OutboundID] INT NULL, [MaterialID] INT NULL, [RequestedQty] FLOAT NULL, [ActualQty] FLOAT NULL, [UnitPrice] DECIMAL(19,4) NULL, [TotalAmount] DECIMAL(19,4) NULL, [ProductionOrderRef] INT NULL)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "WorkshopTransfer", "CREATE TABLE [WorkshopTransfer] ([TransferID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [TransferNo] NVARCHAR(30) NULL, [NoteID] INT NULL, [FromWorkshop] NVARCHAR(20) NULL, [ToWorkshop] NVARCHAR(20) NULL, [RequestQty] FLOAT NULL, [RequestedBy] NVARCHAR(50) NULL, [RequestedAt] DATETIME2(7) NULL, [FulfilledAt] DATETIME2(7) NULL, [Status] NVARCHAR(20) NULL, [Notes] NVARCHAR(MAX) NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' --- 其他配置表 ---
    If CreateTableIfNotExists(conn, "BaseNotes", "CREATE TABLE [BaseNotes] ([BaseNoteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [BaseNoteName] NVARCHAR(100) NOT NULL, [Description] NVARCHAR(MAX) NULL, [Ingredients] NVARCHAR(MAX) NULL, [UnitPrice] DECIMAL(19,4) NULL, [IsActive] BIT NULL DEFAULT 1, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "NoteIngredients", "CREATE TABLE [NoteIngredients] ([ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [NoteID] INT NOT NULL, [BaseNoteID] INT NOT NULL, [Percentage] FLOAT NULL, [CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "BottleStyles", "CREATE TABLE [BottleStyles] ([BottleID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [BottleName] NVARCHAR(50) NOT NULL, [Description] NVARCHAR(MAX) NULL, [ImageURL] NVARCHAR(200) NULL, [PriceAddition] DECIMAL(19,4) NULL DEFAULT 0, [IsActive] BIT NULL DEFAULT 1)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "Volumes", "CREATE TABLE [Volumes] ([VolumeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, [VolumeName] NVARCHAR(50) NOT NULL, [VolumeML] INT NOT NULL, [PriceMultiplier] FLOAT NULL DEFAULT 1.0, [IsActive] BIT NULL DEFAULT 1)") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists(conn, "SiteSettings", "CREATE TABLE [SiteSettings] ([SettingKey] NVARCHAR(50) NULL, [SettingName] NVARCHAR(100) NULL, [SettingValue] NVARCHAR(255) NULL, [Description] NVARCHAR(255) NULL, [UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE())") Then successTables = successTables + 1
    totalTables = totalTables + 1

    Call Log("info", "表结构创建完成: " & successTables & "/" & totalTables & " 张表")
    Call StageEnd(IIf(successTables >= totalTables, "success", "warning"))

    ' ==== Stage 3b: 创建性能索引 ====
    Call StageStart("Stage 3b: 创建性能优化索引")

    Dim idxCount : idxCount = 0
    Dim idxSQL

    ' Orders 表索引
    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_UserID') CREATE NONCLUSTERED INDEX [IX_Orders_UserID] ON [Orders]([UserID]) INCLUDE ([OrderID],[TotalAmount],[Status],[CreatedAt])"
    If SafeExec(conn, idxSQL, "IX_Orders_UserID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_Status') CREATE NONCLUSTERED INDEX [IX_Orders_Status] ON [Orders]([Status]) INCLUDE ([OrderID],[TotalAmount],[CreatedAt])"
    If SafeExec(conn, idxSQL, "IX_Orders_Status") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_CreatedAt') CREATE NONCLUSTERED INDEX [IX_Orders_CreatedAt] ON [Orders]([CreatedAt]) INCLUDE ([OrderID],[TotalAmount],[Status])"
    If SafeExec(conn, idxSQL, "IX_Orders_CreatedAt") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_OrderNo') CREATE NONCLUSTERED INDEX [IX_Orders_OrderNo] ON [Orders]([OrderNo])"
    If SafeExec(conn, idxSQL, "IX_Orders_OrderNo") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_OrderDetails_OrderID') CREATE NONCLUSTERED INDEX [IX_OrderDetails_OrderID] ON [OrderDetails]([OrderID]) INCLUDE ([ProductID],[Quantity],[Subtotal])"
    If SafeExec(conn, idxSQL, "IX_OrderDetails_OrderID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_OrderDetails_ProductID') CREATE NONCLUSTERED INDEX [IX_OrderDetails_ProductID] ON [OrderDetails]([ProductID]) INCLUDE ([OrderID],[Quantity],[Subtotal])"
    If SafeExec(conn, idxSQL, "IX_OrderDetails_ProductID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Products_ProductType') CREATE NONCLUSTERED INDEX [IX_Products_ProductType] ON [Products]([ProductType]) INCLUDE ([ProductID],[ProductName],[BasePrice],[UnitCost],[IsActive])"
    If SafeExec(conn, idxSQL, "IX_Products_ProductType") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Products_IsActive') CREATE NONCLUSTERED INDEX [IX_Products_IsActive] ON [Products]([IsActive]) INCLUDE ([ProductID],[ProductName],[ProductType],[BasePrice])"
    If SafeExec(conn, idxSQL, "IX_Products_IsActive") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Users_Username') CREATE NONCLUSTERED INDEX [IX_Users_Username] ON [Users]([Username])"
    If SafeExec(conn, idxSQL, "IX_Users_Username") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Users_Email') CREATE NONCLUSTERED INDEX [IX_Users_Email] ON [Users]([Email])"
    If SafeExec(conn, idxSQL, "IX_Users_Email") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Users_CreatedAt') CREATE NONCLUSTERED INDEX [IX_Users_CreatedAt] ON [Users]([CreatedAt]) INCLUDE ([UserID])"
    If SafeExec(conn, idxSQL, "IX_Users_CreatedAt") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ProductReviews_ProductID') CREATE NONCLUSTERED INDEX [IX_ProductReviews_ProductID] ON [ProductReviews]([ProductID],[Status]) INCLUDE ([Rating],[UserID])"
    If SafeExec(conn, idxSQL, "IX_ProductReviews_ProductID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_UserFavorites_ProductID') CREATE NONCLUSTERED INDEX [IX_UserFavorites_ProductID] ON [UserFavorites]([ProductID]) INCLUDE ([UserID])"
    If SafeExec(conn, idxSQL, "IX_UserFavorites_ProductID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_UserFavorites_UserID') CREATE NONCLUSTERED INDEX [IX_UserFavorites_UserID] ON [UserFavorites]([UserID]) INCLUDE ([ProductID])"
    If SafeExec(conn, idxSQL, "IX_UserFavorites_UserID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_AdminLogs_CreatedAt') CREATE NONCLUSTERED INDEX [IX_AdminLogs_CreatedAt] ON [AdminLogs]([CreatedAt]) INCLUDE ([AdminID],[ActionType],[ModuleCode])"
    If SafeExec(conn, idxSQL, "IX_AdminLogs_CreatedAt") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_AdminLogs_ModuleCode') CREATE NONCLUSTERED INDEX [IX_AdminLogs_ModuleCode] ON [AdminLogs]([ModuleCode],[CreatedAt])"
    If SafeExec(conn, idxSQL, "IX_AdminLogs_ModuleCode") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_LoginAlerts_CreatedAt') CREATE NONCLUSTERED INDEX [IX_LoginAlerts_CreatedAt] ON [LoginAlerts]([CreatedAt]) INCLUDE ([AlertType],[AlertLevel],[IsRead])"
    If SafeExec(conn, idxSQL, "IX_LoginAlerts_CreatedAt") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_LoginAlerts_IsRead') CREATE NONCLUSTERED INDEX [IX_LoginAlerts_IsRead] ON [LoginAlerts]([IsRead],[CreatedAt])"
    If SafeExec(conn, idxSQL, "IX_LoginAlerts_IsRead") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_IPBlacklist_IPAddress') CREATE NONCLUSTERED INDEX [IX_IPBlacklist_IPAddress] ON [IPBlacklist]([IPAddress],[IsActive])"
    If SafeExec(conn, idxSQL, "IX_IPBlacklist_IPAddress") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_IPBlacklist_BlockedAt') CREATE NONCLUSTERED INDEX [IX_IPBlacklist_BlockedAt] ON [IPBlacklist]([BlockedAt]) INCLUDE ([IsActive])"
    If SafeExec(conn, idxSQL, "IX_IPBlacklist_BlockedAt") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_RawMaterialInventory_ItemCode') CREATE NONCLUSTERED INDEX [IX_RawMaterialInventory_ItemCode] ON [RawMaterialInventory]([ItemCode]) INCLUDE ([MaterialID],[ItemName],[StockQty],[UnitPrice],[SafetyStock])"
    If SafeExec(conn, idxSQL, "IX_RawMaterialInventory_ItemCode") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_NoteInventory_NoteID') CREATE NONCLUSTERED INDEX [IX_NoteInventory_NoteID] ON [NoteInventory]([NoteID]) INCLUDE ([StockQuantity])"
    If SafeExec(conn, idxSQL, "IX_NoteInventory_NoteID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_FragranceNotes_NoteType') CREATE NONCLUSTERED INDEX [IX_FragranceNotes_NoteType] ON [FragranceNotes]([NoteType]) INCLUDE ([NoteID],[NoteName],[PriceAddition])"
    If SafeExec(conn, idxSQL, "IX_FragranceNotes_NoteType") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ProductInventory_ProductID') CREATE NONCLUSTERED INDEX [IX_ProductInventory_ProductID] ON [ProductInventory]([ProductID]) INCLUDE ([StockQty],[UnitCost])"
    If SafeExec(conn, idxSQL, "IX_ProductInventory_ProductID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_SupplierPrices_ItemCode') CREATE NONCLUSTERED INDEX [IX_SupplierPrices_ItemCode] ON [SupplierPrices]([ItemCode],[IsActive]) INCLUDE ([UnitPrice],[CreatedAt])"
    If SafeExec(conn, idxSQL, "IX_SupplierPrices_ItemCode") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ProductionOrders_Status') CREATE NONCLUSTERED INDEX [IX_ProductionOrders_Status] ON [ProductionOrders]([Status]) INCLUDE ([ProductionID],[CreatedAt])"
    If SafeExec(conn, idxSQL, "IX_ProductionOrders_Status") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_PurchaseOrders_Status') CREATE NONCLUSTERED INDEX [IX_PurchaseOrders_Status] ON [PurchaseOrders]([Status]) INCLUDE ([PurchaseID],[TotalAmount])"
    If SafeExec(conn, idxSQL, "IX_PurchaseOrders_Status") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_RecipeAccordMaterials_AccordRecipeID') CREATE NONCLUSTERED INDEX [IX_RecipeAccordMaterials_AccordRecipeID] ON [RecipeAccordMaterials]([AccordRecipeID]) INCLUDE ([MaterialID],[PlannedQty])"
    If SafeExec(conn, idxSQL, "IX_RecipeAccordMaterials_AccordRecipeID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_RecipeAccords_NoteID') CREATE NONCLUSTERED INDEX [IX_RecipeAccords_NoteID] ON [RecipeAccords]([NoteID]) INCLUDE ([AccordRecipeID],[RecipeID])"
    If SafeExec(conn, idxSQL, "IX_RecipeAccords_NoteID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ProductNoteRatios_ProductID') CREATE NONCLUSTERED INDEX [IX_ProductNoteRatios_ProductID] ON [ProductNoteRatios]([ProductID]) INCLUDE ([NoteID],[Percentage])"
    If SafeExec(conn, idxSQL, "IX_ProductNoteRatios_ProductID") Then idxCount = idxCount + 1

    ' --- V10.2 新增: 财务扩展表索引 ---
    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_AccountsReceivable_Status') CREATE NONCLUSTERED INDEX [IX_AccountsReceivable_Status] ON [AccountsReceivable]([Status]) INCLUDE ([ReceivableID],[Amount],[ReceivedAmount],[DueDate])"
    If SafeExec(conn, idxSQL, "IX_AccountsReceivable_Status") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_AccountsReceivable_DueDate') CREATE NONCLUSTERED INDEX [IX_AccountsReceivable_DueDate] ON [AccountsReceivable]([DueDate]) INCLUDE ([Status],[Amount],[ReceivedAmount])"
    If SafeExec(conn, idxSQL, "IX_AccountsReceivable_DueDate") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_AccountsPayable_Status') CREATE NONCLUSTERED INDEX [IX_AccountsPayable_Status] ON [AccountsPayable]([Status]) INCLUDE ([PayableID],[Amount],[PaidAmount],[DueDate])"
    If SafeExec(conn, idxSQL, "IX_AccountsPayable_Status") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_AccountsPayable_DueDate') CREATE NONCLUSTERED INDEX [IX_AccountsPayable_DueDate] ON [AccountsPayable]([DueDate]) INCLUDE ([Status],[Amount],[PaidAmount])"
    If SafeExec(conn, idxSQL, "IX_AccountsPayable_DueDate") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_CostCenters_Code') CREATE NONCLUSTERED INDEX [IX_CostCenters_Code] ON [CostCenters]([CenterCode]) INCLUDE ([CenterName],[CenterType],[IsActive])"
    If SafeExec(conn, idxSQL, "IX_CostCenters_Code") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_GLTransactions_CenterID') CREATE NONCLUSTERED INDEX [IX_GLTransactions_CenterID] ON [GLTransactions]([CenterID],[TransactionDate]) INCLUDE ([DebitAmount],[CreditAmount],[RefType])"
    If SafeExec(conn, idxSQL, "IX_GLTransactions_CenterID") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_GLTransactions_GLNo') CREATE NONCLUSTERED INDEX [IX_GLTransactions_GLNo] ON [GLTransactions]([GLNo])"
    If SafeExec(conn, idxSQL, "IX_GLTransactions_GLNo") Then idxCount = idxCount + 1

    ' --- V10.3 新增: 财务报表性能索引 ---
    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_PaymentRecords_CreatedAt') CREATE NONCLUSTERED INDEX [IX_PaymentRecords_CreatedAt] ON [PaymentRecords]([CreatedAt]) INCLUDE ([Amount],[PaymentType],[PaymentMethod],[Status])"
    If SafeExec(conn, idxSQL, "IX_PaymentRecords_CreatedAt") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_PaymentRecords_TransactionType') CREATE NONCLUSTERED INDEX [IX_PaymentRecords_TransactionType] ON [PaymentRecords]([TransactionType]) INCLUDE ([Amount],[CreatedAt],[Status])"
    If SafeExec(conn, idxSQL, "IX_PaymentRecords_TransactionType") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ExpenseRecords_Period') CREATE NONCLUSTERED INDEX [IX_ExpenseRecords_Period] ON [ExpenseRecords]([Period]) INCLUDE ([ExpenseType],[Amount],[AllocationMethod])"
    If SafeExec(conn, idxSQL, "IX_ExpenseRecords_Period") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_BudgetPlans_Category_Period') CREATE NONCLUSTERED INDEX [IX_BudgetPlans_Category_Period] ON [BudgetPlans]([Category],[Period]) INCLUDE ([BudgetAmount],[ActualAmount],[GMVAmount],[ROI])"
    If SafeExec(conn, idxSQL, "IX_BudgetPlans_Category_Period") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_ProfitAmount_Status') CREATE NONCLUSTERED INDEX [IX_Orders_ProfitAmount_Status] ON [Orders]([Status]) INCLUDE ([OrderID],[TotalAmount],[ProfitAmount],[CostAmount],[CreatedAt])"
    If SafeExec(conn, idxSQL, "IX_Orders_ProfitAmount_Status") Then idxCount = idxCount + 1

    idxSQL = "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_GLTransactions_Date') CREATE NONCLUSTERED INDEX [IX_GLTransactions_Date] ON [GLTransactions]([TransactionDate]) INCLUDE ([DebitAmount],[CreditAmount],[AccountCode],[CenterID])"
    If SafeExec(conn, idxSQL, "IX_GLTransactions_Date") Then idxCount = idxCount + 1

    Call Log("success", "索引创建完成: " & idxCount & " 个")
    Call StageEnd("success")

    ' ==== Stage 4: 插入种子数据 ====
    Call StageStart("Stage 4: 插入种子数据（管理员、角色、配置等）")

    ' 4.1 AdminRoles
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM AdminRoles WHERE RoleCode='super_admin') INSERT INTO [AdminRoles]([RoleCode],[RoleName],[Description],[Permissions],[CreatedAt]) VALUES('super_admin','超级管理员','拥有系统所有权限','all',GETDATE())", "AdminRoles - super_admin") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM AdminRoles WHERE RoleCode='admin') INSERT INTO [AdminRoles]([RoleCode],[RoleName],[Description],[Permissions],[CreatedAt]) VALUES('admin','管理员','日常管理权限','operation,finance,inventory,purchase,logistics',GETDATE())", "AdminRoles - admin") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM AdminRoles WHERE RoleCode='finance') INSERT INTO [AdminRoles]([RoleCode],[RoleName],[Description],[Permissions],[CreatedAt]) VALUES('finance','财务','财务管理权限','finance',GETDATE())", "AdminRoles - finance") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM AdminRoles WHERE RoleCode='editor') INSERT INTO [AdminRoles]([RoleCode],[RoleName],[Description],[Permissions],[CreatedAt]) VALUES('editor','编辑','内容编辑权限','operation',GETDATE())", "AdminRoles - editor") Then : End If

    ' 4.2 AdminUser (默认密码: admin123)
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM AdminUsers WHERE Username='admin') INSERT INTO [AdminUsers]([Username],[PasswordHash],[Email],[FullName],[Department],[RoleID],[IsActive],[CreatedAt]) VALUES('admin','ef92b778bafe771e89245b89ecbc08a44a4e166c06659911881f383d4473e94f','admin@perfumeshop.com','系统管理员','管理部',(SELECT TOP 1 RoleID FROM AdminRoles WHERE RoleCode='super_admin'),1,GETDATE())", "AdminUser - admin") Then : End If

    ' 4.3 ProductTypeConfig
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM ProductTypeConfig) INSERT INTO [ProductTypeConfig]([TypeCode],[DisplayName],[NavName],[Description],[DisplayOrder],[RequiresRatio],[RequiresReview],[IsActive]) VALUES('custom','定制香氛','定制','个性化定制香水',1,1,1,1)", "ProductTypeConfig - custom") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM ProductTypeConfig WHERE TypeCode='standard') INSERT INTO [ProductTypeConfig]([TypeCode],[DisplayName],[NavName],[Description],[DisplayOrder],[RequiresRatio],[RequiresReview],[IsActive]) VALUES('standard','标准香氛','标准','标准系列香水',2,0,0,1)", "ProductTypeConfig - standard") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM ProductTypeConfig WHERE TypeCode='kol') INSERT INTO [ProductTypeConfig]([TypeCode],[DisplayName],[NavName],[Description],[DisplayOrder],[RequiresRatio],[RequiresReview],[IsActive]) VALUES('kol','KOL联名','KOL','KOL联名款',3,0,1,1)", "ProductTypeConfig - kol") Then : End If

    ' 4.4 Volumes
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Volumes) INSERT INTO [Volumes]([VolumeName],[VolumeML],[PriceMultiplier],[IsActive]) VALUES('30ml',30,1.0,1)", "Volumes - 30ml") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Volumes WHERE VolumeML=50) INSERT INTO [Volumes]([VolumeName],[VolumeML],[PriceMultiplier],[IsActive]) VALUES('50ml',50,1.5,1)", "Volumes - 50ml") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Volumes WHERE VolumeML=100) INSERT INTO [Volumes]([VolumeName],[VolumeML],[PriceMultiplier],[IsActive]) VALUES('100ml',100,2.5,1)", "Volumes - 100ml") Then : End If

    ' 4.5 Categories
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Categories) INSERT INTO [Categories]([CategoryName],[SortOrder],[IsActive]) VALUES('花香调',1,1)", "Categories - 花香调") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Categories WHERE CategoryName='果香调') INSERT INTO [Categories]([CategoryName],[SortOrder],[IsActive]) VALUES('果香调',2,1)", "Categories - 果香调") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Categories WHERE CategoryName='木质调') INSERT INTO [Categories]([CategoryName],[SortOrder],[IsActive]) VALUES('木质调',3,1)", "Categories - 木质调") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Categories WHERE CategoryName='东方调') INSERT INTO [Categories]([CategoryName],[SortOrder],[IsActive]) VALUES('东方调',4,1)", "Categories - 东方调") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM Categories WHERE CategoryName='清新调') INSERT INTO [Categories]([CategoryName],[SortOrder],[IsActive]) VALUES('清新调',5,1)", "Categories - 清新调") Then : End If

    ' 4.6 CostCenters (V10.2新增)
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM CostCenters) INSERT INTO [CostCenters]([CenterCode],[CenterName],[CenterType],[BudgetAmount]) VALUES('RAW_MAT','原料采购','Procurement',0),('PACKAGING','包装物采购','Procurement',0),('BOTTLE','瓶子采购','Procurement',0),('PRODUCTION','生产制造','Production',0),('LOGISTICS','物流运输','Logistics',0),('MARKETING','市场营销','Marketing',0),('ADMIN','行政管理','Admin',0),('RND','研发设计','R&D',0)", "CostCenters种子数据") Then : End If

    ' 4.7 SiteSettings
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM SiteSettings WHERE SettingKey='site_name') INSERT INTO [SiteSettings]([SettingKey],[SettingName],[SettingValue],[Description]) VALUES('site_name','站点名称','香氛定制','网站名称')", "SiteSettings - site_name") Then : End If
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM SiteSettings WHERE SettingKey='site_version') INSERT INTO [SiteSettings]([SettingKey],[SettingName],[SettingValue],[Description]) VALUES('site_version','系统版本','V10.4','当前系统版本号')", "SiteSettings - version") Then : End If

    ' 4.7 FundAccounts
    If SafeExec(conn, "IF NOT EXISTS (SELECT * FROM FundAccounts) INSERT INTO [FundAccounts]([AccountName],[AccountType],[TotalBalance],[AvailableBalance],[AlertThreshold],[IsActive],[CreatedAt]) VALUES('主账户','cash',100000.0000,100000.0000,10000.0000,1,GETDATE())", "FundAccounts - 主账户") Then : End If

    Call Log("success", "种子数据插入完成")
    Call StageEnd("success")

    ' ==== Stage 5: 验证并完善权限配置 ====
    Call StageStart("Stage 5: 验证并完善权限配置")
    
    ' 注：db_ddladmin/db_datareader/db_datawriter 已在 Stage 2b 授予
    ' 此阶段进行权限验证和兼容性配置
    
    ' 5.1 验证当前用户权限
    Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_ddladmin') AS HasDDL")
    If Not rs.EOF Then
        If rs.Fields("HasDDL").Value = 1 Then
            Call Log("success", "当前用户拥有 db_ddladmin 角色 ✓")
        Else
            Call Log("warning", "当前用户缺少 db_ddladmin 角色")
        End If
    End If
    rs.Close : Set rs = Nothing
    
    Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_datareader') AS HasReader")
    If Not rs.EOF Then
        If rs.Fields("HasReader").Value = 1 Then
            Call Log("success", "当前用户拥有 db_datareader 角色 ✓")
        Else
            Call Log("warning", "当前用户缺少 db_datareader 角色")
        End If
    End If
    rs.Close : Set rs = Nothing
    
    Set rs = conn.Execute("SELECT IS_ROLEMEMBER('db_datawriter') AS HasWriter")
    If Not rs.EOF Then
        If rs.Fields("HasWriter").Value = 1 Then
            Call Log("success", "当前用户拥有 db_datawriter 角色 ✓")
        Else
            Call Log("warning", "当前用户缺少 db_datawriter 角色")
        End If
    End If
    rs.Close : Set rs = Nothing
    
    ' 5.2 为 NETWORK SERVICE 授予权限（兼容旧版应用池身份）
    Call SafeExec(conn, "IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name='NT AUTHORITY\\NETWORK SERVICE') CREATE USER [NT AUTHORITY\\NETWORK SERVICE] FOR LOGIN [NT AUTHORITY\\NETWORK SERVICE]", "创建 NETWORK SERVICE 用户")
    Call SafeExec(conn, "ALTER ROLE db_datareader ADD MEMBER [NT AUTHORITY\\NETWORK SERVICE]", "授予 NETWORK SERVICE db_datareader")
    Call SafeExec(conn, "ALTER ROLE db_datawriter ADD MEMBER [NT AUTHORITY\\NETWORK SERVICE]", "授予 NETWORK SERVICE db_datawriter")
    Call SafeExec(conn, "GRANT BACKUP DATABASE TO [NT AUTHORITY\\NETWORK SERVICE]", "授予 NETWORK SERVICE BACKUP DATABASE")
    
    Call Log("success", "权限配置与验证完成")
    Call StageEnd("success")

    ' ==== Stage 6: 验证 ====
    Call StageStart("Stage 6: 部署验证")

    Dim verifyErr
    verifyErr = 0

    Set rs = conn.Execute("SELECT COUNT(*) FROM sys.tables")
    If Not rs.EOF Then Call Log("info", "数据库表总数: " & rs.Fields(0).Value)
    rs.Close : Set rs = Nothing

    ' 验证关键表
    Dim keyTables(10)
    keyTables(0) = "Users"
    keyTables(1) = "AdminUsers"
    keyTables(2) = "AdminRoles"
    keyTables(3) = "Products"
    keyTables(4) = "Orders"
    keyTables(5) = "OrderDetails"
    keyTables(6) = "FragranceNotes"
    keyTables(7) = "PaymentRecords"
    keyTables(8) = "PurchaseOrders"
    keyTables(9) = "ProductionOrders"
    keyTables(10) = "ProductInventory"

    Dim i, t
    For i = 0 To UBound(keyTables)
        t = keyTables(i)
        Set rs = conn.Execute("SELECT COUNT(*) FROM sys.tables WHERE name='" & t & "'")
        If rs.Fields(0).Value > 0 Then
            Call Log("success", ChrW(&H2713) & " " & t)
        Else
            verifyErr = verifyErr + 1
            Call Log("error", ChrW(&H2717) & " " & t & " - 未找到!")
        End If
        rs.Close : Set rs = Nothing
    Next

    ' 验证种子数据
    Set rs = conn.Execute("SELECT COUNT(*) FROM AdminRoles")
    If rs.Fields(0).Value > 0 Then Call Log("success", ChrW(&H2713) & " AdminRoles: " & rs.Fields(0).Value & " 条") Else Call Log("warning", "AdminRoles 为空")
    rs.Close : Set rs = Nothing

    Set rs = conn.Execute("SELECT COUNT(*) FROM AdminUsers")
    If rs.Fields(0).Value > 0 Then Call Log("success", ChrW(&H2713) & " AdminUsers: " & rs.Fields(0).Value & " 条") Else Call Log("warning", "AdminUsers 为空")
    rs.Close : Set rs = Nothing

    Set rs = conn.Execute("SELECT COUNT(*) FROM ProductTypeConfig")
    If rs.Fields(0).Value > 0 Then Call Log("success", ChrW(&H2713) & " ProductTypeConfig: " & rs.Fields(0).Value & " 条") Else Call Log("warning", "ProductTypeConfig 为空")
    rs.Close : Set rs = Nothing

    If verifyErr = 0 Then
        Call StageEnd("success")
    Else
        Call StageEnd("warning")
    End If

    conn.Close : Set conn = Nothing

    ' ==== 最终报告 ====
    Response.Write "<div class=""summary-box summary-success"">"
    Response.Write "<h2>部署完成!</h2>"
    Response.Write "<p><strong>数据库:</strong> PerfumeShop (SQL Server 2017, MSSQLSERVER)</p>"
    Response.Write "<p><strong>表总数:</strong> " & successTables & "/" & totalTables & " 张表创建成功</p>"
    Response.Write "<p><strong>索引:</strong> " & idxCount & " 个</p>"
    Response.Write "<p><strong>种子数据:</strong> 管理员账号、角色、分类、容量、产品类型、站点设置已插入</p>"
    If verifyErr > 0 Then Response.Write "<p style=""color:#e65100""><strong>警告:</strong> " & verifyErr & " 个关键表验证失败</p>"
    Response.Write "<hr>"
    Response.Write "<p><strong>管理员登录凭据:</strong></p>"
    Response.Write "<p>用户名: <code>admin</code> | 密码: <code>admin123</code></p>"
    Response.Write "<p style=""margin-top:10px""><strong>下一步:</strong></p>"
    Response.Write "<ol><li>运行 <a href=""data_migrate.asp"">数据迁移工具</a> 从 Access 导入现有数据</li>"
    Response.Write "<li>运行 <a href=""env_check.asp"">环境验证</a> 确认全部配置</li>"
    Response.Write "<li>运行 <a href=""test_suite.asp"">功能测试套件</a> 验证系统</li></ol>"
    Response.Write "</div>"
End Sub
%>
</div>

<div class="actions">
    <a href="deploy.asp"><button class="btn btn-primary">↻ 返回部署首页</button></a>
    <a href="env_check.asp"><button class="btn btn-success">环境诊断</button></a>
    <a href="../admin/login.asp"><button class="btn btn-primary" style="background:#1a237e">后台登录</button></a>
</div>

</body>
</html>
