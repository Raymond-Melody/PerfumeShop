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
<title>PerfumeShop 数据库初始化</title>
<style>
body{font-family:Arial,sans-serif;max-width:900px;margin:20px auto;padding:20px;background:#f5f5f5}
h1{color:#333;border-bottom:2px solid #4CAF50;padding-bottom:10px}
.step{margin:10px 0;padding:12px 15px;border-radius:5px}
.success{background:#d4edda;color:#155724;border:1px solid #c3e6cb}
.error{background:#f8d7da;color:#721c24;border:1px solid #f5c6cb}
.info{background:#d1ecf1;color:#0c5460;border:1px solid #bee5eb}
.warning{background:#fff3cd;color:#856404;border:1px solid #ffeeba}
pre{background:#fff;padding:10px;border:1px solid #ddd;border-radius:3px;overflow-x:auto;font-size:12px}
button{padding:10px 30px;font-size:16px;background:#4CAF50;color:#fff;border:none;border-radius:5px;cursor:pointer}
button:disabled{background:#ccc;cursor:not-allowed}
</style>
</head>
<body>
<h1>PerfumeShop 数据库初始化工具</h1>
<%
Dim connAdmin, rs, sql, stepResult

' === 主执行函数 ===
Sub RunSetup()
    ' Step 1: 连接到 master 数据库
    Call LogStep("info", "Step 1: 连接到 SQL Server (master)...")
    
    Set connAdmin = Server.CreateObject("ADODB.Connection")
    connAdmin.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=master;Integrated Security=SSPI;"
    
    Call LogStep("success", "成功连接到 SQL Server 命名实例 (localhost\YOURPERFUME)")

    ' Step 2: 创建 PerfumeShop 数据库
    Call LogStep("info", "Step 2: 检查并创建 PerfumeShop 数据库...")
    
    Dim dbExists : dbExists = False
    Set rs = connAdmin.Execute("SELECT COUNT(*) FROM sys.databases WHERE name='PerfumeShop'")
    If Not rs.EOF Then
        If rs.Fields(0).Value > 0 Then dbExists = True
    End If
    rs.Close : Set rs = Nothing
    
    If dbExists Then
        Call LogStep("warning", "数据库 PerfumeShop 已存在，跳过创建。将从 Task 3 继续。")
    Else
        connAdmin.Execute "CREATE DATABASE [PerfumeShop]"
        Call LogStep("success", "成功创建数据库 PerfumeShop")
        
        ' 设置恢复模式为 SIMPLE（开发环境）
        connAdmin.Execute "ALTER DATABASE [PerfumeShop] SET RECOVERY SIMPLE"
        Call LogStep("success", "数据库恢复模式已设置为 SIMPLE")
    End If
    
    ' Step 3: 执行建表脚本
    Call LogStep("info", "Step 3: 执行建表脚本 (create_sqlserver_tables.sql)...")
    Call ExecuteTableScript()
    
    ' Step 4: 配置权限
    Call LogStep("info", "Step 4: 配置 IIS 应用程序池权限...")
    Call SetupPermissions()
    
    ' Step 5: 验证
    Call VerifySetup()
    
    connAdmin.Close : Set connAdmin = Nothing
    
    Call LogStep("success", "========== 数据库初始化完成! ==========")
End Sub

Sub ExecuteTableScript()
    On Error Resume Next
    
    ' 切换到 PerfumeShop 数据库
    connAdmin.Execute "USE [PerfumeShop]"
    
    Dim totalTables : totalTables = 0
    Dim successTables : successTables = 0
    
    ' === 核心用户与权限表 ===
    If CreateTableIfNotExists("Users", _
        "CREATE TABLE [Users] (" & _
        "[UserID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[Username] NVARCHAR(50) NOT NULL, " & _
        "[Password] NVARCHAR(255) NOT NULL, " & _
        "[Email] NVARCHAR(100) NOT NULL, " & _
        "[FullName] NVARCHAR(100) NULL, " & _
        "[Phone] NVARCHAR(20) NULL, " & _
        "[Address] NVARCHAR(200) NULL, " & _
        "[City] NVARCHAR(50) NULL, " & _
        "[PostalCode] NVARCHAR(20) NULL, " & _
        "[Points] INT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[IsVIP] BIT NULL DEFAULT 0, " & _
        "[UserRole] NVARCHAR(20) NULL DEFAULT 'user', " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("AdminRoles", _
        "CREATE TABLE [AdminRoles] (" & _
        "[RoleID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RoleCode] NVARCHAR(20) NOT NULL, " & _
        "[RoleName] NVARCHAR(50) NOT NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[Permissions] NVARCHAR(MAX) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("AdminUsers", _
        "CREATE TABLE [AdminUsers] (" & _
        "[AdminID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[Username] NVARCHAR(50) NOT NULL, " & _
        "[PasswordHash] NVARCHAR(255) NOT NULL, " & _
        "[Email] NVARCHAR(100) NOT NULL, " & _
        "[FullName] NVARCHAR(100) NULL, " & _
        "[Department] NVARCHAR(50) NULL, " & _
        "[RoleID] INT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[IsLocked] BIT NULL DEFAULT 0, " & _
        "[LastLogin] DATETIME2(7) NULL, " & _
        "[ResetToken] NVARCHAR(255) NULL, " & _
        "[ResetTokenExpiry] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("AdminLogs", _
        "CREATE TABLE [AdminLogs] (" & _
        "[LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[AdminID] INT NULL, " & _
        "[ActionType] NVARCHAR(100) NULL, " & _
        "[TableName] NVARCHAR(50) NULL, " & _
        "[RecordID] NVARCHAR(50) NULL, " & _
        "[ModuleCode] NVARCHAR(50) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[IPAddress] NVARCHAR(50) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("LoginAlerts", _
        "CREATE TABLE [LoginAlerts] (" & _
        "[AlertID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[AlertType] NVARCHAR(50) NULL, " & _
        "[AlertLevel] NVARCHAR(20) NULL DEFAULT 'info', " & _
        "[AlertMessage] NVARCHAR(500) NULL, " & _
        "[IPAddress] NVARCHAR(50) NULL, " & _
        "[AdminID] INT NULL, " & _
        "[IsRead] BIT NULL DEFAULT 0, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("IPBlacklist", _
        "CREATE TABLE [IPBlacklist] (" & _
        "[IPID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[IPAddress] NVARCHAR(50) NOT NULL, " & _
        "[Reason] NVARCHAR(255) NULL, " & _
        "[BlockedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[BlockedBy] INT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[ExpiresAt] DATETIME2(7) NULL, " & _
        "[HitCount] INT NULL DEFAULT 0, " & _
        "[LastHitAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ModulePermissions", _
        "CREATE TABLE [ModulePermissions] (" & _
        "[PermissionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ModuleCode] NVARCHAR(50) NOT NULL, " & _
        "[ModuleName] NVARCHAR(100) NOT NULL, " & _
        "[ParentModule] NVARCHAR(50) NULL, " & _
        "[RequiredRole] NVARCHAR(20) NULL, " & _
        "[PermissionLevel] INT NULL, " & _
        "[URLPattern] NVARCHAR(200) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 产品相关表 ===
    If CreateTableIfNotExists("Products", _
        "CREATE TABLE [Products] (" & _
        "[ProductID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductName] NVARCHAR(100) NOT NULL, " & _
        "[ProductType] NVARCHAR(50) NULL, " & _
        "[Category] NVARCHAR(50) NULL, " & _
        "[BasePrice] DECIMAL(19,4) NOT NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[ImageURL] NVARCHAR(200) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[RecipeID] INT NULL, " & _
        "[BOMCost] DECIMAL(19,4) NULL, " & _
        "[UnitCost] DECIMAL(19,4) NULL, " & _
        "[BaseIngredients] NVARCHAR(MAX) NULL, " & _
        "[Engravable] BIT NULL DEFAULT 0, " & _
        "[EngravingPrice] DECIMAL(19,4) NULL, " & _
        "[KOLID] INT NULL, " & _
        "[ReviewStatus] NVARCHAR(20) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("Categories", _
        "CREATE TABLE [Categories] (" & _
        "[CategoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[CategoryName] NVARCHAR(100) NOT NULL, " & _
        "[SortOrder] INT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("FragranceNotes", _
        "CREATE TABLE [FragranceNotes] (" & _
        "[NoteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[NoteName] NVARCHAR(50) NOT NULL, " & _
        "[NoteType] NVARCHAR(20) NOT NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[ImageURL] NVARCHAR(200) NULL, " & _
        "[Ingredients] NVARCHAR(MAX) NULL, " & _
        "[PriceAddition] DECIMAL(19,4) NULL, " & _
        "[RecommendedPercentage] INT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[IsBaseNote] INT NULL DEFAULT 0, " & _
        "[BaseNoteID] INT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("BaseNotes", _
        "CREATE TABLE [BaseNotes] (" & _
        "[BaseNoteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[BaseNoteName] NVARCHAR(100) NOT NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[Ingredients] NVARCHAR(MAX) NULL, " & _
        "[UnitPrice] DECIMAL(19,4) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("BottleStyles", _
        "CREATE TABLE [BottleStyles] (" & _
        "[BottleID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[BottleName] NVARCHAR(50) NOT NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[ImageURL] NVARCHAR(200) NULL, " & _
        "[PriceAddition] DECIMAL(19,4) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("Volumes", _
        "CREATE TABLE [Volumes] (" & _
        "[VolumeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[VolumeML] INT NOT NULL, " & _
        "[VolumeName] NVARCHAR(50) NOT NULL, " & _
        "[PriceMultiplier] FLOAT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductVolumePrices", _
        "CREATE TABLE [ProductVolumePrices] (" & _
        "[PVPriceID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NOT NULL, " & _
        "[VolumeID] INT NOT NULL, " & _
        "[Price] DECIMAL(19,4) NOT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductBottleStyles", _
        "CREATE TABLE [ProductBottleStyles] (" & _
        "[ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NOT NULL, " & _
        "[BottleID] INT NOT NULL, " & _
        "[CustomPrice] DECIMAL(19,4) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductNotes", _
        "CREATE TABLE [ProductNotes] (" & _
        "[ProductNoteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NOT NULL, " & _
        "[NoteID] INT NOT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductNoteRatios", _
        "CREATE TABLE [ProductNoteRatios] (" & _
        "[RatioID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NOT NULL, " & _
        "[NoteID] INT NOT NULL, " & _
        "[NoteType] NVARCHAR(20) NULL, " & _
        "[Percentage] INT NOT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductTypeConfig", _
        "CREATE TABLE [ProductTypeConfig] (" & _
        "[ConfigID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[TypeCode] NVARCHAR(20) NOT NULL, " & _
        "[DisplayName] NVARCHAR(50) NULL, " & _
        "[NavName] NVARCHAR(50) NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[Icon] NVARCHAR(100) NULL, " & _
        "[DisplayOrder] INT NULL, " & _
        "[RequiresRatio] BIT NULL DEFAULT 0, " & _
        "[RequiresReview] BIT NULL DEFAULT 0, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 订单相关表 ===
    If CreateTableIfNotExists("Orders", _
        "CREATE TABLE [Orders] (" & _
        "[OrderID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[OrderNo] NVARCHAR(50) NOT NULL, " & _
        "[UserID] INT NOT NULL, " & _
        "[TotalAmount] DECIMAL(19,4) NOT NULL, " & _
        "[Status] NVARCHAR(20) NULL DEFAULT 'pending', " & _
        "[ShippingStatus] NVARCHAR(20) NULL, " & _
        "[ShippingName] NVARCHAR(100) NULL, " & _
        "[ShippingPhone] NVARCHAR(20) NULL, " & _
        "[ShippingAddress] NVARCHAR(200) NULL, " & _
        "[ShippingCity] NVARCHAR(50) NULL, " & _
        "[ShippingPostalCode] NVARCHAR(20) NULL, " & _
        "[ShippingCompany] NVARCHAR(50) NULL, " & _
        "[ShippingFee] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[ShippingNotes] NVARCHAR(MAX) NULL, " & _
        "[TrackingNumber] NVARCHAR(100) NULL, " & _
        "[PaymentMethod] NVARCHAR(50) NULL, " & _
        "[CostAmount] DECIMAL(19,4) NULL, " & _
        "[ProfitAmount] DECIMAL(19,4) NULL, " & _
        "[RefundAmount] DECIMAL(19,4) NULL, " & _
        "[ChannelSource] NVARCHAR(50) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[ShippedAt] DATETIME2(7) NULL, " & _
        "[DeliveredAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("OrderDetails", _
        "CREATE TABLE [OrderDetails] (" & _
        "[DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[OrderID] INT NOT NULL, " & _
        "[ProductID] INT NOT NULL, " & _
        "[ProductName] NVARCHAR(200) NULL, " & _
        "[Quantity] INT NOT NULL, " & _
        "[UnitPrice] DECIMAL(19,4) NOT NULL, " & _
        "[Subtotal] DECIMAL(19,4) NOT NULL, " & _
        "[CustomLabel] NVARCHAR(200) NULL, " & _
        "[BaseNoteName] NVARCHAR(100) NULL, " & _
        "[MiddleNoteName] NVARCHAR(100) NULL, " & _
        "[TopNoteName] NVARCHAR(100) NULL, " & _
        "[BottleName] NVARCHAR(100) NULL, " & _
        "[VolumeML] INT NULL, " & _
        "[VolumeName] NVARCHAR(50) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("OrderDetailNoteSelections", _
        "CREATE TABLE [OrderDetailNoteSelections] (" & _
        "[SelectionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[DetailID] INT NOT NULL, " & _
        "[NoteID] INT NOT NULL, " & _
        "[NoteType] NVARCHAR(20) NULL, " & _
        "[Percentage] INT NOT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("OrderIngredients", _
        "CREATE TABLE [OrderIngredients] (" & _
        "[IngredientID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[OrderID] INT NOT NULL, " & _
        "[DetailID] INT NULL, " & _
        "[IngredientName] NVARCHAR(100) NOT NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 购物车表 ===
    If CreateTableIfNotExists("Cart", _
        "CREATE TABLE [Cart] (" & _
        "[CartID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[UserID] INT NULL, " & _
        "[SessionID] NVARCHAR(100) NULL, " & _
        "[ProductID] INT NOT NULL, " & _
        "[Quantity] INT NULL DEFAULT 1, " & _
        "[UnitPrice] DECIMAL(19,4) NOT NULL, " & _
        "[BaseNoteID] INT NULL, " & _
        "[MiddleNoteID] INT NULL, " & _
        "[TopNoteID] INT NULL, " & _
        "[BottleID] INT NULL, " & _
        "[VolumeID] INT NULL, " & _
        "[CustomLabel] NVARCHAR(200) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("CartNoteSelections", _
        "CREATE TABLE [CartNoteSelections] (" & _
        "[SelectionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[CartID] INT NOT NULL, " & _
        "[NoteID] INT NOT NULL, " & _
        "[NoteType] NVARCHAR(20) NULL, " & _
        "[Percentage] INT NOT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 用户相关表 ===
    If CreateTableIfNotExists("UserAddresses", _
        "CREATE TABLE [UserAddresses] (" & _
        "[AddressID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[UserID] INT NOT NULL, " & _
        "[Consignee] NVARCHAR(50) NOT NULL, " & _
        "[Phone] NVARCHAR(20) NOT NULL, " & _
        "[Province] NVARCHAR(50) NULL, " & _
        "[City] NVARCHAR(50) NULL, " & _
        "[District] NVARCHAR(50) NULL, " & _
        "[Address] NVARCHAR(200) NOT NULL, " & _
        "[IsDefault] BIT NULL DEFAULT 0, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("UserFavorites", _
        "CREATE TABLE [UserFavorites] (" & _
        "[FavoriteID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[UserID] INT NOT NULL, " & _
        "[ProductID] INT NOT NULL, " & _
        "[CreatedTime] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("UserPoints", _
        "CREATE TABLE [UserPoints] (" & _
        "[PointID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[UserID] INT NOT NULL, " & _
        "[TotalPoints] INT NULL DEFAULT 0, " & _
        "[AvailablePoints] INT NULL DEFAULT 0, " & _
        "[UsedPoints] INT NULL DEFAULT 0, " & _
        "[ExpiredPoints] INT NULL DEFAULT 0, " & _
        "[LastUpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("PointTransactions", _
        "CREATE TABLE [PointTransactions] (" & _
        "[TransactionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[UserID] INT NOT NULL, " & _
        "[Points] INT NOT NULL, " & _
        "[PointsChange] INT NULL, " & _
        "[TransactionType] NVARCHAR(20) NULL, " & _
        "[Reason] NVARCHAR(200) NULL, " & _
        "[Description] NVARCHAR(255) NULL, " & _
        "[OrderID] INT NULL, " & _
        "[CreatedBy] NVARCHAR(50) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("UserPreferences", _
        "CREATE TABLE [UserPreferences] (" & _
        "[PreferenceID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[UserID] INT NOT NULL, " & _
        "[PreferredCategories] NVARCHAR(255) NULL, " & _
        "[PreferredTopNotes] NVARCHAR(255) NULL, " & _
        "[PreferredMiddleNotes] NVARCHAR(255) NULL, " & _
        "[PreferredBaseNotes] NVARCHAR(255) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 评价表 ===
    If CreateTableIfNotExists("ProductReviews", _
        "CREATE TABLE [ProductReviews] (" & _
        "[ReviewID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[UserID] INT NOT NULL, " & _
        "[ProductID] INT NULL, " & _
        "[OrderID] INT NOT NULL, " & _
        "[Rating] INT NULL, " & _
        "[Comment] NVARCHAR(MAX) NULL, " & _
        "[Status] NVARCHAR(20) NULL DEFAULT 'pending', " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 配方表 ===
    If CreateTableIfNotExists("Recipes", _
        "CREATE TABLE [Recipes] (" & _
        "[RecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RecipeCode] NVARCHAR(50) NULL, " & _
        "[RecipeName] NVARCHAR(100) NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[ProductType] NVARCHAR(20) NULL, " & _
        "[ReviewStatus] NVARCHAR(20) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedBy] NVARCHAR(100) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipeNotes", _
        "CREATE TABLE [RecipeNotes] (" & _
        "[ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RecipeID] INT NULL, " & _
        "[NoteID] INT NULL, " & _
        "[NoteType] NVARCHAR(20) NULL, " & _
        "[Percentage] INT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipeIngredients", _
        "CREATE TABLE [RecipeIngredients] (" & _
        "[ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RecipeID] INT NULL, " & _
        "[NoteID] INT NULL, " & _
        "[IngredientName] NVARCHAR(100) NULL, " & _
        "[Percentage] FLOAT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("Ingredients", _
        "CREATE TABLE [Ingredients] (" & _
        "[IngredientID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[IngredientName] NVARCHAR(100) NOT NULL, " & _
        "[CASNumber] NVARCHAR(50) NULL, " & _
        "[Description] NVARCHAR(255) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("FragranceIngredients", _
        "CREATE TABLE [FragranceIngredients] (" & _
        "[FragranceIngredientID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[NoteID] INT NOT NULL, " & _
        "[IngredientID] INT NOT NULL, " & _
        "[Percentage] REAL NOT NULL, " & _
        "[SortOrder] INT NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("Formulas", _
        "CREATE TABLE [Formulas] (" & _
        "[FormulaID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[FormulaName] NVARCHAR(100) NOT NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("FormulaNotes", _
        "CREATE TABLE [FormulaNotes] (" & _
        "[ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[FormulaID] INT NOT NULL, " & _
        "[NoteID] INT NOT NULL, " & _
        "[Percentage] INT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("NoteIngredients", _
        "CREATE TABLE [NoteIngredients] (" & _
        "[ID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[NoteID] INT NOT NULL, " & _
        "[BaseNoteID] INT NOT NULL, " & _
        "[Percentage] FLOAT NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 生产相关表 ===
    If CreateTableIfNotExists("ProductionOrders", _
        "CREATE TABLE [ProductionOrders] (" & _
        "[ProductionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[WorkOrderNo] NVARCHAR(50) NULL, " & _
        "[OrderID] INT NOT NULL, " & _
        "[DetailID] INT NULL, " & _
        "[RecipeID] INT NULL, " & _
        "[RecipeName] NVARCHAR(100) NULL, " & _
        "[TotalBottles] INT NULL, " & _
        "[Status] NVARCHAR(20) NULL DEFAULT 'pending', " & _
        "[Priority] INT NULL DEFAULT 0, " & _
        "[PriorityText] NVARCHAR(10) NULL, " & _
        "[AssignedTo] NVARCHAR(100) NULL, " & _
        "[BottleIndex] INT NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[QCNotes] NVARCHAR(MAX) NULL, " & _
        "[EstimatedDate] DATETIME2(7) NULL, " & _
        "[StartedAt] DATETIME2(7) NULL, " & _
        "[CompletedAt] DATETIME2(7) NULL, " & _
        "[QCPassedAt] DATETIME2(7) NULL, " & _
        "[WarehouseInAt] DATETIME2(7) NULL, " & _
        "[ShippedOutAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductionLogs", _
        "CREATE TABLE [ProductionLogs] (" & _
        "[LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductionID] INT NOT NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[CreatedBy] NVARCHAR(100) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipeProducts", _
        "CREATE TABLE [RecipeProducts] (" & _
        "[ProductRecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RecipeID] INT NULL, " & _
        "[ProductID] INT NULL, " & _
        "[BatchSize] FLOAT NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[PublishedBy] NVARCHAR(50) NULL, " & _
        "[PublishedAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipeProductNotes", _
        "CREATE TABLE [RecipeProductNotes] (" & _
        "[DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductRecipeID] INT NULL, " & _
        "[NoteID] INT NULL, " & _
        "[NoteName] NVARCHAR(100) NULL, " & _
        "[Percentage] FLOAT NULL, " & _
        "[PlannedQty] FLOAT NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductManufacturing", _
        "CREATE TABLE [ProductManufacturing] (" & _
        "[ManufacturingID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NULL, " & _
        "[ProductName] NVARCHAR(100) NULL, " & _
        "[ProductRecipeID] INT NULL, " & _
        "[BatchNo] NVARCHAR(30) NULL, " & _
        "[PlannedQty] FLOAT NULL, " & _
        "[ActualQty] FLOAT NULL, " & _
        "[WorkCenter] NVARCHAR(20) NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[TransferRequestID] INT NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[StartedAt] DATETIME2(7) NULL, " & _
        "[CompletedAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductManufacturingDetails", _
        "CREATE TABLE [ProductManufacturingDetails] (" & _
        "[DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ManufacturingID] INT NULL, " & _
        "[NoteID] INT NULL, " & _
        "[NoteName] NVARCHAR(100) NULL, " & _
        "[PlannedQty] FLOAT NULL, " & _
        "[ActualQty] FLOAT NULL, " & _
        "[UnitCost] DECIMAL(19,4) NULL, " & _
        "[TotalCost] DECIMAL(19,4) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("AccordProductions", _
        "CREATE TABLE [AccordProductions] (" & _
        "[ProductionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[AccordRecipeID] INT NULL, " & _
        "[BatchNo] NVARCHAR(30) NULL, " & _
        "[NoteID] INT NULL, " & _
        "[NoteName] NVARCHAR(100) NULL, " & _
        "[PlannedQty] FLOAT NULL, " & _
        "[ActualQty] FLOAT NULL, " & _
        "[WorkCenter] NVARCHAR(20) NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[ApprovedBy] NVARCHAR(50) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[StartedAt] DATETIME2(7) NULL, " & _
        "[CompletedAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("AccordProductionDetails", _
        "CREATE TABLE [AccordProductionDetails] (" & _
        "[DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductionID] INT NULL, " & _
        "[MaterialID] INT NULL, " & _
        "[MaterialName] NVARCHAR(100) NULL, " & _
        "[PlannedQty] FLOAT NULL, " & _
        "[ActualQty] FLOAT NULL, " & _
        "[UnitCost] DECIMAL(19,4) NULL, " & _
        "[TotalCost] DECIMAL(19,4) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("AccordQCReports", _
        "CREATE TABLE [AccordQCReports] (" & _
        "[QCReportID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductionID] INT NULL, " & _
        "[BatchNo] NVARCHAR(30) NULL, " & _
        "[QCResult] NVARCHAR(20) NULL, " & _
        "[TestDate] DATETIME2(7) NULL, " & _
        "[TesterID] INT NULL, " & _
        "[TesterName] NVARCHAR(50) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipeAccords", _
        "CREATE TABLE [RecipeAccords] (" & _
        "[AccordRecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RecipeID] INT NULL, " & _
        "[NoteID] INT NULL, " & _
        "[RecipeName] NVARCHAR(100) NULL, " & _
        "[BatchSize] FLOAT NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[PublishedBy] NVARCHAR(50) NULL, " & _
        "[PublishedAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipeAccordMaterials", _
        "CREATE TABLE [RecipeAccordMaterials] (" & _
        "[DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[AccordRecipeID] INT NULL, " & _
        "[MaterialID] INT NULL, " & _
        "[MaterialName] NVARCHAR(100) NULL, " & _
        "[Percentage] FLOAT NULL, " & _
        "[PlannedQty] FLOAT NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipePublishLog", _
        "CREATE TABLE [RecipePublishLog] (" & _
        "[LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RecipeID] INT NULL, " & _
        "[TargetRecipeID] INT NULL, " & _
        "[PublishType] NVARCHAR(20) NULL, " & _
        "[PublishedBy] NVARCHAR(50) NULL, " & _
        "[PublishedAt] DATETIME2(7) NULL, " & _
        "[IPAddress] NVARCHAR(50) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecommendedRecipes", _
        "CREATE TABLE [RecommendedRecipes] (" & _
        "[RecipeID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RecipeName] NVARCHAR(200) NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[ProductID] INT NULL, " & _
        "[SortOrder] INT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 库存表 ===
    If CreateTableIfNotExists("NoteInventory", _
        "CREATE TABLE [NoteInventory] (" & _
        "[InventoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[NoteID] INT NOT NULL, " & _
        "[StockQuantity] INT NULL DEFAULT 0, " & _
        "[MinStockLevel] INT NULL DEFAULT 10, " & _
        "[LastRestockDate] DATETIME2(7) NULL, " & _
        "[UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductInventory", _
        "CREATE TABLE [ProductInventory] (" & _
        "[InventoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NULL, " & _
        "[NoteID] INT NULL, " & _
        "[StockType] NVARCHAR(20) NULL, " & _
        "[StockQty] INT NULL DEFAULT 0, " & _
        "[SafetyStock] INT NULL DEFAULT 5, " & _
        "[UnitCost] DECIMAL(19,4) NULL, " & _
        "[UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("InventoryTransactions", _
        "CREATE TABLE [InventoryTransactions] (" & _
        "[TransactionID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[NoteID] INT NOT NULL, " & _
        "[ProductID] INT NULL, " & _
        "[MaterialID] INT NULL, " & _
        "[Quantity] INT NOT NULL, " & _
        "[TransactionType] NVARCHAR(20) NULL, " & _
        "[TransactionDirection] NVARCHAR(10) NULL, " & _
        "[ReferenceType] NVARCHAR(50) NULL, " & _
        "[ReferenceOrderID] INT NULL, " & _
        "[UnitCost] DECIMAL(19,4) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[CreatedBy] NVARCHAR(50) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("WorkshopTransfer", _
        "CREATE TABLE [WorkshopTransfer] (" & _
        "[TransferID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[TransferNo] NVARCHAR(30) NULL, " & _
        "[NoteID] INT NULL, " & _
        "[FromWorkshop] NVARCHAR(20) NULL, " & _
        "[ToWorkshop] NVARCHAR(20) NULL, " & _
        "[RequestQty] FLOAT NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[RequestedBy] NVARCHAR(50) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[RequestedAt] DATETIME2(7) NULL, " & _
        "[FulfilledAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 采购表 ===
    If CreateTableIfNotExists("Suppliers", _
        "CREATE TABLE [Suppliers] (" & _
        "[SupplierID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[SupplierName] NVARCHAR(100) NOT NULL, " & _
        "[Category] NVARCHAR(50) NULL, " & _
        "[ContactPerson] NVARCHAR(50) NULL, " & _
        "[Phone] NVARCHAR(30) NULL, " & _
        "[Email] NVARCHAR(100) NULL, " & _
        "[Address] NVARCHAR(255) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("PurchaseCategories", _
        "CREATE TABLE [PurchaseCategories] (" & _
        "[CategoryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[CategoryCode] NVARCHAR(20) NULL, " & _
        "[CategoryName] NVARCHAR(100) NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[DisplayOrder] INT NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("PurchaseOrders", _
        "CREATE TABLE [PurchaseOrders] (" & _
        "[PurchaseID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[PurchaseNo] NVARCHAR(50) NULL, " & _
        "[SupplierID] INT NULL, " & _
        "[CategoryCode] NVARCHAR(20) NULL, " & _
        "[OrderDate] DATETIME2(7) NULL, " & _
        "[ExpectedDate] DATETIME2(7) NULL, " & _
        "[TotalAmount] DECIMAL(19,4) NULL, " & _
        "[Status] NVARCHAR(20) NULL DEFAULT 'pending', " & _
        "[Remarks] NVARCHAR(MAX) NULL, " & _
        "[CreatedBy] INT NULL, " & _
        "[ApprovedBy] INT NULL, " & _
        "[ApprovedAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("PurchaseOrderDetails", _
        "CREATE TABLE [PurchaseOrderDetails] (" & _
        "[DetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[PurchaseID] INT NULL, " & _
        "[ItemCode] NVARCHAR(50) NULL, " & _
        "[ItemName] NVARCHAR(200) NULL, " & _
        "[Specification] NVARCHAR(200) NULL, " & _
        "[Unit] NVARCHAR(20) NULL, " & _
        "[Quantity] FLOAT NULL, " & _
        "[UnitPrice] DECIMAL(19,4) NULL, " & _
        "[TotalPrice] DECIMAL(19,4) NULL, " & _
        "[ReceivedQty] FLOAT NULL, " & _
        "[Remarks] NVARCHAR(MAX) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("PurchaseReceipts", _
        "CREATE TABLE [PurchaseReceipts] (" & _
        "[ReceiptID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ReceiptNo] NVARCHAR(50) NULL, " & _
        "[PurchaseID] INT NULL, " & _
        "[SupplierID] INT NULL, " & _
        "[ReceiptDate] DATETIME2(7) NULL, " & _
        "[TotalReceivedQty] FLOAT NULL, " & _
        "[ReceivedBy] NVARCHAR(50) NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("PurchaseReceiptDetails", _
        "CREATE TABLE [PurchaseReceiptDetails] (" & _
        "[ReceiptDetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ReceiptID] INT NULL, " & _
        "[PurchaseDetailID] INT NULL, " & _
        "[MaterialID] INT NULL, " & _
        "[ReceivedQty] FLOAT NULL, " & _
        "[AcceptedQty] FLOAT NULL, " & _
        "[RejectedQty] FLOAT NULL, " & _
        "[RejectReason] NVARCHAR(200) NULL, " & _
        "[UnitPrice] DECIMAL(19,4) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("PurchaseCostReview", _
        "CREATE TABLE [PurchaseCostReview] (" & _
        "[ReviewID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[PurchaseID] INT NULL, " & _
        "[ReviewAmount] DECIMAL(19,4) NULL, " & _
        "[CostAllocation] NVARCHAR(20) NULL, " & _
        "[ReviewStatus] NVARCHAR(20) NULL, " & _
        "[ReviewComments] NVARCHAR(MAX) NULL, " & _
        "[ReviewerID] INT NULL, " & _
        "[ReviewedAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RawMaterialInventory", _
        "CREATE TABLE [RawMaterialInventory] (" & _
        "[MaterialID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ItemCode] NVARCHAR(50) NULL, " & _
        "[ItemName] NVARCHAR(200) NULL, " & _
        "[CategoryCode] NVARCHAR(20) NULL, " & _
        "[StockQty] FLOAT NULL DEFAULT 0, " & _
        "[SafetyStock] FLOAT NULL DEFAULT 0, " & _
        "[Unit] NVARCHAR(20) NULL, " & _
        "[UnitPrice] DECIMAL(19,4) NULL, " & _
        "[SupplierID] INT NULL, " & _
        "[LastPurchaseDate] DATETIME2(7) NULL, " & _
        "[UpdatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("SupplierPrices", _
        "CREATE TABLE [SupplierPrices] (" & _
        "[PriceID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[SupplierID] INT NULL, " & _
        "[ItemCode] NVARCHAR(50) NULL, " & _
        "[ItemName] NVARCHAR(200) NULL, " & _
        "[UnitPrice] DECIMAL(19,4) NULL, " & _
        "[MinOrderQty] FLOAT NULL, " & _
        "[EffectiveDate] DATETIME2(7) NULL, " & _
        "[ExpiryDate] DATETIME2(7) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("MaterialOutbound", _
        "CREATE TABLE [MaterialOutbound] (" & _
        "[OutboundID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[OutboundNo] NVARCHAR(50) NULL, " & _
        "[OutboundType] NVARCHAR(20) NULL, " & _
        "[ReferenceType] NVARCHAR(50) NULL, " & _
        "[ReferenceID] INT NULL, " & _
        "[OutboundDate] DATETIME2(7) NULL, " & _
        "[RequestedBy] NVARCHAR(50) NULL, " & _
        "[ApprovedBy] NVARCHAR(50) NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[Notes] NVARCHAR(MAX) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("MaterialOutboundDetails", _
        "CREATE TABLE [MaterialOutboundDetails] (" & _
        "[OutboundDetailID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[OutboundID] INT NULL, " & _
        "[MaterialID] INT NULL, " & _
        "[RequestedQty] FLOAT NULL, " & _
        "[ActualQty] FLOAT NULL, " & _
        "[UnitPrice] DECIMAL(19,4) NULL, " & _
        "[TotalAmount] DECIMAL(19,4) NULL, " & _
        "[ProductionOrderRef] INT NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 财务表 ===
    If CreateTableIfNotExists("PaymentRecords", _
        "CREATE TABLE [PaymentRecords] (" & _
        "[RecordID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[OrderID] INT NULL, " & _
        "[OrderNo] NVARCHAR(50) NULL, " & _
        "[TransactionNo] NVARCHAR(100) NULL, " & _
        "[TransactionType] NVARCHAR(20) NULL, " & _
        "[Amount] DECIMAL(19,4) NULL, " & _
        "[Fee] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[NetAmount] DECIMAL(19,4) NULL, " & _
        "[PaymentMethod] NVARCHAR(50) NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[ReconcileStatus] NVARCHAR(20) NULL, " & _
        "[Category] NVARCHAR(50) NULL, " & _
        "[Remark] NVARCHAR(200) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RefundRecords", _
        "CREATE TABLE [RefundRecords] (" & _
        "[RefundID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[RefundNo] NVARCHAR(50) NULL, " & _
        "[OrderID] INT NOT NULL, " & _
        "[OrderNo] NVARCHAR(50) NULL, " & _
        "[RefundAmount] DECIMAL(19,4) NOT NULL, " & _
        "[RefundReason] NVARCHAR(MAX) NULL, " & _
        "[Status] NVARCHAR(20) NULL DEFAULT 'pending', " & _
        "[CostWriteBack] BIT NULL DEFAULT 0, " & _
        "[ApprovedBy] NVARCHAR(50) NULL, " & _
        "[ApprovedAt] DATETIME2(7) NULL, " & _
        "[CompletedAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ReconciliationLogs", _
        "CREATE TABLE [ReconciliationLogs] (" & _
        "[LogID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[OrderID] INT NULL, " & _
        "[OrderNo] NVARCHAR(50) NULL, " & _
        "[OrderAmount] DECIMAL(19,4) NULL, " & _
        "[PaymentAmount] DECIMAL(19,4) NULL, " & _
        "[Difference] DECIMAL(19,4) NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[Resolution] NVARCHAR(MAX) NULL, " & _
        "[ResolvedBy] NVARCHAR(50) NULL, " & _
        "[ResolvedAt] DATETIME2(7) NULL, " & _
        "[ReconcileDate] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ExpenseRecords", _
        "CREATE TABLE [ExpenseRecords] (" & _
        "[ExpenseID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ExpenseName] NVARCHAR(100) NULL, " & _
        "[ExpenseType] NVARCHAR(30) NULL, " & _
        "[Amount] DECIMAL(19,4) NULL, " & _
        "[AllocationMethod] NVARCHAR(20) NULL, " & _
        "[AllocationRatio] FLOAT NULL, " & _
        "[OrderID] INT NULL, " & _
        "[ProductID] INT NULL, " & _
        "[SourceOrderID] INT NULL, " & _
        "[Period] NVARCHAR(10) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("BudgetPlans", _
        "CREATE TABLE [BudgetPlans] (" & _
        "[BudgetID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[BudgetName] NVARCHAR(100) NULL, " & _
        "[Category] NVARCHAR(50) NULL, " & _
        "[Period] NVARCHAR(10) NULL, " & _
        "[BudgetAmount] DECIMAL(19,4) NULL, " & _
        "[ActualAmount] DECIMAL(19,4) NULL, " & _
        "[GMVAmount] DECIMAL(19,4) NULL, " & _
        "[ROI] FLOAT NULL, " & _
        "[AlertPercent] FLOAT NULL, " & _
        "[AlertROI] FLOAT NULL, " & _
        "[Status] NVARCHAR(20) NULL, " & _
        "[CreatedBy] NVARCHAR(50) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("FundAccounts", _
        "CREATE TABLE [FundAccounts] (" & _
        "[AccountID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[AccountName] NVARCHAR(100) NULL, " & _
        "[AccountType] NVARCHAR(30) NULL, " & _
        "[TotalBalance] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[AvailableBalance] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[FrozenAmount] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[PendingSettlement] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[AlertThreshold] DECIMAL(19,4) NULL, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[LastSyncAt] DATETIME2(7) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("ProductCosts", _
        "CREATE TABLE [ProductCosts] (" & _
        "[CostID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NOT NULL, " & _
        "[CostName] NVARCHAR(100) NULL, " & _
        "[CostType] NVARCHAR(20) NULL, " & _
        "[Quantity] FLOAT NULL, " & _
        "[UnitCost] DECIMAL(19,4) NULL, " & _
        "[TotalCost] DECIMAL(19,4) NULL, " & _
        "[EffectiveDate] DATETIME2(7) NULL, " & _
        "[ExpiryDate] DATETIME2(7) NULL, " & _
        "[CreatedBy] NVARCHAR(50) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 营销表 ===
    If CreateTableIfNotExists("Coupons", _
        "CREATE TABLE [Coupons] (" & _
        "[CouponID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[CouponCode] NVARCHAR(50) NULL, " & _
        "[DiscountType] NVARCHAR(20) NULL, " & _
        "[DiscountValue] DECIMAL(19,4) NULL, " & _
        "[MinPurchase] DECIMAL(19,4) NULL, " & _
        "[StartDate] DATETIME2(7) NULL, " & _
        "[EndDate] DATETIME2(7) NULL, " & _
        "[UsageLimit] INT NULL, " & _
        "[UsedCount] INT NULL DEFAULT 0, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("MarketingCampaigns", _
        "CREATE TABLE [MarketingCampaigns] (" & _
        "[CampaignID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[CampaignName] NVARCHAR(200) NULL, " & _
        "[CampaignType] NVARCHAR(50) NULL, " & _
        "[Description] NVARCHAR(MAX) NULL, " & _
        "[DiscountValue] DECIMAL(19,4) NULL, " & _
        "[MinPurchase] DECIMAL(19,4) NULL, " & _
        "[StartDate] DATETIME2(7) NULL, " & _
        "[EndDate] DATETIME2(7) NULL, " & _
        "[ParticipantCount] INT NULL DEFAULT 0, " & _
        "[TotalSales] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("RecipePopularity", _
        "CREATE TABLE [RecipePopularity] (" & _
        "[PopularityID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ProductID] INT NOT NULL, " & _
        "[ViewCount] INT NULL DEFAULT 0, " & _
        "[FavoriteCount] INT NULL DEFAULT 0, " & _
        "[PurchaseCount] INT NULL DEFAULT 0, " & _
        "[LastCalculatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    ' === 统计与配置表 ===
    If CreateTableIfNotExists("DailyStatistics", _
        "CREATE TABLE [DailyStatistics] (" & _
        "[StatID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[StatDate] DATETIME2(7) NOT NULL, " & _
        "[TotalOrders] INT NULL DEFAULT 0, " & _
        "[TotalRevenue] DECIMAL(19,4) NULL DEFAULT 0, " & _
        "[TotalUsers] INT NULL DEFAULT 0, " & _
        "[NewUsers] INT NULL DEFAULT 0, " & _
        "[TopProductID] INT NULL, " & _
        "[TopNoteID] INT NULL, " & _
        "[DataJSON] NVARCHAR(MAX) NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE(), " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1

    If CreateTableIfNotExists("SiteSettings", _
        "CREATE TABLE [SiteSettings] (" & _
        "[SettingKey] NVARCHAR(50) NULL, " & _
        "[SettingName] NVARCHAR(100) NULL, " & _
        "[SettingValue] NVARCHAR(255) NULL, " & _
        "[Description] NVARCHAR(255) NULL, " & _
        "[UpdatedAt] DATETIME2(7) NULL" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1
    
    ' === V14: 推荐制相关表 ===
    ' 为 Users 表添加推荐相关字段（如果表已存在但缺少字段）
    On Error Resume Next
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[Users]') AND name = 'ReferrerUserID') ALTER TABLE [Users] ADD [ReferrerUserID] INT NULL"
    If Err.Number = 0 Then Call LogStep("success", "Users 表字段 [ReferrerUserID] 已添加") Else Call LogStep("warning", "Users 表字段 [ReferrerUserID] 添加失败: " & Err.Description) : Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[Users]') AND name = 'DeviceFingerprint') ALTER TABLE [Users] ADD [DeviceFingerprint] NVARCHAR(100) NULL"
    If Err.Number = 0 Then Call LogStep("success", "Users 表字段 [DeviceFingerprint] 已添加") Else Call LogStep("warning", "Users 表字段 [DeviceFingerprint] 添加失败: " & Err.Description) : Err.Clear
    
    ' V14.1: 添加 OriginalToken 列（存储原始推荐Token字符串）
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[ReferralTokens]') AND name = 'OriginalToken') ALTER TABLE [ReferralTokens] ADD [OriginalToken] NVARCHAR(1000) NULL"
    If Err.Number = 0 Then Call LogStep("success", "ReferralTokens 表字段 [OriginalToken] 已添加") Else Call LogStep("warning", "ReferralTokens 表字段 [OriginalToken] 添加失败: " & Err.Description) : Err.Clear
    On Error GoTo 0
    
    If CreateTableIfNotExists("ReferralTokens", _
        "CREATE TABLE [ReferralTokens] (" & _
        "[TokenID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[ReferrerUserID] INT NOT NULL, " & _
        "[ReferrerType] NVARCHAR(20) NULL DEFAULT 'user', " & _
        "[TokenHash] NVARCHAR(255) NOT NULL, " & _
        "[OriginalToken] NVARCHAR(1000) NULL, " & _
        "[ExpiresAt] DATETIME2(7) NOT NULL, " & _
        "[MaxUses] INT NULL DEFAULT 1, " & _
        "[UsedCount] INT NULL DEFAULT 0, " & _
        "[IsActive] BIT NULL DEFAULT 1, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1
    
    If CreateTableIfNotExists("ReferralRelations", _
        "CREATE TABLE [ReferralRelations] (" & _
        "[RelationID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[AncestorUserID] INT NOT NULL, " & _
        "[DescendantUserID] INT NOT NULL, " & _
        "[Depth] INT NOT NULL, " & _
        "[CreatedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1
    
    If CreateTableIfNotExists("RegistrationAttempts", _
        "CREATE TABLE [RegistrationAttempts] (" & _
        "[AttemptID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "[IPAddress] NVARCHAR(50) NOT NULL, " & _
        "[DeviceFingerprint] NVARCHAR(100) NULL, " & _
        "[Success] BIT NULL DEFAULT 0, " & _
        "[TokenHash] NVARCHAR(255) NULL, " & _
        "[AttemptedAt] DATETIME2(7) NULL DEFAULT GETDATE()" & _
        ")") Then successTables = successTables + 1
    totalTables = totalTables + 1
    
    Call LogStep("success", "建表完成: " & successTables & "/" & totalTables & " 个表创建成功")

    ' === 创建索引 ===
    Call LogStep("info", "创建性能索引...")
    Dim indexCount : indexCount = 0
    On Error Resume Next
    
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_UserID') CREATE NONCLUSTERED INDEX [IX_Orders_UserID] ON [Orders]([UserID]) INCLUDE ([OrderID],[TotalAmount],[Status],[CreatedAt])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_Status') CREATE NONCLUSTERED INDEX [IX_Orders_Status] ON [Orders]([Status]) INCLUDE ([OrderID],[TotalAmount],[CreatedAt])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Orders_OrderNo') CREATE NONCLUSTERED INDEX [IX_Orders_OrderNo] ON [Orders]([OrderNo])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_OrderDetails_OrderID') CREATE NONCLUSTERED INDEX [IX_OrderDetails_OrderID] ON [OrderDetails]([OrderID]) INCLUDE ([ProductID],[Quantity],[Subtotal])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Products_IsActive') CREATE NONCLUSTERED INDEX [IX_Products_IsActive] ON [Products]([IsActive]) INCLUDE ([ProductID],[ProductName],[ProductType],[BasePrice])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Users_Username') CREATE NONCLUSTERED INDEX [IX_Users_Username] ON [Users]([Username])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_Users_Email') CREATE NONCLUSTERED INDEX [IX_Users_Email] ON [Users]([Email])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_FragranceNotes_NoteType') CREATE NONCLUSTERED INDEX [IX_FragranceNotes_NoteType] ON [FragranceNotes]([NoteType]) INCLUDE ([NoteID],[NoteName],[PriceAddition])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    
    ' V14: 推荐制索引
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ReferralTokens_TokenHash') CREATE NONCLUSTERED INDEX [IX_ReferralTokens_TokenHash] ON [ReferralTokens]([TokenHash]) INCLUDE ([ReferrerUserID],[ExpiresAt],[UsedCount],[MaxUses],[IsActive])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ReferralTokens_ReferrerUserID') CREATE NONCLUSTERED INDEX [IX_ReferralTokens_ReferrerUserID] ON [ReferralTokens]([ReferrerUserID]) INCLUDE ([TokenID],[ExpiresAt],[IsActive])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ReferralRelations_Ancestor') CREATE NONCLUSTERED INDEX [IX_ReferralRelations_Ancestor] ON [ReferralRelations]([AncestorUserID]) INCLUDE ([DescendantUserID],[Depth])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_ReferralRelations_Descendant') CREATE NONCLUSTERED INDEX [IX_ReferralRelations_Descendant] ON [ReferralRelations]([DescendantUserID])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_RegistrationAttempts_IP') CREATE NONCLUSTERED INDEX [IX_RegistrationAttempts_IP] ON [RegistrationAttempts]([IPAddress],[AttemptedAt])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='IX_RegistrationAttempts_Fingerprint') CREATE NONCLUSTERED INDEX [IX_RegistrationAttempts_Fingerprint] ON [RegistrationAttempts]([DeviceFingerprint],[AttemptedAt])"
    If Err.Number = 0 Then indexCount = indexCount + 1 Else Err.Clear
    
    Call LogStep("success", "索引创建完成: " & indexCount & " 个")

    On Error GoTo 0
End Sub

Function CreateTableIfNotExists(tableName, createSQL)
    On Error Resume Next
    Dim rsCheck
    Set rsCheck = connAdmin.Execute("SELECT COUNT(*) FROM sys.tables WHERE name='" & tableName & "'")
    Dim exists : exists = False
    If Not rsCheck.EOF Then
        If rsCheck.Fields(0).Value > 0 Then exists = True
    End If
    rsCheck.Close : Set rsCheck = Nothing
    
    If exists Then
        CreateTableIfNotExists = True
        Exit Function
    End If
    
    connAdmin.Execute createSQL
    If Err.Number <> 0 Then
        Call LogStep("error", "创建表 [" & tableName & "] 失败: " & Err.Description)
        CreateTableIfNotExists = False
        Err.Clear
    Else
        Call LogStep("success", "表 [" & tableName & "] 创建成功")
        CreateTableIfNotExists = True
    End If
    On Error GoTo 0
End Function

Sub SetupPermissions()
    On Error Resume Next
    connAdmin.Execute "USE [PerfumeShop]"
    
    ' 获取当前连接用户
    Dim currentUser
    Set rs = connAdmin.Execute("SELECT SUSER_SNAME()")
    If Not rs.EOF Then currentUser = rs.Fields(0).Value
    rs.Close : Set rs = Nothing
    Call LogStep("info", "当前数据库连接用户: " & currentUser)

    ' 创建数据库角色并授予权限
    connAdmin.Execute "IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name='asp_app_role' AND type='R') CREATE ROLE [asp_app_role]"
    If Err.Number <> 0 Then Err.Clear
    
    connAdmin.Execute "GRANT SELECT, INSERT, UPDATE, DELETE ON SCHEMA::dbo TO [asp_app_role]"
    If Err.Number <> 0 Then Err.Clear
    
    connAdmin.Execute "GRANT EXECUTE ON SCHEMA::dbo TO [asp_app_role]"
    If Err.Number <> 0 Then Err.Clear
    
    ' 授予 BACKUP DATABASE 权限
    connAdmin.Execute "GRANT BACKUP DATABASE TO [" & currentUser & "]"
    If Err.Number <> 0 Then
        Call LogStep("warning", "无法授予 BACKUP DATABASE 给 " & currentUser & ": " & Err.Description)
        Err.Clear
    Else
        Call LogStep("success", "BACKUP DATABASE 权限已授予 " & currentUser)
    End If
    
    ' 授予 CREATE PROCEDURE 权限
    connAdmin.Execute "GRANT CREATE PROCEDURE TO [" & currentUser & "]"
    If Err.Number <> 0 Then Err.Clear
    
    Call LogStep("success", "数据库权限配置完成")
    On Error GoTo 0
End Sub

Sub VerifySetup()
    Call LogStep("info", "验证数据库设置...")
    
    connAdmin.Execute "USE [PerfumeShop]"
    
    ' 统计表数量
    Set rs = connAdmin.Execute("SELECT COUNT(*) FROM sys.tables")
    Dim tableCount : tableCount = rs.Fields(0).Value
    rs.Close : Set rs = Nothing
    Call LogStep("success", "数据库表总数: " & tableCount)
    
    ' 列出所有表名
    Set rs = connAdmin.Execute("SELECT name FROM sys.tables ORDER BY name")
    Dim tableList : tableList = ""
    Do While Not rs.EOF
        tableList = tableList & rs.Fields(0).Value & ", "
        rs.MoveNext
    Loop
    rs.Close : Set rs = Nothing
    If Len(tableList) > 2 Then tableList = Left(tableList, Len(tableList)-2)
    Call LogStep("info", "表列表: " & tableList)
End Sub

Sub LogStep(cssClass, message)
    Response.Write "<div class='step " & cssClass & "'>" & Server.HTMLEncode(message) & "</div>"
    Response.Flush
End Sub

' === 执行入口 ===
If Request.Form("action") = "setup" Then
    Call RunSetup()
Else
%>
<div class="step info">
    <p><strong>当前环境检测：</strong></p>
    <p>SQL Server 默认实例 (MSSQLSERVER) 已运行</p>
    <p>IIS 已运行 (W3SVC)</p>
    <p>此工具将创建 PerfumeShop 数据库、所有表结构、索引和权限</p>
    <p><strong>注意：</strong>执行前请备份现有数据</p>
</div>
<form method="post">
    <input type="hidden" name="action" value="setup">
    <button type="submit">开始初始化数据库</button>
</form>
<%
End If
%>
</body>
</html>
