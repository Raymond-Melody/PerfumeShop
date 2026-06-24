<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<%
' ============================================
' V11 智能补货Schema初始化页面
' 从 replenishment.asp 中提取，避免阻塞正常页面加载
' 管理员首次使用智能补货时手动执行一次即可
' ============================================
Call OpenConnection()
Server.ScriptTimeout = 120
conn.CommandTimeout = 60

Dim resultMsg, resultType, details
resultMsg = ""
resultType = "success"
details = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        resultMsg = "安全令牌验证失败，请刷新页面后重试"
        resultType = "error"
    Else
        Dim migSQL
        migSQL = ""
        ' --- RawMaterialInventory 字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='AvgDailyUsage') ALTER TABLE RawMaterialInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='LeadTimeDays') ALTER TABLE RawMaterialInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='LastReplenishDate') ALTER TABLE RawMaterialInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('RawMaterialInventory') AND name='ReorderPoint') ALTER TABLE RawMaterialInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- PackagingInventory 字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='AvgDailyUsage') ALTER TABLE PackagingInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='LeadTimeDays') ALTER TABLE PackagingInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='LastReplenishDate') ALTER TABLE PackagingInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PackagingInventory') AND name='ReorderPoint') ALTER TABLE PackagingInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- BottleStyles 字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='AvgDailyUsage') ALTER TABLE BottleStyles ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='LeadTimeDays') ALTER TABLE BottleStyles ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='LastReplenishDate') ALTER TABLE BottleStyles ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('BottleStyles') AND name='ReorderPoint') ALTER TABLE BottleStyles ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- PrintingInventory 建表+字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='PrintingInventory') CREATE TABLE PrintingInventory (PrintingID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), ItemCode NVARCHAR(50), StockQty DECIMAL(10,2) DEFAULT 0, SafetyStock DECIMAL(10,2) DEFAULT 0, Unit NVARCHAR(20) DEFAULT N'张', UnitPrice DECIMAL(10,2) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()); "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='AvgDailyUsage') ALTER TABLE PrintingInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='LeadTimeDays') ALTER TABLE PrintingInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='LastReplenishDate') ALTER TABLE PrintingInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('PrintingInventory') AND name='ReorderPoint') ALTER TABLE PrintingInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- SprayHeadInventory 建表+字段 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='SprayHeadInventory') CREATE TABLE SprayHeadInventory (SprayHeadID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100), ItemCode NVARCHAR(50), StockQty DECIMAL(10,2) DEFAULT 0, SafetyStock DECIMAL(10,2) DEFAULT 0, Unit NVARCHAR(20) DEFAULT N'个', UnitPrice DECIMAL(10,2) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()); "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='AvgDailyUsage') ALTER TABLE SprayHeadInventory ADD AvgDailyUsage DECIMAL(19,6) DEFAULT 0; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='LeadTimeDays') ALTER TABLE SprayHeadInventory ADD LeadTimeDays INT DEFAULT 7; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='LastReplenishDate') ALTER TABLE SprayHeadInventory ADD LastReplenishDate DATETIME; "
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('SprayHeadInventory') AND name='ReorderPoint') ALTER TABLE SprayHeadInventory ADD ReorderPoint DECIMAL(19,4) DEFAULT 0; "
        ' --- PurchaseHistoryStats 建表 ---
        migSQL = migSQL & "IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='PurchaseHistoryStats') CREATE TABLE PurchaseHistoryStats (StatID INT IDENTITY(1,1) PRIMARY KEY, ItemType NVARCHAR(30) NOT NULL, ItemCode NVARCHAR(100), ItemName NVARCHAR(200), Avg30DayUsage DECIMAL(19,6) DEFAULT 0, Avg90DayUsage DECIMAL(19,6) DEFAULT 0, LastOrderDate DATETIME, TotalOrders90Days INT DEFAULT 0, PreferredSupplierID INT, PreferredUnitPrice DECIMAL(19,4) DEFAULT 0, UpdatedAt DATETIME DEFAULT GETDATE()); "

        On Error Resume Next
        conn.Execute migSQL
        If Err.Number <> 0 Then
            resultMsg = "Schema初始化失败: " & Server.HTMLEncode(Err.Description)
            resultType = "error"
        Else
            resultMsg = "Schema初始化成功！智能补货功能已就绪。"
            resultType = "success"
            Session("ReplenishSchemaReady") = "1"
        End If
        On Error GoTo 0
    End If
End If

' 检查当前状态
Dim schemaStatus
schemaStatus = "未知"
On Error Resume Next
conn.Execute "SELECT TOP 1 StatID FROM PurchaseHistoryStats WHERE 1=0"
If Err.Number = 0 Then
    schemaStatus = "<span style='color:#4CAF50;'><i class='fas fa-check-circle'></i> Schema已就绪</span>"
Else
    schemaStatus = "<span style='color:#e74c3c;'><i class='fas fa-exclamation-triangle'></i> Schema未就绪（需要初始化）</span>"
    Err.Clear
End If
On Error GoTo 0

Call CloseConnection()
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>智能补货Schema初始化 - 采购中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #FF9800; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 0; }
        .container { max-width: 700px; margin: 80px auto; padding: 30px; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); }
        h1 { color: var(--accent); font-size: 22px; margin-bottom: 10px; }
        .status-box { padding: 16px; background: rgba(255,255,255,0.03); border-radius: 8px; margin: 15px 0; }
        .alert { padding: 12px 18px; border-radius: 8px; margin-bottom: 15px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #4CAF50; border-left: 3px solid #4CAF50; }
        .alert-error { background: rgba(231,76,60,0.15); color: #e74c3c; border-left: 3px solid #e74c3c; }
        .btn { display: inline-block; padding: 10px 24px; border-radius: 8px; cursor: pointer; text-decoration: none; font-size: 14px; border: none; }
        .btn-primary { background: var(--accent); color: #fff; }
        .btn-outline { background: transparent; border: 1px solid rgba(255,255,255,0.2); color: #aaa; }
        .info-text { color: #888; font-size: 13px; line-height: 1.6; margin: 15px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1><i class="fas fa-cog"></i> 智能补货Schema初始化</h1>
        <p class="info-text">此页面用于初始化智能补货功能所需的数据库表和字段。操作是幂等的（重复执行不会出错），只需执行一次。</p>
        
        <% If resultMsg <> "" Then %>
        <div class="alert alert-<%=resultType%>"><%=resultMsg%></div>
        <% End If %>
        
        <div class="status-box">
            <strong>当前状态：</strong><%=schemaStatus%>
        </div>
        
        <div class="info-text">
            <strong>需要创建的Schema：</strong>
            <ul style="margin-top:8px;padding-left:20px;">
                <li>RawMaterialInventory: AvgDailyUsage, LeadTimeDays, LastReplenishDate, ReorderPoint</li>
                <li>PackagingInventory: AvgDailyUsage, LeadTimeDays, LastReplenishDate, ReorderPoint</li>
                <li>BottleStyles: AvgDailyUsage, LeadTimeDays, LastReplenishDate, ReorderPoint</li>
                <li>PrintingInventory: 新表 + 补货字段</li>
                <li>SprayHeadInventory: 新表 + 补货字段</li>
                <li>PurchaseHistoryStats: 新表（采购历史统计）</li>
            </ul>
        </div>
        
        <form method="post">
            <input type="hidden" name="action" value="init_schema">
            <%=GetCSRFTokenField()%>
            <button type="submit" class="btn btn-primary" onclick="return confirm('确认执行Schema初始化？此操作可能需要30-60秒。')">
                <i class="fas fa-play"></i> 执行Schema初始化
            </button>
            <a href="replenishment.asp" class="btn btn-outline" style="margin-left:10px;">
                <i class="fas fa-arrow-left"></i> 返回智能补货
            </a>
        </form>
    </div>
</body>
</html>
