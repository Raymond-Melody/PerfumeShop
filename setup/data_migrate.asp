<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.Charset = "UTF-8"
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>数据迁移: Access -> SQL Server</title>
<style>
body{font-family:Arial,sans-serif;max-width:1000px;margin:20px auto;padding:20px;background:#f5f5f5}
h1{color:#333;border-bottom:2px solid #FF9800;padding-bottom:10px}
.step{margin:8px 0;padding:10px 15px;border-radius:5px}
.success{background:#d4edda;color:#155724;border:1px solid #c3e6cb}
.error{background:#f8d7da;color:#721c24;border:1px solid #f5c6cb}
.info{background:#d1ecf1;color:#0c5460;border:1px solid #bee5eb}
.warning{background:#fff3cd;color:#856404;border:1px solid #ffeeba}
table{width:100%;border-collapse:collapse;margin:10px 0;font-size:13px}
th,td{border:1px solid #ddd;padding:6px 10px;text-align:left}
th{background:#f2f2f2}
button{padding:10px 30px;font-size:16px;background:#FF9800;color:#fff;border:none;border-radius:5px;cursor:pointer;margin:5px}
button:disabled{background:#ccc;cursor:not-allowed}
button.green{background:#4CAF50}
.progress-bar{width:100%;background:#ddd;border-radius:4px;margin:10px 0}
.progress-fill{height:20px;border-radius:4px;background:#4CAF50;text-align:center;color:#fff;line-height:20px;font-size:12px}
</style>
</head>
<body>
<h1>数据迁移工具: Access (PerfumeShop.mdb) → SQL Server</h1>
<%
Dim connAcc, connSQL, rs, sql

' === 获取表列表 ===
Function GetAccessTables(conn)
    Dim tableRS, tables
    Set tableRS = conn.OpenSchema(20) ' adSchemaTables
    tables = Array()
    Do While Not tableRS.EOF
        If tableRS.Fields("TABLE_TYPE").Value = "TABLE" Then
            ' Exclude system tables
            If Not (Left(tableRS.Fields("TABLE_NAME").Value, 4) = "MSys" Or Left(tableRS.Fields("TABLE_NAME").Value, 1) = "~") Then
                ReDim Preserve tables(UBound(tables) + 1)
                tables(UBound(tables)) = tableRS.Fields("TABLE_NAME").Value
            End If
        End If
        tableRS.MoveNext
    Loop
    tableRS.Close : Set tableRS = Nothing
    GetAccessTables = tables
End Function

Function GetSQLTables(conn)
    Dim rsT, tables, i
    tables = Array()
    Set rsT = conn.Execute("SELECT name FROM sys.tables ORDER BY name")
    i = -1
    Do While Not rsT.EOF
        i = i + 1
        ReDim Preserve tables(i)
        tables(i) = rsT.Fields(0).Value
        rsT.MoveNext
    Loop
    rsT.Close : Set rsT = Nothing
    GetSQLTables = tables
End Function

' === 获取表列信息 ===
Function GetColumnInfo(conn, tableName, isSQL)
    Dim rsSch, cols, i
    cols = Array()
    i = -1
    If isSQL Then
        Set rsSch = conn.Execute("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" & tableName & "' ORDER BY ORDINAL_POSITION")
    Else
        Set rsSch = conn.OpenSchema(4, Array(Empty, Empty, tableName)) ' adSchemaColumns
    End If
    
    Do While Not rsSch.EOF
        i = i + 1
        ReDim Preserve cols(i)
        If isSQL Then
            cols(i) = rsSch.Fields("COLUMN_NAME").Value
        Else
            cols(i) = rsSch.Fields("COLUMN_NAME").Value
        End If
        rsSch.MoveNext
    Loop
    rsSch.Close : Set rsSch = Nothing
    GetColumnInfo = cols
End Function

Sub RunMigration()
    Dim accessTables, sqlTables, tableName, colNames, i, j, accCols, startCol
    Dim totalRows, successRows, failedTables, successTables
    Dim colList, valList, colName, val
    
    On Error Resume Next
    
    ' 连接 Access
    Call LogStep("info", "Connecting to Access database...")
    Set connAcc = Server.CreateObject("ADODB.Connection")
    connAcc.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("../database/PerfumeShop.mdb") & ";"
    If Err.Number <> 0 Then
        Err.Clear
        ' 尝试旧版驱动
        connAcc.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../database/PerfumeShop.mdb") & ";"
        If Err.Number <> 0 Then
            Call LogStep("error", "Cannot connect Access: " & Err.Description & ". Install ACE.OLEDB.12.0 driver.")
            Exit Sub
        End If
    End If
    Call LogStep("success", "Access connected OK")
    
    ' 连接 SQL Server
    Call LogStep("info", "Connecting to SQL Server...")
    Set connSQL = Server.CreateObject("ADODB.Connection")
    connSQL.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"
    If Err.Number <> 0 Then
        Call LogStep("error", "Cannot connect SQL Server: " & Err.Description)
        connAcc.Close : Set connAcc = Nothing
        Exit Sub
    End If
    Call LogStep("success", "SQL Server connected OK")
    
    ' 获取两个数据库的表列表
    accessTables = GetAccessTables(connAcc)
    sqlTables = GetSQLTables(connSQL)
    
    Call LogStep("info", "Access tables: " & (UBound(accessTables)+1) & " | SQL Server tables: " & (UBound(sqlTables)+1))
    
    ' 迁移顺序：先迁无外键依赖的基础表
    Dim migrationOrder
    migrationOrder = Array( _
        "Users", "AdminRoles", "AdminUsers", "Categories", "ProductTypeConfig", _
        "Volumes", "FragranceNotes", "BaseNotes", "BottleStyles", "Ingredients", _
        "Products", "ProductVolumePrices", "ProductNotes", "ProductNoteRatios", _
        "ProductBottleStyles", "FragranceIngredients", "FormulaNotes", "Formulas", _
        "NoteIngredients", "NoteInventory", "ProductInventory", _
        "Recipes", "RecipeNotes", "RecipeIngredients", "RecipeProducts", "RecipeProductNotes", _
        "RecipeAccords", "RecipeAccordMaterials", "RecipePublishLog", "RecommendedRecipes", _
        "RecipePopularity", _
        "Suppliers", "PurchaseCategories", "SupplierPrices", _
        "PurchaseOrders", "PurchaseOrderDetails", "PurchaseReceipts", "PurchaseReceiptDetails", _
        "PurchaseCostReview", "RawMaterialInventory", _
        "MaterialOutbound", "MaterialOutboundDetails", _
        "ProductionOrders", "ProductionLogs", "ProductManufacturing", "ProductManufacturingDetails", _
        "AccordProductions", "AccordProductionDetails", "AccordQCReports", _
        "WorkshopTransfer", _
        "Orders", "OrderDetails", "OrderDetailNoteSelections", "OrderIngredients", _
        "Cart", "CartNoteSelections", _
        "UserAddresses", "UserFavorites", "UserPoints", "PointTransactions", _
        "UserPreferences", "ProductReviews", _
        "PaymentRecords", "RefundRecords", "ReconciliationLogs", _
        "ExpenseRecords", "BudgetPlans", "FundAccounts", "ProductCosts", _
        "Coupons", "MarketingCampaigns", _
        "ModulePermissions", "AdminLogs", _
        "DailyStatistics", "SiteSettings" _
    )
    
    totalRows = 0
    successRows = 0
    failedTables = 0
    successTables = 0
    
    For i = 0 To UBound(migrationOrder)
        tableName = migrationOrder(i)
        
        ' 检查 SQL 中是否有此表
        Dim foundInSQL : foundInSQL = False
        For j = 0 To UBound(sqlTables)
            If UCase(sqlTables(j)) = UCase(tableName) Then
                foundInSQL = True
                Exit For
            End If
        Next
        
        If Not foundInSQL Then
            Call LogStep("warning", "Skip [" & tableName & "]: table not in SQL Server")
        Else
            ' 检查 Access 中是否有此表且有数据
            Dim foundInAcc : foundInAcc = False
            For j = 0 To UBound(accessTables)
                If UCase(accessTables(j)) = UCase(tableName) Then
                    foundInAcc = True
                    Exit For
                End If
            Next
            
            If Not foundInAcc Then
                Call LogStep("info", "Skip [" & tableName & "]: table not in Access")
            Else
                ' 检查 SQL 表是否已有数据
                Dim rsCount
                Set rsCount = connSQL.Execute("SELECT COUNT(*) FROM [" & tableName & "]")
                Dim existingCount : existingCount = rsCount.Fields(0).Value
                rsCount.Close : Set rsCount = Nothing
                
                If existingCount > 0 Then
                    Call LogStep("info", "Skip [" & tableName & "]: already has " & existingCount & " rows")
                Else
                    ' 读 Access 数据
                    Dim rsAcc
                    Set rsAcc = connAcc.Execute("SELECT * FROM [" & tableName & "]")
                    
                    If rsAcc.EOF Then
                        Call LogStep("info", "Skip [" & tableName & "]: no data in Access")
                        rsAcc.Close : Set rsAcc = Nothing
                    Else
                        ' 获取列名
                        Dim colCount : colCount = rsAcc.Fields.Count - 1
                        ReDim accCols(colCount)
                        For j = 0 To colCount
                            accCols(j) = rsAcc.Fields(j).Name
                        Next
                        
                        ' 禁用/启用 IDENTITY_INSERT
                        Dim hasIdentity : hasIdentity = True
                        connSQL.Execute "SET IDENTITY_INSERT [" & tableName & "] ON"
                        If Err.Number <> 0 Then
                            hasIdentity = False
                            Err.Clear
                        End If
                        
                        ' 如果无法启用 IDENTITY_INSERT，跳过第一列（标识列）
                        startCol = 0
                        If Not hasIdentity Then startCol = 1
                        
                        ' 逐行插入
                        Dim rowCount : rowCount = 0
                        Dim errorCount : errorCount = 0
                        
                        Do While Not rsAcc.EOF
                            colList = ""
                            valList = ""
                            
                            For j = startCol To colCount
                                If colList <> "" Then colList = colList & ", "
                                colList = colList & "[" & accCols(j) & "]"
                                
                                If valList <> "" Then valList = valList & ", "
                                val = rsAcc.Fields(j).Value
                                
                                If IsNull(val) Then
                                    valList = valList & "NULL"
                                Else
                                    Select Case VarType(val)
                                        Case 2,3,4,5,6 ' 数字类型
                                            valList = valList & Replace(CStr(val), ",", ".")
                                        Case 7 ' Date
                                            valList = valList & "'" & FormatDateTime(val, 2) & " " & FormatDateTime(val, 4) & "'"
                                        Case 11 ' Boolean
                                            valList = valList & IIf(val, "1", "0")
                                        Case Else ' String
                                            valList = valList & "N'" & Replace(Replace(CStr(val), "'", "''"), "\", "\\") & "'"
                                    End Select
                                End If
                            Next
                            
                            On Error Resume Next
                            connSQL.Execute "INSERT INTO [" & tableName & "] (" & colList & ") VALUES (" & valList & ")"
                            If Err.Number <> 0 Then
                                errorCount = errorCount + 1
                                If errorCount <= 3 Then
                                    ' 只显示前3个错误
                                    Call LogStep("warning", "[" & tableName & "] row insert error: " & Err.Description)
                                End If
                                Err.Clear
                            Else
                                rowCount = rowCount + 1
                            End If
                            Err.Clear
                            
                            rsAcc.MoveNext
                        Loop
                        
                        If hasIdentity Then
                            connSQL.Execute "SET IDENTITY_INSERT [" & tableName & "] OFF"
                            Err.Clear
                        End If
                        
                        totalRows = totalRows + rowCount
                        successRows = successRows + rowCount
                        
                        Dim statusClass
                        If errorCount = 0 Then
                            statusClass = "success"
                            successTables = successTables + 1
                        Else
                            statusClass = "warning"
                        End If
                        
                        Call LogStep(statusClass, "[" & tableName & "]: Migrated " & rowCount & " rows" & IIf(errorCount>0, " (skipped " & errorCount & " errors)", ""))
                        
                        rsAcc.Close : Set rsAcc = Nothing
                    End If
                End If
            End If
        End If
    Next
    
    Call LogStep("success", "========================================")
    Call LogStep("success", "Migration complete! Total: " & totalRows & " rows, " & successTables & " tables")
    Call LogStep("success", "========================================")
    
    connAcc.Close : Set connAcc = Nothing
    connSQL.Close : Set connSQL = Nothing
End Sub

Function IIf(cond, tv, fv)
    If cond Then IIf = tv Else IIf = fv
End Function

Sub LogStep(cssClass, message)
    Response.Write "<div class='step " & cssClass & "'>" & Server.HTMLEncode(message) & "</div>"
    Response.Flush
End Sub

If Request.Form("action") = "migrate" Then
    Call RunMigration()
Else
%>
<div class="step info">
    <p><strong>数据迁移说明：</strong></p>
    <p>此工具将从 Access 数据库 (database/PerfumeShop.mdb) 读取数据，逐表迁移到 SQL Server。</p>
    <p><strong>前提条件：</strong></p>
    <ul>
        <li>SQL Server PerfumeShop 数据库已通过 db_setup.asp 创建</li>
        <li>服务器安装了 Microsoft ACE OLEDB 12.0 驱动</li>
        <li>IIS 应用程序池对 database 文件夹有读取权限</li>
    </ul>
    <p><strong>注意：</strong>已存在数据的表将被跳过，不会重复导入。</p>
</div>
<form method="post">
    <input type="hidden" name="action" value="migrate">
    <button type="submit" class="green">开始数据迁移</button>
</form>
<%
End If
%>
</body>
</html>
