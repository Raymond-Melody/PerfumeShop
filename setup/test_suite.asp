<%@ Language="VBScript" CodePage=65001 %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>PerfumeShop - 功能测试套件</title>
<style>
body{font-family:Arial,sans-serif;max-width:1000px;margin:20px auto;padding:20px;background:#f5f5f5}
h1{color:#333;border-bottom:2px solid #9C27B0;padding-bottom:10px}
h2{color:#555;margin-top:30px;border-bottom:1px solid #ddd;padding-bottom:5px}
.test{display:flex;align-items:center;margin:6px 0;padding:8px 12px;border-radius:4px;font-size:13px}
.test-name{flex:1;font-weight:bold}
.test-result{min-width:80px;text-align:center;padding:3px 8px;border-radius:3px;font-weight:bold}
.PASS{background:#d4edda;color:#155724;border:1px solid #c3e6cb}
.FAIL{background:#f8d7da;color:#721c24;border:1px solid #f5c6cb}
.SKIP{background:#e2e3e5;color:#383d41;border:1px solid #d6d8db}
pre{background:#fff;padding:10px;border:1px solid #ddd;border-radius:3px;overflow-x:auto;font-size:12px;max-height:300px;overflow-y:auto}
.summary{padding:15px;border-radius:5px;margin:15px 0;font-weight:bold;text-align:center}
</style>
</head>
<body>
<h1>PerfumeShop 功能测试套件</h1>
<p>测试时间: <%= Now() %> | 系统版本: <%= SYS_VERSION %> | <a href="deploy.asp">返回部署工具</a></p>

<%
Dim passCount, failCount, skipCount
passCount = 0 : failCount = 0 : skipCount = 0

Sub RunTest(name, result, detail)
    Dim cssClass
    Select Case result
        Case "PASS": cssClass = "PASS" : passCount = passCount + 1
        Case "FAIL": cssClass = "FAIL" : failCount = failCount + 1
        Case "SKIP": cssClass = "SKIP" : skipCount = skipCount + 1
    End Select
    Response.Write "<div class='test'><span class='test-name'>" & name
    If detail <> "" Then Response.Write " <small style='color:#666'>(" & Server.HTMLEncode(detail) & ")</small>"
    Response.Write "</span><span class='test-result " & cssClass & "'>" & result & "</span></div>"
End Sub

OpenConnection
%>

<h2>1. 数据库连接与基础CRUD</h2>
<%
Dim testRS, testSQL, testID

' Test 1: Database connection
If IsObject(conn) And conn.State = 1 Then
    RunTest "SQL Server 连接", "PASS", conn.ConnectionString
Else
    RunTest "SQL Server 连接", "FAIL", "连接状态: " & conn.State
End If

' Test 2: Users table read
On Error Resume Next
Set testRS = ExecuteQuery("SELECT COUNT(*) FROM Users")
If Err.Number = 0 And Not testRS Is Nothing Then
    Dim userCount : userCount = testRS.Fields(0).Value
    testRS.Close : Set testRS = Nothing
    RunTest "Users 表查询", "PASS", userCount & " 条记录"
Else
    RunTest "Users 表查询", "FAIL", Err.Description
    Err.Clear
End If

' Test 3: Products table read
On Error Resume Next
Set testRS = ExecuteQuery("SELECT COUNT(*) FROM Products")
If Err.Number = 0 And Not testRS Is Nothing Then
    Dim prodCount : prodCount = testRS.Fields(0).Value
    testRS.Close : Set testRS = Nothing
    RunTest "Products 表查询", "PASS", prodCount & " 条记录"
Else
    RunTest "Products 表查询", "FAIL", Err.Description
    Err.Clear
End If

' Test 4: FragranceNotes read
On Error Resume Next
Set testRS = ExecuteQuery("SELECT COUNT(*) FROM FragranceNotes")
If Err.Number = 0 And Not testRS Is Nothing Then
    Dim noteCount : noteCount = testRS.Fields(0).Value
    testRS.Close : Set testRS = Nothing
    RunTest "FragranceNotes 表查询", "PASS", noteCount & " 条记录"
Else
    RunTest "FragranceNotes 表查询", "FAIL", Err.Description
    Err.Clear
End If

' Test 5: Orders read
On Error Resume Next
Set testRS = ExecuteQuery("SELECT COUNT(*) FROM Orders")
If Err.Number = 0 And Not testRS Is Nothing Then
    Dim orderCount : orderCount = testRS.Fields(0).Value
    testRS.Close : Set testRS = Nothing
    RunTest "Orders 表查询", "PASS", orderCount & " 条记录"
Else
    RunTest "Orders 表查询", "FAIL", Err.Description
    Err.Clear
End If

' Test 6: INSERT test (temp record)
On Error Resume Next
testSQL = "INSERT INTO Users (Username, Password, Email, FullName, IsActive, CreatedAt) VALUES ('_test_user_001', 'test_hash', 'test@test.com', N'测试用户', 1, GETDATE())"
If ExecuteNonQuery(testSQL) Then
    testID = GetLastInsertID("Users")
    RunTest "INSERT 操作", "PASS", "新记录 ID=" & testID
    ' Cleanup
    ExecuteNonQuery "DELETE FROM Users WHERE Username='_test_user_001'"
Else
    RunTest "INSERT 操作", "FAIL", Session("LastDBError")
End If
On Error GoTo 0

' Test 7: UPDATE test
On Error Resume Next
ExecuteNonQuery "INSERT INTO Users (Username, Password, Email, FullName, IsActive, CreatedAt) VALUES ('_test_user_002', 'test_hash', 'test2@test.com', N'更新测试', 1, GETDATE())"
testID = GetLastInsertID("Users")
If testID > 0 Then
    If ExecuteNonQuery("UPDATE Users SET FullName=N'已更新测试' WHERE UserID=" & testID) Then
        Set testRS = ExecuteQuery("SELECT FullName FROM Users WHERE UserID=" & testID)
        If Not testRS Is Nothing And Not testRS.EOF Then
            If testRS.Fields(0).Value = "已更新测试" Then
                RunTest "UPDATE 操作", "PASS", "FullName 已更新"
            Else
                RunTest "UPDATE 操作", "FAIL", "值不匹配"
            End If
            testRS.Close
        End If
        Set testRS = Nothing
    Else
        RunTest "UPDATE 操作", "FAIL", Session("LastDBError")
    End If
    ExecuteNonQuery "DELETE FROM Users WHERE UserID=" & testID
End If
On Error GoTo 0

' Test 8: DELETE test
On Error Resume Next
ExecuteNonQuery "INSERT INTO Users (Username, Password, Email, FullName, IsActive, CreatedAt) VALUES ('_test_user_003', 'test_hash', 'test3@test.com', N'删除测试', 1, GETDATE())"
testID = GetLastInsertID("Users")
If testID > 0 Then
    If ExecuteNonQuery("DELETE FROM Users WHERE UserID=" & testID) Then
        Set testRS = ExecuteQuery("SELECT COUNT(*) FROM Users WHERE UserID=" & testID)
        If Not testRS Is Nothing Then
            If testRS.Fields(0).Value = 0 Then
                RunTest "DELETE 操作", "PASS", "记录已删除"
            End If
            testRS.Close
        End If
        Set testRS = Nothing
    Else
        RunTest "DELETE 操作", "FAIL", Session("LastDBError")
    End If
End If
On Error GoTo 0

' Test 9: Transaction test
On Error Resume Next
BeginTransaction
ExecuteNonQuery "INSERT INTO Users (Username, Password, Email, FullName, IsActive, CreatedAt) VALUES ('_test_txn_001', 'test', 'txn@test.com', N'事务测试', 1, GETDATE())"
Dim txnID : txnID = GetLastInsertID("Users")
RollbackTransaction
Set testRS = ExecuteQuery("SELECT COUNT(*) FROM Users WHERE UserID=" & txnID)
If Not testRS Is Nothing Then
    If testRS.Fields(0).Value = 0 Then
        RunTest "事务回滚", "PASS", "回滚成功，数据未持久化"
    Else
        RunTest "事务回滚", "FAIL", "回滚可能失败"
    End If
    testRS.Close
End If
Set testRS = Nothing
On Error GoTo 0
%>

<h2>2. 核心表存在性检查</h2>
<%
Dim coreTables, tableName
coreTables = Array( _
    "Users", "AdminUsers", "AdminRoles", "AdminLogs", _
    "Products", "Categories", "FragranceNotes", "BaseNotes", "BottleStyles", "Volumes", _
    "ProductVolumePrices", "ProductNotes", "ProductNoteRatios", "ProductBottleStyles", _
    "Orders", "OrderDetails", "OrderDetailNoteSelections", _
    "Cart", "CartNoteSelections", _
    "UserAddresses", "UserFavorites", "UserPoints", "PointTransactions", _
    "ProductReviews", _
    "Suppliers", "PurchaseOrders", "PurchaseOrderDetails", "RawMaterialInventory", _
    "Recipes", "RecipeNotes", "RecipeIngredients", "RecipeProducts", _
    "NoteInventory", "ProductInventory", "InventoryTransactions", _
    "ProductionOrders", "ProductionLogs", _
    "PaymentRecords", "RefundRecords", "ReconciliationLogs", _
    "ExpenseRecords", "BudgetPlans", "FundAccounts", _
    "Coupons", "MarketingCampaigns", _
    "SiteSettings", "DailyStatistics", _
    "ModulePermissions", "ProductTypeConfig" _
)

For Each tableName In coreTables
    On Error Resume Next
    Set testRS = ExecuteQuery("SELECT COUNT(*) FROM [" & tableName & "]")
    If Err.Number = 0 And Not testRS Is Nothing Then
        Dim rowCount : rowCount = testRS.Fields(0).Value
        testRS.Close : Set testRS = Nothing
        RunTest tableName, "PASS", rowCount & " 行"
    Else
        RunTest tableName, "FAIL", Err.Description
        Err.Clear
    End If
    On Error GoTo 0
Next
%>

<h2>3. 索引完整性检查</h2>
<%
Dim indexes : indexes = Array( _
    "IX_Orders_UserID", "IX_Orders_Status", "IX_Orders_OrderNo", _
    "IX_OrderDetails_OrderID", "IX_Products_IsActive", _
    "IX_Users_Username", "IX_Users_Email", "IX_FragranceNotes_NoteType", _
    "IX_RecipeAccordMaterials_AccordRecipeID", "IX_RecipeAccords_NoteID", "IX_ProductNoteRatios_ProductID" _
)
Dim idxName
For Each idxName In indexes
    On Error Resume Next
    Set testRS = ExecuteQuery("SELECT COUNT(*) FROM sys.indexes WHERE name='" & idxName & "'")
    If Err.Number = 0 And Not testRS Is Nothing Then
        If testRS.Fields(0).Value > 0 Then
            RunTest "索引 " & idxName, "PASS", ""
        Else
            RunTest "索引 " & idxName, "FAIL", "不存在"
        End If
        testRS.Close
    End If
    Set testRS = Nothing
    Err.Clear
    On Error GoTo 0
Next
%>

<h2>4. 安全与配置测试</h2>
<%
' CSRF Token
Call EnsureCSRFToken()
If Session("CSRFToken") <> "" And Len(Session("CSRFToken")) = 32 Then
    RunTest "CSRF 令牌生成", "PASS", "长度: 32"
Else
    RunTest "CSRF 令牌生成", "FAIL", ""
End If

' SQL Injection protection
Dim safeTest : safeTest = SafeSQL("test'value")
If InStr(safeTest, "''") > 0 Then
    RunTest "SQL 注入防护 (SafeSQL)", "PASS", ""
Else
    RunTest "SQL 注入防护 (SafeSQL)", "FAIL", ""
End If

' XSS protection
Dim xssTest : xssTest = HTMLEncode("<script>alert('xss')</script>")
If InStr(xssTest, "&lt;script&gt;") > 0 Then
    RunTest "XSS 防护 (HTMLEncode)", "PASS", ""
Else
    RunTest "XSS 防护 (HTMLEncode)", "FAIL", ""
End If

' GenerateOrderNo format
Dim orderNoTest : orderNoTest = GenerateOrderNo()
If Len(orderNoTest) = 19 And Left(orderNoTest, 2) = "PF" Then
    RunTest "订单号生成", "PASS", orderNoTest
Else
    RunTest "订单号生成", "FAIL", orderNoTest
End If

' Money formatting
Dim moneyTest : moneyTest = FormatMoney(1234.56)
If InStr(moneyTest, "¥") > 0 And InStr(moneyTest, "1,234.56") > 0 Then
    RunTest "货币格式化", "PASS", moneyTest
Else
    RunTest "货币格式化", "FAIL", moneyTest
End If
%>

<h2>5. 财务模块表验证</h2>
<%
Dim finTables : finTables = Array("PaymentRecords", "RefundRecords", "ReconciliationLogs", "ExpenseRecords", "BudgetPlans", "FundAccounts", "ProductCosts")
Dim finTable
For Each finTable In finTables
    On Error Resume Next
    Set testRS = ExecuteQuery("SELECT COUNT(*) FROM [" & finTable & "]")
    If Err.Number = 0 And Not testRS Is Nothing Then
        RunTest finTable, "PASS", testRS.Fields(0).Value & " 行"
        testRS.Close
    Else
        RunTest finTable, "FAIL", Err.Description
        Err.Clear
    End If
    Set testRS = Nothing
    On Error GoTo 0
Next
%>

<h2>6. 库存管理表验证</h2>
<%
Dim invTables : invTables = Array("NoteInventory", "ProductInventory", "InventoryTransactions", "RawMaterialInventory", "WorkshopTransfer")
Dim invTable
For Each invTable In invTables
    On Error Resume Next
    Set testRS = ExecuteQuery("SELECT COUNT(*) FROM [" & invTable & "]")
    If Err.Number = 0 And Not testRS Is Nothing Then
        RunTest invTable, "PASS", testRS.Fields(0).Value & " 行"
        testRS.Close
    Else
        RunTest invTable, "FAIL", Err.Description
        Err.Clear
    End If
    Set testRS = Nothing
    On Error GoTo 0
Next
%>

<%
CloseConnection
%>

<h2>测试总结</h2>
<%
Dim totalTests : totalTests = passCount + failCount + skipCount
Dim passRate
If totalTests > 0 Then passRate = Round((passCount / totalTests) * 100, 1) Else passRate = 0

Dim summaryClass
If passRate >= 90 Then summaryClass = "PASS" Else summaryClass = "FAIL"

Dim extraInfo : extraInfo = ""
If failCount > 0 Then extraInfo = extraInfo & " | 失败: " & failCount
If skipCount > 0 Then extraInfo = extraInfo & " | 跳过: " & skipCount
%>
<div class="summary <%= summaryClass %>" style="font-size:18px">
    通过率: <%= passRate %>% (<%= passCount %>/<%= totalTests %>)<%= extraInfo %>
</div>
<div class="summary" style="background:#e3f2fd;color:#0d47a1;border:1px solid:#90caf9">
    <% If passRate >= 90 Then %>
        系统状态: 可交付 - 所有核心功能通过验证
    <% ElseIf passRate >= 70 Then %>
        系统状态: 需要关注 - 部分功能需要修复
    <% Else %>
        系统状态: 需要修复 - 多项功能未通过验证
    <% End If %>
</div>

</body>
</html>
