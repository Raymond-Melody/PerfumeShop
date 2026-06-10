<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection()

' ========== SafeNum函数：安全处理数值空值 ==========
Function SafeNum(val)
    If IsNull(val) Then
        SafeNum = 0
    ElseIf val = "" Then
        SafeNum = 0
    ElseIf Not IsNumeric(val) Then
        SafeNum = 0
    Else
        SafeNum = CDbl(val)
    End If
End Function

' ========== SafeDiv函数：安全除法，防止除零 ==========
Function SafeDiv(numerator, denominator)
    If SafeNum(denominator) = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = SafeNum(numerator) / SafeNum(denominator)
    End If
End Function

' 权限检查
Dim canModify
canModify = False
If Session("AdminRoleCode") = "SUPER_ADMIN" Then
    canModify = True
ElseIf Session("AdminRoleCode") = "PURCHASE_MANAGER" Then
    canModify = True
End If

' V8：自动创建 SupplierPrices 表（如果不存在）
On Error Resume Next
conn.Execute "SELECT TOP 1 1 FROM SupplierPrices"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE SupplierPrices (PriceID INT IDENTITY(1,1) PRIMARY KEY, SupplierID INT, ItemName NVARCHAR(200), ItemCode NVARCHAR(100), UnitPrice DECIMAL(19,4) DEFAULT 0, MinOrderQty INT DEFAULT 1, EffectiveDate DATE, ExpiryDate DATE, IsActive BIT DEFAULT 1, CreatedAt DATETIME DEFAULT GETDATE(), PriceType NVARCHAR(30) DEFAULT 'RawMaterial', ValidFrom DATE, ValidTo DATE, DiscountType NVARCHAR(20) DEFAULT 'None', DiscountRule NVARCHAR(500), ApprovalStatus NVARCHAR(20) DEFAULT 'Approved', ApprovedBy NVARCHAR(100), ApprovedAt DATETIME)"
    If Err.Number <> 0 Then Err.Clear
End If
Err.Clear
' V8：自动添加 PriceType 字段
conn.Execute "SELECT PriceType FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD PriceType NVARCHAR(30) DEFAULT 'RawMaterial'"
conn.Execute "SELECT ValidFrom FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD ValidFrom DATE"
conn.Execute "SELECT ValidTo FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD ValidTo DATE"
' V11: 折扣与审批字段
conn.Execute "SELECT DiscountType FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD DiscountType NVARCHAR(20) DEFAULT 'None'"
conn.Execute "SELECT DiscountRule FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD DiscountRule NVARCHAR(500)"
conn.Execute "SELECT ApprovalStatus FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD ApprovalStatus NVARCHAR(20) DEFAULT 'Approved'"
conn.Execute "SELECT ApprovedBy FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD ApprovedBy NVARCHAR(100)"
conn.Execute "SELECT ApprovedAt FROM SupplierPrices WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE SupplierPrices ADD ApprovedAt DATETIME"
' 价格变更日志表
conn.Execute "SELECT TOP 1 1 FROM PriceChangeLog"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PriceChangeLog (LogID INT IDENTITY(1,1) PRIMARY KEY, PriceID INT, FieldChanged NVARCHAR(50), OldValue NVARCHAR(200), NewValue NVARCHAR(200), ChangedBy NVARCHAR(100), ChangedAt DATETIME DEFAULT GETDATE())"
    If Err.Number <> 0 Then Err.Clear
End If
Err.Clear
On Error GoTo 0

' ==================== POST 处理：CRUD ====================
Dim action
action = Request.Form("action")

If action = "add" And canModify Then
    Dim addSupplierId, itemName, itemCode, unitPrice, minOrderQty, effectiveDate, expiryDate, priceType
    addSupplierId = Request.Form("supplierId")
    itemName = SafeSQL(Request.Form("itemName"))
    itemCode = SafeSQL(Request.Form("itemCode"))
    unitPrice = SafeSQL(Request.Form("unitPrice"))
    minOrderQty = SafeSQL(Request.Form("minOrderQty"))
    effectiveDate = SafeSQL(Request.Form("effectiveDate"))
    expiryDate = SafeSQL(Request.Form("expiryDate"))
    priceType = SafeSQL(Request.Form("priceType"))
    If priceType = "" Then priceType = "RawMaterial"
    
    If IsNumeric(addSupplierId) Then
        If itemName <> "" Then
            Dim discountTypeAdd, discountRuleAdd
            discountTypeAdd = SafeSQL(Request.Form("discountType"))
            discountRuleAdd = SafeSQL(Request.Form("discountRule"))
            If discountTypeAdd = "" Then discountTypeAdd = "None"
            ' 验证日期
            Dim errMsg : errMsg = ""
            If effectiveDate <> "" And Not IsDate(effectiveDate) Then
                errMsg = "生效日期格式无效"
            End If
            If expiryDate <> "" And Not IsDate(expiryDate) Then
                errMsg = "过期日期格式无效"
            End If
            If errMsg <> "" Then
                Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode(errMsg)
                Response.End
            End If
            Dim insertSql
            insertSql = "INSERT INTO SupplierPrices (SupplierID, ItemName, ItemCode, UnitPrice, MinOrderQty, EffectiveDate, ExpiryDate, IsActive, PriceType, DiscountType, DiscountRule, ApprovalStatus, CreatedAt) VALUES (" & _
                CInt(addSupplierId) & ", " & _
                "'" & itemName & "', " & _
                "'" & itemCode & "', " & _
                SafeNum(unitPrice) & ", " & _
                SafeNum(minOrderQty) & ", " & _
                "'" & effectiveDate & "', " & _
                "'" & expiryDate & "', " & _
                "1, '" & priceType & "', '" & discountTypeAdd & "', '" & discountRuleAdd & "', 'Approved', GETDATE())"
            If ExecuteNonQuery(insertSql) Then
                Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("报价添加成功")
            Else
                Dim dbgMsg
                dbgMsg = Session("LastDBError")
                If dbgMsg <> "" Then
                    Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("报价添加失败：" & Mid(dbgMsg, 1, 100))
                Else
                    Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("报价添加失败")
                End If
            End If
        Else
            Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("请填写物料名称")
        End If
    Else
        Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("请选择供应商")
    End If
ElseIf action = "edit" And canModify Then
    Dim editPriceId
    editPriceId = Request.Form("priceId")
    addSupplierId = Request.Form("supplierId")
    itemName = SafeSQL(Request.Form("itemName"))
    itemCode = SafeSQL(Request.Form("itemCode"))
    unitPrice = SafeSQL(Request.Form("unitPrice"))
    minOrderQty = SafeSQL(Request.Form("minOrderQty"))
    effectiveDate = SafeSQL(Request.Form("effectiveDate"))
    expiryDate = SafeSQL(Request.Form("expiryDate"))
    priceType = SafeSQL(Request.Form("priceType"))
    If priceType = "" Then priceType = "RawMaterial"
    
    If IsNumeric(editPriceId) Then
        If IsNumeric(addSupplierId) Then
            If itemName <> "" Then
                ' V11: 获取旧值用于变更日志
                Dim oldPrice, oldDiscount
                oldPrice = 0
                On Error Resume Next
                Dim rsOld : Set rsOld = conn.Execute("SELECT UnitPrice, DiscountType FROM SupplierPrices WHERE PriceID=" & CInt(editPriceId))
                If Not rsOld Is Nothing Then
                    If Not rsOld.EOF Then
                        oldPrice = SafeNum(rsOld("UnitPrice"))
                        oldDiscount = rsOld("DiscountType") & ""
                    End If
                    rsOld.Close : Set rsOld = Nothing
                End If
                Err.Clear : On Error GoTo 0
                
                Dim discountTypeEdit, discountRuleEdit
                discountTypeEdit = SafeSQL(Request.Form("discountType"))
                discountRuleEdit = SafeSQL(Request.Form("discountRule"))
                If discountTypeEdit = "" Then discountTypeEdit = "None"
                
                Dim updateSql
                updateSql = "UPDATE SupplierPrices SET " & _
                    "SupplierID = " & CInt(addSupplierId) & ", " & _
                    "ItemName = '" & itemName & "', " & _
                    "ItemCode = '" & itemCode & "', " & _
                    "UnitPrice = " & SafeNum(unitPrice) & ", " & _
                    "MinOrderQty = " & SafeNum(minOrderQty) & ", " & _
                    "EffectiveDate = '" & effectiveDate & "', " & _
                    "ExpiryDate = '" & expiryDate & "', " & _
                    "PriceType = '" & priceType & "', " & _
                    "DiscountType = '" & discountTypeEdit & "', " & _
                    "DiscountRule = '" & discountRuleEdit & "', " & _
                    "ApprovalStatus = 'Approved' " & _
                    "WHERE PriceID = " & CInt(editPriceId)
                If ExecuteNonQuery(updateSql) Then
                    ' V11: 记录价格变更日志
                    If Abs(SafeNum(unitPrice) - oldPrice) > 0.0001 Then
                        conn.Execute "INSERT INTO PriceChangeLog (PriceID, FieldChanged, OldValue, NewValue, ChangedBy) VALUES (" & CInt(editPriceId) & ", 'UnitPrice', '" & FormatNumber(oldPrice, 4) & "', '" & FormatNumber(SafeNum(unitPrice), 4) & "', '" & SafeSQL(Session("AdminUsername")) & "')"
                    End If
                    If discountTypeEdit <> (oldDiscount & "") Then
                        conn.Execute "INSERT INTO PriceChangeLog (PriceID, FieldChanged, OldValue, NewValue, ChangedBy) VALUES (" & CInt(editPriceId) & ", 'DiscountType', '" & SafeSQL(oldDiscount) & "', '" & discountTypeEdit & "', '" & SafeSQL(Session("AdminUsername")) & "')"
                    End If
                    Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("报价更新成功")
                Else
                    Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("报价更新失败")
                End If
            Else
                Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("请填写物料名称")
            End If
        Else
            Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("请选择供应商")
        End If
    Else
        Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("报价ID无效")
    End If
ElseIf action = "toggle" And canModify Then
    Dim togglePriceId
    togglePriceId = Request.Form("priceId")
    If IsNumeric(togglePriceId) Then
        Dim toggleSql
        toggleSql = "UPDATE SupplierPrices SET IsActive = IIF(IsActive <> 0, 0, 1) WHERE PriceID = " & CInt(togglePriceId)
        If ExecuteNonQuery(toggleSql) Then
            Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("状态切换成功")
        Else
            Response.Redirect "price_management.asp?tab=manage&msg=" & Server.URLEncode("状态切换失败")
        End If
    End If
ElseIf action = "approve" And canModify Then
    ' V11: 审批通过
    Dim approvePriceId : approvePriceId = Request.Form("priceId")
    If IsNumeric(approvePriceId) Then
        conn.Execute "UPDATE SupplierPrices SET ApprovalStatus='Approved', ApprovedBy='" & SafeSQL(Session("AdminUsername")) & "', ApprovedAt=GETDATE() WHERE PriceID=" & CInt(approvePriceId)
        ' 记录变更日志
        conn.Execute "INSERT INTO PriceChangeLog (PriceID, FieldChanged, OldValue, NewValue, ChangedBy) VALUES (" & CInt(approvePriceId) & ", 'ApprovalStatus', 'Pending', 'Approved', '" & SafeSQL(Session("AdminUsername")) & "')"
        Response.Redirect "price_management.asp?tab=approval&msg=" & Server.URLEncode("报价已审批通过")
    End If
ElseIf action = "reject" And canModify Then
    ' V11: 审批拒绝
    Dim rejectPriceId : rejectPriceId = Request.Form("priceId")
    If IsNumeric(rejectPriceId) Then
        conn.Execute "UPDATE SupplierPrices SET ApprovalStatus='Rejected', ApprovedBy='" & SafeSQL(Session("AdminUsername")) & "', ApprovedAt=GETDATE() WHERE PriceID=" & CInt(rejectPriceId)
        conn.Execute "INSERT INTO PriceChangeLog (PriceID, FieldChanged, OldValue, NewValue, ChangedBy) VALUES (" & CInt(rejectPriceId) & ", 'ApprovalStatus', 'Pending', 'Rejected', '" & SafeSQL(Session("AdminUsername")) & "')"
        Response.Redirect "price_management.asp?tab=approval&msg=" & Server.URLEncode("报价已拒绝")
    End If
End If

' ==================== 获取Tab参数 ====================
Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then
    currentTab = "manage"
End If

' ==================== 获取筛选参数 ====================
Dim filterSupplier, filterItem, filterStatus, filterPriceType
filterSupplier = Request.QueryString("supplier")
filterItem = Request.QueryString("item")
filterStatus = Request.QueryString("status")
filterPriceType = Request.QueryString("priceType")

' ==================== 统计数据 ====================
Dim totalPrices, activePrices, expiredPrices
totalPrices = GetScalar("SELECT COUNT(*) FROM SupplierPrices")
activePrices = GetScalar("SELECT COUNT(*) FROM SupplierPrices WHERE IsActive <> 0 AND (ExpiryDate >= CAST(GETDATE() AS DATE) OR ExpiryDate IS NULL)")
expiredPrices = GetScalar("SELECT COUNT(*) FROM SupplierPrices WHERE ExpiryDate < CAST(GETDATE() AS DATE)")

' V8 有效期管理：即将过期统计
Dim expiringSoon
expiringSoon = GetScalar("SELECT COUNT(*) FROM SupplierPrices WHERE IsActive <> 0 AND ExpiryDate >= CAST(GETDATE() AS DATE) AND ExpiryDate <= DATEADD(DAY, 30, CAST(GETDATE() AS DATE))")
' V11: 待审批统计
Dim pendingApproval
pendingApproval = GetScalar("SELECT COUNT(*) FROM SupplierPrices WHERE ApprovalStatus = 'Pending'")

' ==================== 获取供应商列表（用于下拉选择）====================
Dim rsSuppliers, supplierList(), supplierCount, i
supplierCount = 0
Set rsSuppliers = ExecuteQuery("SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive <> 0 ORDER BY SupplierName")
If Not rsSuppliers Is Nothing Then
    If Not rsSuppliers.EOF Then
        rsSuppliers.MoveLast
        supplierCount = rsSuppliers.RecordCount
        rsSuppliers.MoveFirst
        ReDim supplierList(supplierCount - 1, 1)
        i = 0
        Do While Not rsSuppliers.EOF
            supplierList(i, 0) = rsSuppliers("SupplierID")
            supplierList(i, 1) = rsSuppliers("SupplierName")
            i = i + 1
            rsSuppliers.MoveNext
        Loop
    End If
    rsSuppliers.Close
    Set rsSuppliers = Nothing
End If

' ==================== 构建报价查询条件 ====================
Dim priceWhereClause, priceHasCondition
priceWhereClause = ""
priceHasCondition = False

If filterSupplier <> "" Then
    If IsNumeric(filterSupplier) Then
        priceWhereClause = priceWhereClause & "p.SupplierID = " & CInt(filterSupplier)
        priceHasCondition = True
    End If
End If

If filterItem <> "" Then
    If priceHasCondition Then
        priceWhereClause = priceWhereClause & " AND "
    End If
    priceWhereClause = priceWhereClause & "(p.ItemName LIKE '%" & SafeSQL(filterItem) & "%' OR p.ItemCode LIKE '%" & SafeSQL(filterItem) & "%')"
    priceHasCondition = True
End If

If filterStatus = "active" Then
    If priceHasCondition Then
        priceWhereClause = priceWhereClause & " AND "
    End If
    priceWhereClause = priceWhereClause & "p.IsActive <> 0"
    priceHasCondition = True
ElseIf filterStatus = "inactive" Then
    If priceHasCondition Then
        priceWhereClause = priceWhereClause & " AND "
    End If
    priceWhereClause = priceWhereClause & "p.IsActive = 0"
    priceHasCondition = True
ElseIf filterStatus = "expiring" Then
    If priceHasCondition Then
        priceWhereClause = priceWhereClause & " AND "
    End If
    priceWhereClause = priceWhereClause & "p.IsActive <> 0 AND p.ExpiryDate >= CAST(GETDATE() AS DATE) AND p.ExpiryDate <= DATEADD(DAY, 30, CAST(GETDATE() AS DATE))"
    priceHasCondition = True
ElseIf filterStatus = "expired" Then
    If priceHasCondition Then
        priceWhereClause = priceWhereClause & " AND "
    End If
    priceWhereClause = priceWhereClause & "p.ExpiryDate < CAST(GETDATE() AS DATE)"
    priceHasCondition = True
End If

If filterPriceType <> "" Then
    If priceHasCondition Then
        priceWhereClause = priceWhereClause & " AND "
    End If
    priceWhereClause = priceWhereClause & "p.PriceType = '" & SafeSQL(filterPriceType) & "'"
    priceHasCondition = True
End If

If priceHasCondition Then
    priceWhereClause = "WHERE " & priceWhereClause
End If

' ==================== 获取报价列表 ====================
Dim rsPrices, priceList(), priceCount
priceCount = 0
Dim priceSql
priceSql = "SELECT p.*, s.SupplierName FROM SupplierPrices p LEFT JOIN Suppliers s ON p.SupplierID = s.SupplierID " & priceWhereClause & " ORDER BY p.CreatedAt DESC"
Set rsPrices = ExecuteQuery(priceSql)
If Not rsPrices Is Nothing Then
    If Not rsPrices.EOF Then
        rsPrices.MoveLast
        priceCount = rsPrices.RecordCount
        rsPrices.MoveFirst
        ReDim priceList(priceCount - 1, 15)
        i = 0
        On Error Resume Next
        Do While Not rsPrices.EOF
            priceList(i, 0) = rsPrices("PriceID")
            priceList(i, 1) = rsPrices("SupplierID")
            priceList(i, 2) = rsPrices("ItemName")
            priceList(i, 3) = rsPrices("ItemCode")
            priceList(i, 4) = rsPrices("UnitPrice")
            priceList(i, 5) = rsPrices("MinOrderQty")
            priceList(i, 6) = rsPrices("EffectiveDate")
            priceList(i, 7) = rsPrices("ExpiryDate")
            priceList(i, 8) = rsPrices("IsActive")
            priceList(i, 9) = rsPrices("CreatedAt")
            priceList(i, 10) = rsPrices("SupplierName")
            Err.Clear
            priceList(i, 11) = rsPrices("PriceType") & ""
            If Err.Number <> 0 Then Err.Clear : priceList(i, 11) = "RawMaterial"
            priceList(i, 12) = rsPrices("DiscountType") & ""
            If Err.Number <> 0 Then Err.Clear : priceList(i, 12) = "None"
            priceList(i, 13) = rsPrices("DiscountRule") & ""
            If Err.Number <> 0 Then Err.Clear : priceList(i, 13) = ""
            priceList(i, 14) = rsPrices("ApprovalStatus") & ""
            If Err.Number <> 0 Then Err.Clear : priceList(i, 14) = "Approved"
            priceList(i, 15) = rsPrices("ApprovedAt")
            If Err.Number <> 0 Then Err.Clear : priceList(i, 15) = Null
            i = i + 1
            rsPrices.MoveNext
        Loop
        On Error GoTo 0
    End If
    rsPrices.Close
    Set rsPrices = Nothing
End If

' ==================== 获取价格对比数据（按物料分组）====================
Dim rsCompare, compareList(), compareCount
compareCount = 0
Set rsCompare = ExecuteQuery("SELECT ItemName, ItemCode, COUNT(*) AS SupplierCount FROM SupplierPrices WHERE IsActive <> 0 GROUP BY ItemName, ItemCode HAVING COUNT(*) > 1 ORDER BY ItemName")
If Not rsCompare Is Nothing Then
    If Not rsCompare.EOF Then
        rsCompare.MoveLast
        compareCount = rsCompare.RecordCount
        rsCompare.MoveFirst
        ReDim compareList(compareCount - 1, 2)
        i = 0
        Do While Not rsCompare.EOF
            compareList(i, 0) = rsCompare("ItemName")
            compareList(i, 1) = rsCompare("ItemCode")
            compareList(i, 2) = rsCompare("SupplierCount")
            i = i + 1
            rsCompare.MoveNext
        Loop
    End If
    rsCompare.Close
    Set rsCompare = Nothing
End If

' ==================== 获取历史价格追踪数据 ====================
Dim historyItem, rsHistory, historyList(), historyCount
historyItem = Request.QueryString("historyItem")
historyCount = 0
If historyItem <> "" Then
    Set rsHistory = ExecuteQuery("SELECT p.*, s.SupplierName FROM SupplierPrices p LEFT JOIN Suppliers s ON p.SupplierID = s.SupplierID WHERE p.ItemName = '" & SafeSQL(historyItem) & "' OR p.ItemCode = '" & SafeSQL(historyItem) & "' ORDER BY p.EffectiveDate DESC")
    If Not rsHistory Is Nothing Then
        If Not rsHistory.EOF Then
            rsHistory.MoveLast
            historyCount = rsHistory.RecordCount
            rsHistory.MoveFirst
            ReDim historyList(historyCount - 1, 6)
            i = 0
            Do While Not rsHistory.EOF
                historyList(i, 0) = rsHistory("PriceID")
                historyList(i, 1) = rsHistory("SupplierName")
                historyList(i, 2) = rsHistory("UnitPrice")
                historyList(i, 3) = rsHistory("EffectiveDate")
                historyList(i, 4) = rsHistory("ExpiryDate")
                historyList(i, 5) = rsHistory("IsActive")
                historyList(i, 6) = rsHistory("CreatedAt")
                i = i + 1
                rsHistory.MoveNext
            Loop
        End If
        rsHistory.Close
        Set rsHistory = Nothing
    End If
End If

' ==================== V11: 获取待审批列表 ====================
Dim rsApproval, approvalList(), approvalCount
approvalCount = 0
If currentTab = "approval" Then
    Dim approvalWhere : approvalWhere = ""
    If filterStatus = "approved" Then
        approvalWhere = " AND p.ApprovalStatus = 'Approved'"
    ElseIf filterStatus = "rejected" Then
        approvalWhere = " AND p.ApprovalStatus = 'Rejected'"
    Else
        approvalWhere = " AND p.ApprovalStatus = 'Pending'"
    End If
    If filterItem <> "" Then
        approvalWhere = approvalWhere & " AND (p.ItemName LIKE '%" & SafeSQL(filterItem) & "%' OR p.ItemCode LIKE '%" & SafeSQL(filterItem) & "%')"
    End If
    
    Set rsApproval = ExecuteQuery("SELECT p.*, s.SupplierName FROM SupplierPrices p LEFT JOIN Suppliers s ON p.SupplierID = s.SupplierID WHERE 1=1 " & approvalWhere & " ORDER BY p.CreatedAt DESC")
    If Not rsApproval Is Nothing Then
        If Not rsApproval.EOF Then
            rsApproval.MoveLast
            approvalCount = rsApproval.RecordCount
            rsApproval.MoveFirst
            ReDim approvalList(approvalCount - 1, 10)
            i = 0
            Do While Not rsApproval.EOF
                approvalList(i, 0) = rsApproval("PriceID")
                approvalList(i, 1) = rsApproval("SupplierID")
                approvalList(i, 2) = rsApproval("ItemName")
                approvalList(i, 3) = rsApproval("ItemCode")
                approvalList(i, 4) = rsApproval("UnitPrice")
                approvalList(i, 5) = rsApproval("MinOrderQty")
                approvalList(i, 6) = rsApproval("PriceType") & ""
                approvalList(i, 7) = rsApproval("SupplierName")
                approvalList(i, 8) = rsApproval("ApprovalStatus") & ""
                approvalList(i, 9) = rsApproval("ApprovedAt")
                approvalList(i, 10) = rsApproval("CreatedAt")
                i = i + 1
                rsApproval.MoveNext
            Loop
        End If
        rsApproval.Close
        Set rsApproval = Nothing
    End If
End If

' ==================== V11: 获取价格变更日志 ====================
Dim logPriceId, rsPriceLog, priceLogList(), priceLogCount
logPriceId = Request.QueryString("logPriceId")
priceLogCount = 0
If logPriceId <> "" And IsNumeric(logPriceId) Then
    Set rsPriceLog = ExecuteQuery("SELECT * FROM PriceChangeLog WHERE PriceID = " & CInt(logPriceId) & " ORDER BY ChangedAt DESC")
    If Not rsPriceLog Is Nothing Then
        If Not rsPriceLog.EOF Then
            rsPriceLog.MoveLast
            priceLogCount = rsPriceLog.RecordCount
            rsPriceLog.MoveFirst
            ReDim priceLogList(priceLogCount - 1, 4)
            i = 0
            Do While Not rsPriceLog.EOF
                priceLogList(i, 0) = rsPriceLog("FieldChanged")
                priceLogList(i, 1) = rsPriceLog("OldValue") & ""
                priceLogList(i, 2) = rsPriceLog("NewValue") & ""
                priceLogList(i, 3) = rsPriceLog("ChangedBy") & ""
                priceLogList(i, 4) = rsPriceLog("ChangedAt")
                i = i + 1
                rsPriceLog.MoveNext
            Loop
        End If
        rsPriceLog.Close
        Set rsPriceLog = Nothing
    End If
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>价格管理 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* Tab导航 */
        .tab-nav {
            display: flex;
            gap: 10px;
            margin-bottom: 25px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            padding-bottom: 15px;
        }
        .tab-btn {
            padding: 12px 25px;
            background: #2d2d44;
            border: none;
            border-radius: 8px;
            color: #888;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s ease;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        .tab-btn:hover {
            background: #3a3a5c;
            color: #fff;
        }
        .tab-btn.active {
            background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%);
            color: white;
        }
        
        /* 统计卡片 */
        .stats-grid { 
            display: grid; 
            grid-template-columns: repeat(3, 1fr); 
            gap: 20px; 
            margin-bottom: 25px; 
        }
        .stat-card { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            padding: 20px; 
            border-radius: 12px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            text-align: center;
            transition: transform 0.2s ease;
        }
        .stat-card:hover {
            transform: translateY(-3px);
        }
        .stat-card.total { border-top: 4px solid #FF9800; }
        .stat-card.active { border-top: 4px solid #4CAF50; }
        .stat-card.expired { border-top: 4px solid #f44336; }
        .stat-value { font-size: 28px; font-weight: bold; color: #fff; }
        .stat-label { color: #888; margin-top: 8px; font-size: 13px; }
        
        /* 筛选栏 */
        .filter-bar { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            padding: 20px; 
            border-radius: 12px; 
            margin-bottom: 20px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            display: flex; 
            gap: 15px; 
            flex-wrap: wrap; 
            align-items: center; 
        }
        .filter-bar select, .filter-bar input { 
            padding: 10px 15px; 
            border: 1px solid rgba(255,255,255,0.1); 
            border-radius: 8px; 
            font-size: 14px; 
            background: #1e1e32;
            color: #e0e0e0;
        }
        .filter-bar input[type="text"] { width: 200px; }
        .filter-bar button { 
            padding: 10px 20px; 
            border: none; 
            border-radius: 8px; 
            cursor: pointer; 
            font-size: 14px;
            transition: all 0.2s ease;
        }
        .btn-search { background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%); color: white; }
        .btn-search:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(255,152,0,0.3); }
        .btn-reset { background: #3a3a5c; color: #e0e0e0; }
        .btn-reset:hover { background: #4a4a6c; }
        .btn-add { 
            background: linear-gradient(135deg, #4CAF50 0%, #388E3C 100%); 
            color: white; 
            margin-left: auto;
        }
        .btn-add:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(76,175,80,0.3); }
        .btn-add.disabled {
            background: #3a3a5c;
            cursor: not-allowed;
        }
        
        /* 表格样式 */
        .data-table { 
            width: 100%; 
            border-collapse: collapse; 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px; 
            overflow: hidden; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
        }
        .data-table th { 
            background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%);
            color: white; 
            padding: 15px; 
            text-align: left;
            font-weight: 600;
        }
        .data-table td { 
            padding: 15px; 
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #e0e0e0;
        }
        .data-table tr:hover { background: rgba(255,255,255,0.02); }
        
        .status-badge { 
            display: inline-block; 
            padding: 5px 12px; 
            border-radius: 12px; 
            font-size: 12px; 
            font-weight: 500; 
        }
        .status-active { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-inactive { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        .status-expired { background: rgba(244,67,54,0.2); color: #f44336; }
        
        .price-value {
            font-weight: 600;
            color: #FF9800;
        }
        
        .btn-action { 
            padding: 6px 12px; 
            border: none; 
            border-radius: 6px; 
            cursor: pointer; 
            font-size: 12px; 
            margin-right: 5px;
            transition: all 0.2s ease;
        }
        .btn-action:hover { transform: translateY(-2px); }
        .btn-edit { background: linear-gradient(135deg, #2196F3 0%, #1976D2 100%); color: white; }
        .btn-toggle { background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%); color: white; }
        .btn-toggle.disabled { background: #3a3a5c; cursor: not-allowed; }
        .btn-history { background: linear-gradient(135deg, #9C27B0 0%, #7B1FA2 100%); color: white; }
        
        /* 模态框 */
        .modal { 
            display: none; 
            position: fixed; 
            top: 0; 
            left: 0; 
            width: 100%; 
            height: 100%; 
            background: rgba(0,0,0,0.7); 
            z-index: 1000;
            backdrop-filter: blur(5px);
        }
        .modal-content { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            width: 90%; 
            max-width: 500px; 
            margin: 50px auto; 
            padding: 30px; 
            border-radius: 12px; 
            max-height: 80vh; 
            overflow-y: auto;
            border: 1px solid rgba(255,255,255,0.1);
            box-shadow: 0 10px 40px rgba(0,0,0,0.5);
        }
        .modal-header { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            margin-bottom: 25px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            padding-bottom: 15px;
        }
        .modal-header h3 { 
            margin: 0; 
            color: #fff;
            font-size: 20px;
        }
        .close-btn { 
            background: none; 
            border: none; 
            font-size: 28px; 
            cursor: pointer; 
            color: #888;
            transition: color 0.2s ease;
        }
        .close-btn:hover { color: #fff; }
        .form-group { margin-bottom: 20px; }
        .form-group label { 
            display: block; 
            margin-bottom: 8px; 
            font-weight: 500; 
            color: #b0b0b0; 
        }
        .form-group input, .form-group select { 
            width: 100%; 
            padding: 12px; 
            border: 1px solid rgba(255,255,255,0.1); 
            border-radius: 8px; 
            font-size: 14px; 
            box-sizing: border-box;
            background: #1e1e32;
            color: #e0e0e0;
        }
        .form-group input:focus, .form-group select:focus {
            outline: none;
            border-color: #FF9800;
        }
        .required { color: #f44336; }
        .form-actions { 
            text-align: right; 
            margin-top: 25px;
            padding-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.1);
        }
        .form-actions button { 
            padding: 12px 25px; 
            border: none; 
            border-radius: 8px; 
            cursor: pointer; 
            font-size: 14px; 
            margin-left: 10px;
            transition: all 0.2s ease;
        }
        .btn-cancel { background: #3a3a5c; color: #e0e0e0; }
        .btn-cancel:hover { background: #4a4a6c; }
        .btn-save { 
            background: linear-gradient(135deg, #4CAF50 0%, #388E3C 100%); 
            color: white; 
        }
        .btn-save:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(76,175,80,0.3); }
        
        /* 价格对比卡片 */
        .compare-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 20px;
        }
        .compare-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
        }
        .compare-card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .compare-item-name {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
        }
        .compare-item-code {
            font-size: 12px;
            color: #888;
            margin-top: 3px;
        }
        .compare-count {
            background: rgba(255,152,0,0.2);
            color: #FF9800;
            padding: 5px 12px;
            border-radius: 12px;
            font-size: 12px;
        }
        .compare-table {
            width: 100%;
            font-size: 13px;
        }
        .compare-table th {
            text-align: left;
            padding: 8px 5px;
            color: #888;
            font-weight: 500;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }
        .compare-table td {
            padding: 10px 5px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }
        .compare-price {
            font-weight: 600;
            color: #FF9800;
        }
        .compare-price.best {
            color: #4CAF50;
        }
        
        /* 历史价格 */
        .history-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
        }
        .history-title {
            font-size: 16px;
            color: #fff;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .history-title i {
            color: #FF9800;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .compare-grid { grid-template-columns: 1fr; }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .btn-add { margin-left: 0; margin-top: 10px; }
            .tab-nav { flex-wrap: wrap; }
        }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-tags"></i> 价格管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>价格管理</span>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success" style="background: rgba(76,175,80,0.2); color: #4CAF50; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <!-- Tab导航 -->
        <div class="tab-nav">
            <a href="?tab=manage" class="tab-btn <%= IIf(currentTab = "manage", "active", "") %>">
                <i class="fas fa-list"></i> 报价管理
            </a>
            <a href="?tab=compare" class="tab-btn <%= IIf(currentTab = "compare", "active", "") %>">
                <i class="fas fa-balance-scale"></i> 价格对比
            </a>
            <a href="?tab=approval" class="tab-btn <%= IIf(currentTab = "approval", "active", "") %>">
                <i class="fas fa-check-circle"></i> 审批管理
                <% If pendingApproval > 0 Then %>
                <span style="background:#f44336;color:#fff;padding:1px 7px;border-radius:10px;font-size:11px;margin-left:4px;"><%= pendingApproval %></span>
                <% End If %>
            </a>
        </div>
        
        <% If currentTab = "manage" Then %>
        <!-- 报价管理 Tab -->
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card total">
                <div class="stat-value"><%= totalPrices %></div>
                <div class="stat-label">总报价数</div>
            </div>
            <div class="stat-card active">
                <div class="stat-value"><%= activePrices %></div>
                <div class="stat-label">有效报价</div>
            </div>
            <div class="stat-card" style="border-top:4px solid #FF9800;">
                <div class="stat-value" <% If expiringSoon > 0 Then %>style="color:#FF9800;"<% End If %>><%= expiringSoon %></div>
                <div class="stat-label">即将过期(30天)</div>
            </div>
            <div class="stat-card expired">
                <div class="stat-value"><%= expiredPrices %></div>
                <div class="stat-label">已过期报价</div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <form method="get" action="price_management.asp" style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center; flex: 1;">
                <input type="hidden" name="tab" value="manage">
                <select name="supplier">
                    <option value="">全部供应商</option>
                    <% For i = 0 To supplierCount - 1 %>
                    <option value="<%= supplierList(i, 0) %>" <% If CStr(filterSupplier) = CStr(supplierList(i, 0)) Then Response.Write "selected" %>><%= HTMLEncode(supplierList(i, 1)) %></option>
                    <% Next %>
                </select>
                <input type="text" name="item" placeholder="搜索物料名称或编码" value="<%= HTMLEncode(filterItem) %>">
                <select name="status">
                    <option value="">全部状态</option>
                    <option value="active" <% If filterStatus = "active" Then Response.Write "selected" %>>有效</option>
                    <option value="expiring" <% If filterStatus = "expiring" Then Response.Write "selected" %>>即将过期(30天)</option>
                    <option value="expired" <% If filterStatus = "expired" Then Response.Write "selected" %>>已过期</option>
                    <option value="inactive" <% If filterStatus = "inactive" Then Response.Write "selected" %>>无效</option>
                </select>
                <select name="priceType">
                    <option value="">全部类型</option>
                    <option value="RawMaterial" <% If filterPriceType = "RawMaterial" Then Response.Write "selected" %>>原料</option>
                    <option value="BaseNote" <% If filterPriceType = "BaseNote" Then Response.Write "selected" %>>基香原料</option>
                    <option value="Packaging" <% If filterPriceType = "Packaging" Then Response.Write "selected" %>>包装物</option>
                    <option value="Bottle" <% If filterPriceType = "Bottle" Then Response.Write "selected" %>>瓶子</option>
                    <option value="Printing" <% If filterPriceType = "Printing" Then Response.Write "selected" %>>印刷品</option>
                    <option value="SprayHead" <% If filterPriceType = "SprayHead" Then Response.Write "selected" %>>喷头</option>
                </select>
                <button type="submit" class="btn-search"><i class="fas fa-search"></i> 搜索</button>
                <a href="price_management.asp?tab=manage" class="btn-reset" style="text-decoration: none; display: inline-block; padding: 10px 20px;"><i class="fas fa-undo"></i> 重置</a>
            </form>
            <% If canModify Then %>
            <button class="btn-add" onclick="openAddModal()"><i class="fas fa-plus"></i> 新增报价</button>
            <% Else %>
            <button class="btn-add disabled" disabled title="无权限"><i class="fas fa-plus"></i> 新增报价</button>
            <% End If %>
        </div>
        
        <!-- 报价列表 -->
        <table class="data-table">
            <thead>
                <tr>
                    <th>ID</th>
                    <th>类型</th>
                    <th>供应商</th>
                    <th>物料名称</th>
                    <th>物料编码</th>
                    <th>单价</th>
                    <th>最小起订量</th>
                    <th>生效日期</th>
                    <th>失效日期</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If priceCount > 0 Then %>
                <% For i = 0 To priceCount - 1 %>
                <tr>
                    <td>#<%= priceList(i, 0) %></td>
                    <td>
                        <%
                            Dim ptLabel : ptLabel = priceList(i, 11) & ""
                            If ptLabel = "Packaging" Then
                                Response.Write "<span style='background:rgba(156,39,176,0.2);color:#CE93D8;padding:3px 10px;border-radius:12px;font-size:12px;'>包装物</span>"
                            ElseIf ptLabel = "Bottle" Then
                                Response.Write "<span style='background:rgba(33,150,243,0.2);color:#64B5F6;padding:3px 10px;border-radius:12px;font-size:12px;'>瓶子</span>"
                            ElseIf ptLabel = "Printing" Then
                                Response.Write "<span style='background:rgba(0,188,212,0.2);color:#00BCD4;padding:3px 10px;border-radius:12px;font-size:12px;'>印刷品</span>"
                            ElseIf ptLabel = "SprayHead" Then
                                Response.Write "<span style='background:rgba(255,87,34,0.2);color:#FF5722;padding:3px 10px;border-radius:12px;font-size:12px;'>喷头</span>"
                            ElseIf ptLabel = "BaseNote" Then
                                Response.Write "<span style='background:rgba(76,175,80,0.2);color:#81C784;padding:3px 10px;border-radius:12px;font-size:12px;'>基香原料</span>"
                            Else
                                Response.Write "<span style='background:rgba(255,152,0,0.2);color:#FFB74D;padding:3px 10px;border-radius:12px;font-size:12px;'>原料</span>"
                            End If
                        %>
                    </td>
                    <td><%= HTMLEncode(priceList(i, 10)) %></td>
                    <td><%= HTMLEncode(priceList(i, 2)) %></td>
                    <td><%= HTMLEncode(priceList(i, 3)) %></td>
                    <td class="price-value">¥<%= FormatNumber(SafeNum(priceList(i, 4)), 2) %></td>
                    <td><%= FormatNumber(SafeNum(priceList(i, 5)), 0) %></td>
                    <td><%
                            Dim effDt : effDt = priceList(i, 6)
                            If Not IsNull(effDt) And IsDate(effDt) Then
                                Response.Write FormatDateTime(effDt, 2)
                            End If
                        %></td>
                    <td>
                        <%
                            Dim expDt : expDt = priceList(i, 7)
                            If Not IsNull(expDt) And IsDate(expDt) Then
                                If CDate(expDt) < Date() Then
                        %>
                            <span style="color: #f44336;"><%= FormatDateTime(expDt, 2) %></span>
                            <% Else %>
                            <%= FormatDateTime(expDt, 2) %>
                            <% End If %>
                        <% Else %>
                        <span style="color: #888;">无期限</span>
                        <% End If %>
                    </td>
                    <td>
                        <%
                        Dim statExpDt : statExpDt = priceList(i, 7)
                        Dim isExp : isExp = False
                        Dim isExpiringSoon : isExpiringSoon = False
                        If priceList(i, 8) = 0 Then %>
                        <span class="status-badge status-inactive"><i class="fas fa-ban"></i> 无效</span>
                        <%
                        Else
                            If Not IsNull(statExpDt) Then
                                If IsDate(statExpDt) Then
                                    If CDate(statExpDt) < Date() Then
                                        isExp = True
                                    ElseIf CDate(statExpDt) <= DateAdd("d", 30, Date()) Then
                                        isExpiringSoon = True
                                    End If
                                End If
                            End If
                            If isExp Then
                        %>
                        <span class="status-badge status-expired"><i class="fas fa-clock"></i> 已过期</span>
                        <% ElseIf isExpiringSoon Then %>
                        <span class="status-badge" style="background:rgba(255,152,0,0.2);color:#FF9800;"><i class="fas fa-exclamation-triangle"></i> 即将过期</span>
                        <% Else %>
                        <span class="status-badge status-active"><i class="fas fa-check-circle"></i> 有效</span>
                        <% End If %>
                        <% End If %>
                    </td>
                    <td>
                        <button class="btn-action btn-edit" onclick="openEditModal('<%= priceList(i, 0) %>', '<%= priceList(i, 1) %>', '<%= Server.HTMLEncode(priceList(i, 2)) %>', '<%= Server.HTMLEncode(priceList(i, 3)) %>', '<%= priceList(i, 4) %>', '<%= priceList(i, 5) %>', '<%= priceList(i, 6) %>', '<%= priceList(i, 7) %>', '<%= priceList(i, 11) & "" %>', '<%= priceList(i, 12) & "" %>', '<%= Server.HTMLEncode(priceList(i, 13) & "") %>')">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <% If canModify Then %>
                        <form method="post" action="price_management.asp?tab=manage" style="display: inline;">
                            <input type="hidden" name="action" value="toggle">
                            <input type="hidden" name="priceId" value="<%= priceList(i, 0) %>">
                            <button type="submit" class="btn-action btn-toggle" onclick="return confirm('确定要切换该报价的状态吗？')">
                                <i class="fas fa-exchange-alt"></i> <%= IIf(priceList(i, 8) <> 0, "无效", "有效") %>
                            </button>
                        </form>
                        <% Else %>
                        <button class="btn-action btn-toggle disabled" disabled title="无权限">
                            <i class="fas fa-exchange-alt"></i> <%= IIf(priceList(i, 8) <> 0, "无效", "有效") %>
                        </button>
                        <% End If %>
                    </td>
                </tr>
                <% Next %>
                <% Else %>
                <tr>
                    <td colspan="11" style="text-align: center; padding: 40px; color: #666;">
                        <i class="fas fa-inbox" style="font-size: 48px; display: block; margin-bottom: 15px;"></i>
                        暂无报价数据
                    </td>
                </tr>
                <% End If %>
            </tbody>
        </table>
        
        <% ElseIf currentTab = "compare" Then %>
        <!-- 价格对比 Tab -->
        <div class="filter-bar">
            <form method="get" action="price_management.asp" style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center; flex: 1;">
                <input type="hidden" name="tab" value="compare">
                <input type="text" name="historyItem" placeholder="输入物料名称或编码查看历史价格" value="<%= HTMLEncode(historyItem) %>">
                <button type="submit" class="btn-history"><i class="fas fa-history"></i> 查看历史价格</button>
                <% If historyItem <> "" Then %>
                <a href="price_management.asp?tab=compare" class="btn-reset" style="text-decoration: none; display: inline-block; padding: 10px 20px;"><i class="fas fa-undo"></i> 清除</a>
                <% End If %>
            </form>
        </div>
        
        <% If historyItem <> "" And historyCount > 0 Then %>
        <!-- 历史价格追踪 -->
        <div class="history-section">
            <div class="history-title">
                <i class="fas fa-chart-line"></i>
                "<%= HTMLEncode(historyItem) %>" 的历史价格追踪
            </div>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>供应商</th>
                        <th>单价</th>
                        <th>生效日期</th>
                        <th>失效日期</th>
                        <th>状态</th>
                        <th>创建时间</th>
                    </tr>
                </thead>
                <tbody>
                    <% For i = 0 To historyCount - 1 %>
                    <tr>
                        <td><%= HTMLEncode(historyList(i, 1)) %></td>
                        <td class="price-value">¥<%= FormatNumber(SafeNum(historyList(i, 2)), 2) %></td>
                        <td><% If IsDate(historyList(i, 3)) Then Response.Write FormatDateTime(historyList(i, 3), 2) End If %></td>
                        <td>
                            <% If Not IsNull(historyList(i, 4)) And IsDate(historyList(i, 4)) Then %>
                            <%= FormatDateTime(historyList(i, 4), 2) %>
                            <% Else %>
                            <span style="color: #888;">无期限</span>
                            <% End If %>
                        </td>
                        <td>
                            <% If historyList(i, 5) = 0 Then %>
                            <span class="status-badge status-inactive">无效</span>
                            <% Else %>
                            <span class="status-badge status-active">有效</span>
                            <% End If %>
                        </td>
                        <td><% If IsDate(historyList(i, 6)) Then Response.Write FormatDateTime(historyList(i, 6), 2) End If %></td>
                    </tr>
                    <% Next %>
                </tbody>
            </table>
        </div>
        <% End If %>
        
        <!-- 价格对比卡片 -->
        <% If compareCount > 0 Then %>
        <div class="compare-grid">
            <% For i = 0 To compareCount - 1 %>
            <div class="compare-card">
                <div class="compare-card-header">
                    <div>
                        <div class="compare-item-name"><%= HTMLEncode(compareList(i, 0)) %></div>
                        <div class="compare-item-code"><%= HTMLEncode(compareList(i, 1)) %></div>
                    </div>
                    <div class="compare-count"><%= compareList(i, 2) %> 家供应商</div>
                </div>
                <table class="compare-table">
                    <thead>
                        <tr>
                            <th>供应商</th>
                            <th>单价</th>
                            <th>起订量</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% 
                        ' 获取该物料的所有报价
                        Dim rsItemPrices, minPrice
                        minPrice = 999999999
                        Set rsItemPrices = ExecuteQuery("SELECT p.UnitPrice, p.MinOrderQty, s.SupplierName FROM SupplierPrices p LEFT JOIN Suppliers s ON p.SupplierID = s.SupplierID WHERE p.ItemName = '" & SafeSQL(compareList(i, 0)) & "' AND p.IsActive <> 0 AND (p.ExpiryDate >= CAST(GETDATE() AS DATE) OR p.ExpiryDate IS NULL) ORDER BY p.UnitPrice ASC")
                        If Not rsItemPrices Is Nothing Then
                            If Not rsItemPrices.EOF Then
                                ' 先找最低价
                                rsItemPrices.MoveFirst
                                Do While Not rsItemPrices.EOF
                                    If SafeNum(rsItemPrices("UnitPrice")) < minPrice Then
                                        minPrice = SafeNum(rsItemPrices("UnitPrice"))
                                    End If
                                    rsItemPrices.MoveNext
                                Loop
                                ' 再显示
                                rsItemPrices.MoveFirst
                                Do While Not rsItemPrices.EOF
                        %>
                        <tr>
                            <td><%= HTMLEncode(CStr(rsItemPrices("SupplierName"))) %></td>
                            <td class="compare-price <%= IIf(SafeNum(rsItemPrices("UnitPrice")) = minPrice, "best", "") %>">
                                ¥<%= FormatNumber(SafeNum(rsItemPrices("UnitPrice")), 2) %>
                                <% If SafeNum(rsItemPrices("UnitPrice")) = minPrice Then %>
                                <i class="fas fa-check-circle" style="color: #4CAF50; margin-left: 5px;"></i>
                                <% End If %>
                            </td>
                            <td><%= FormatNumber(SafeNum(rsItemPrices("MinOrderQty")), 0) %></td>
                        </tr>
                        <% 
                                    rsItemPrices.MoveNext
                                Loop
                            End If
                            rsItemPrices.Close
                            Set rsItemPrices = Nothing
                        End If
                        %>
                    </tbody>
                </table>
                <div style="margin-top: 15px; text-align: right;">
                    <a href="?tab=compare&historyItem=<%= Server.URLEncode(compareList(i, 0)) %>" class="btn-history" style="text-decoration: none; padding: 8px 15px; border-radius: 6px; font-size: 12px;">
                        <i class="fas fa-history"></i> 查看历史价格
                    </a>
                </div>
            </div>
            <% Next %>
        </div>
        <% Else %>
        <div style="text-align: center; padding: 60px; color: #666;">
            <i class="fas fa-balance-scale" style="font-size: 64px; display: block; margin-bottom: 20px;"></i>
            暂无可对比的物料（需要同一物料有多个供应商报价）
        </div>
        <% End If %>
        <% ElseIf currentTab = "approval" Then %>
        <!-- V11: 审批管理 Tab -->
        <div class="stats-grid">
            <div class="stat-card" style="border-top:4px solid #FF9800;">
                <div class="stat-value"><%= pendingApproval %></div>
                <div class="stat-label">待审批</div>
            </div>
        </div>
        
        <div class="filter-bar">
            <form method="get" action="price_management.asp" style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center; flex: 1;">
                <input type="hidden" name="tab" value="approval">
                <input type="text" name="item" placeholder="搜索物料名称或编码" value="<%= HTMLEncode(filterItem) %>">
                <select name="status">
                    <option value="">待审批</option>
                    <option value="approved" <% If filterStatus = "approved" Then Response.Write "selected" %>>已通过</option>
                    <option value="rejected" <% If filterStatus = "rejected" Then Response.Write "selected" %>>已拒绝</option>
                </select>
                <button type="submit" class="btn-search"><i class="fas fa-search"></i> 搜索</button>
                <a href="price_management.asp?tab=approval" class="btn-reset" style="text-decoration: none; display: inline-block; padding: 10px 20px;"><i class="fas fa-undo"></i> 重置</a>
            </form>
        </div>
        
        <table class="data-table">
            <thead>
                <tr>
                    <th>ID</th>
                    <th>类型</th>
                    <th>供应商</th>
                    <th>物料名称</th>
                    <th>物料编码</th>
                    <th>单价</th>
                    <th>最小起订量</th>
                    <th>审批状态</th>
                    <th>创建时间</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If approvalCount > 0 Then %>
                <% For i = 0 To approvalCount - 1 %>
                <tr>
                    <td>#<%= approvalList(i, 0) %></td>
                    <td>
                        <%
                            Dim apLabel : apLabel = approvalList(i, 6) & ""
                            If apLabel = "Packaging" Then
                                Response.Write "<span style='background:rgba(156,39,176,0.2);color:#CE93D8;padding:3px 10px;border-radius:12px;font-size:12px;'>包装物</span>"
                            ElseIf apLabel = "Bottle" Then
                                Response.Write "<span style='background:rgba(33,150,243,0.2);color:#64B5F6;padding:3px 10px;border-radius:12px;font-size:12px;'>瓶子</span>"
                            ElseIf apLabel = "Printing" Then
                                Response.Write "<span style='background:rgba(0,188,212,0.2);color:#00BCD4;padding:3px 10px;border-radius:12px;font-size:12px;'>印刷品</span>"
                            ElseIf apLabel = "SprayHead" Then
                                Response.Write "<span style='background:rgba(255,87,34,0.2);color:#FF5722;padding:3px 10px;border-radius:12px;font-size:12px;'>喷头</span>"
                            ElseIf apLabel = "BaseNote" Then
                                Response.Write "<span style='background:rgba(76,175,80,0.2);color:#81C784;padding:3px 10px;border-radius:12px;font-size:12px;'>基香原料</span>"
                            Else
                                Response.Write "<span style='background:rgba(255,152,0,0.2);color:#FFB74D;padding:3px 10px;border-radius:12px;font-size:12px;'>原料</span>"
                            End If
                        %>
                    </td>
                    <td><%= HTMLEncode(approvalList(i, 7)) %></td>
                    <td><%= HTMLEncode(approvalList(i, 2)) %></td>
                    <td><%= HTMLEncode(approvalList(i, 3)) %></td>
                    <td class="price-value">¥<%= FormatNumber(SafeNum(approvalList(i, 4)), 2) %></td>
                    <td><%= FormatNumber(SafeNum(approvalList(i, 5)), 0) %></td>
                    <td>
                        <% If approvalList(i, 8) = "Approved" Then %>
                        <span class="status-badge status-active"><i class="fas fa-check-circle"></i> 已通过</span>
                        <% ElseIf approvalList(i, 8) = "Rejected" Then %>
                        <span class="status-badge status-expired"><i class="fas fa-times-circle"></i> 已拒绝</span>
                        <% Else %>
                        <span class="status-badge" style="background:rgba(255,152,0,0.2);color:#FF9800;"><i class="fas fa-clock"></i> 待审批</span>
                        <% End If %>
                    </td>
                    <td><% If IsDate(approvalList(i, 10)) Then Response.Write FormatDateTime(approvalList(i, 10), 2) End If %></td>
                    <td>
                        <% If approvalList(i, 8) = "Pending" And canModify Then %>
                        <form method="post" action="price_management.asp?tab=approval" style="display:inline;">
                            <input type="hidden" name="action" value="approve">
                            <input type="hidden" name="priceId" value="<%= approvalList(i, 0) %>">
                            <button type="submit" class="btn-action btn-edit" style="background:linear-gradient(135deg,#4CAF50 0%,#388E3C 100%);"><i class="fas fa-check"></i> 通过</button>
                        </form>
                        <form method="post" action="price_management.asp?tab=approval" style="display:inline;">
                            <input type="hidden" name="action" value="reject">
                            <input type="hidden" name="priceId" value="<%= approvalList(i, 0) %>">
                            <button type="submit" class="btn-action" style="background:linear-gradient(135deg,#f44336 0%,#d32f2f 100%);color:#fff;" onclick="return confirm('确定拒绝该报价吗？')"><i class="fas fa-times"></i> 拒绝</button>
                        </form>
                        <% Else %>
                        <span style="color:#888;font-size:12px;">无操作</span>
                        <% End If %>
                    </td>
                </tr>
                <% Next %>
                <% Else %>
                <tr>
                    <td colspan="10" style="text-align: center; padding: 40px; color: #666;">
                        <i class="fas fa-check-circle" style="font-size: 48px; display: block; margin-bottom: 15px;"></i>
                        暂无待审批数据
                    </td>
                </tr>
                <% End If %>
            </tbody>
        </table>
        <% End If %>
    </div>
    
    <!-- 新增报价模态框 -->
    <div id="addModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-plus-circle"></i> 新增报价</h3>
                <button class="close-btn" onclick="closeModal('addModal')">&times;</button>
            </div>
            <form method="post" action="price_management.asp?tab=manage">
                <input type="hidden" name="action" value="add">
                
                <div class="form-group">
                    <label>供应商 <span class="required">*</span></label>
                    <select name="supplierId" required>
                        <option value="">请选择供应商</option>
                        <% For i = 0 To supplierCount - 1 %>
                        <option value="<%= supplierList(i, 0) %>"><%= HTMLEncode(supplierList(i, 1)) %></option>
                        <% Next %>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>物料名称 <span class="required">*</span></label>
                    <input type="text" name="itemName" required placeholder="请输入物料名称">
                </div>
                
                <div class="form-group">
                    <label>物料编码</label>
                    <input type="text" name="itemCode" placeholder="请输入物料编码">
                </div>
                
                <div class="form-group">
                    <label>单价 <span class="required">*</span></label>
                    <input type="number" name="unitPrice" step="0.01" required placeholder="请输入单价">
                </div>
                
                <div class="form-group">
                    <label>最小起订量</label>
                    <input type="number" name="minOrderQty" step="1" value="1" placeholder="请输入最小起订量">
                </div>
                
                <div class="form-group">
                    <label>生效日期</label>
                    <input type="date" name="effectiveDate" value="<%= Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Right("0" & Day(Now()), 2) %>">
                </div>
                
                <div class="form-group">
                    <label>失效日期</label>
                    <input type="date" name="expiryDate" placeholder="留空表示无期限">
                </div>
                
                <div class="form-group">
                    <label>价格类型 <span class="required">*</span></label>
                    <select name="priceType" required>
                        <option value="RawMaterial">原料</option>
                        <option value="BaseNote">基香原料</option>
                        <option value="Packaging">包装物</option>
                        <option value="Bottle">瓶子</option>
                        <option value="Printing">印刷品</option>
                        <option value="SprayHead">喷头</option>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>折扣类型</label>
                    <select name="discountType" id="addDiscountType" onchange="toggleDiscountRule('add')">
                        <option value="None">无折扣</option>
                        <option value="Quantity">数量阶梯</option>
                        <option value="Tiered">分层折扣</option>
                        <option value="Seasonal">季节性调价</option>
                    </select>
                </div>
                
                <div class="form-group" id="addDiscountRuleGroup" style="display:none;">
                    <label>折扣规则 <span style="font-size:11px;color:#888;">（数量阶梯示例：100,10;500,8 表示100件以上¥10，500件以上¥8）</span></label>
                    <input type="text" name="discountRule" id="addDiscountRule" placeholder="例如：100,9.5;500,8.0">
                </div>
                
                <div class="form-actions">
                    <button type="button" class="btn-cancel" onclick="closeModal('addModal')">取消</button>
                    <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 编辑报价模态框 -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-edit"></i> 编辑报价</h3>
                <button class="close-btn" onclick="closeModal('editModal')">&times;</button>
            </div>
            <form method="post" action="price_management.asp?tab=manage">
                <input type="hidden" name="action" value="edit">
                <input type="hidden" name="priceId" id="editPriceId">
                
                <div class="form-group">
                    <label>供应商 <span class="required">*</span></label>
                    <select name="supplierId" id="editSupplierId" required>
                        <option value="">请选择供应商</option>
                        <% For i = 0 To supplierCount - 1 %>
                        <option value="<%= supplierList(i, 0) %>"><%= HTMLEncode(supplierList(i, 1)) %></option>
                        <% Next %>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>物料名称 <span class="required">*</span></label>
                    <input type="text" name="itemName" id="editItemName" required placeholder="请输入物料名称">
                </div>
                
                <div class="form-group">
                    <label>物料编码</label>
                    <input type="text" name="itemCode" id="editItemCode" placeholder="请输入物料编码">
                </div>
                
                <div class="form-group">
                    <label>单价 <span class="required">*</span></label>
                    <input type="number" name="unitPrice" id="editUnitPrice" step="0.01" required placeholder="请输入单价">
                </div>
                
                <div class="form-group">
                    <label>最小起订量</label>
                    <input type="number" name="minOrderQty" id="editMinOrderQty" step="1" placeholder="请输入最小起订量">
                </div>
                
                <div class="form-group">
                    <label>生效日期</label>
                    <input type="date" name="effectiveDate" id="editEffectiveDate">
                </div>
                
                <div class="form-group">
                    <label>失效日期</label>
                    <input type="date" name="expiryDate" id="editExpiryDate" placeholder="留空表示无期限">
                </div>
                
                <div class="form-group">
                    <label>价格类型 <span class="required">*</span></label>
                    <select name="priceType" id="editPriceType" required>
                        <option value="RawMaterial">原料</option>
                        <option value="BaseNote">基香原料</option>
                        <option value="Packaging">包装物</option>
                        <option value="Bottle">瓶子</option>
                        <option value="Printing">印刷品</option>
                        <option value="SprayHead">喷头</option>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>折扣类型</label>
                    <select name="discountType" id="editDiscountType" onchange="toggleDiscountRule('edit')">
                        <option value="None">无折扣</option>
                        <option value="Quantity">数量阶梯</option>
                        <option value="Tiered">分层折扣</option>
                        <option value="Seasonal">季节性调价</option>
                    </select>
                </div>
                
                <div class="form-group" id="editDiscountRuleGroup" style="display:none;">
                    <label>折扣规则</label>
                    <input type="text" name="discountRule" id="editDiscountRule" placeholder="例如：100,9.5;500,8.0">
                </div>
                
                <div class="form-actions">
                    <button type="button" class="btn-cancel" onclick="closeModal('editModal')">取消</button>
                    <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        // V11: 折扣规则显示切换
        function toggleDiscountRule(prefix) {
            var typeSelect = document.getElementById(prefix + 'DiscountType');
            var ruleGroup = document.getElementById(prefix + 'DiscountRuleGroup');
            if (typeSelect && ruleGroup) {
                ruleGroup.style.display = typeSelect.value === 'None' ? 'none' : 'block';
            }
        }
        
        function openAddModal() {
            document.getElementById('addModal').style.display = 'block';
            document.getElementById('addDiscountType').value = 'None';
            document.getElementById('addDiscountRuleGroup').style.display = 'none';
            document.getElementById('addDiscountRule').value = '';
        }
        
        function openEditModal(priceId, supplierId, itemName, itemCode, unitPrice, minOrderQty, effectiveDate, expiryDate, priceType, discountType, discountRule) {
            document.getElementById('editPriceId').value = priceId;
            document.getElementById('editSupplierId').value = supplierId;
            document.getElementById('editItemName').value = itemName;
            document.getElementById('editItemCode').value = itemCode;
            document.getElementById('editUnitPrice').value = unitPrice;
            document.getElementById('editMinOrderQty').value = minOrderQty;
            
            // 设置价格类型
            var ptSelect = document.getElementById('editPriceType');
            if (priceType && priceType != 'null' && priceType != '') {
                ptSelect.value = priceType;
            } else {
                ptSelect.value = 'RawMaterial';
            }
            
            // V11: 设置折扣类型和规则
            var dtSelect = document.getElementById('editDiscountType');
            if (discountType && discountType != 'null' && discountType != '') {
                dtSelect.value = discountType;
            } else {
                dtSelect.value = 'None';
            }
            document.getElementById('editDiscountRule').value = (discountRule && discountRule != 'null') ? discountRule : '';
            toggleDiscountRule('edit');
            
            // 格式化日期
            if (effectiveDate && effectiveDate != 'null') {
                var d = new Date(effectiveDate);
                if (!isNaN(d.getTime())) {
                    document.getElementById('editEffectiveDate').value = d.toISOString().split('T')[0];
                }
            }
            if (expiryDate && expiryDate != 'null') {
                var d = new Date(expiryDate);
                if (!isNaN(d.getTime())) {
                    document.getElementById('editExpiryDate').value = d.toISOString().split('T')[0];
                }
            }
            
            document.getElementById('editModal').style.display = 'block';
        }
        
        function closeModal(modalId) {
            document.getElementById(modalId).style.display = 'none';
        }
        
        window.onclick = function(event) {
            if (event.target.classList.contains('modal')) {
                event.target.style.display = 'none';
            }
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
