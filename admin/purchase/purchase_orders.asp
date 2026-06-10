﻿﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection()

' ========== 确保必要字段存在 ==========
On Error Resume Next
conn.Execute "SELECT OrderType FROM PurchaseOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PurchaseOrders ADD OrderType NVARCHAR(20) DEFAULT 'RawMaterial'"
conn.Execute "SELECT CategoryCode FROM PurchaseOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PurchaseOrders ADD CategoryCode NVARCHAR(20) DEFAULT 'RAW'"
' ========== 分类-类型映射表 ==========
conn.Execute "SELECT TOP 1 1 FROM PurchaseCategories"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PurchaseCategories (CategoryID INT IDENTITY(1,1) PRIMARY KEY, CategoryCode NVARCHAR(20) NOT NULL UNIQUE, CategoryName NVARCHAR(50) NOT NULL, IconClass NVARCHAR(50), SortOrder INT DEFAULT 0, IsActive BIT DEFAULT 1)"
    ' 插入7个默认分类（与现有CategoryCode一致）
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('RAW','原材料','fas fa-flask',1)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('BASE','基香原料','fas fa-leaf',2)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('PACK','包装材料','fas fa-box',3)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('BOTTLE','瓶子包装','fas fa-wine-bottle',4)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('PRINTING','印刷品','fas fa-print',5)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('SPRAYHEAD','喷头配件','fas fa-spray-can',6)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('MARKET','营销物料','fas fa-ad',7)"
End If
conn.Execute "SELECT TOP 1 1 FROM PurchaseCategoryTypes"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PurchaseCategoryTypes (MapID INT IDENTITY(1,1) PRIMARY KEY, CategoryCode NVARCHAR(20) NOT NULL, OrderType NVARCHAR(30) NOT NULL, IsDefault BIT DEFAULT 0, FOREIGN KEY (CategoryCode) REFERENCES PurchaseCategories(CategoryCode))"
    ' 插入默认映射关系（与现有代码保持一致）
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('RAW','RawMaterial',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('BASE','RawMaterial',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('PACK','Packaging',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('BOTTLE','Bottle',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('PRINTING','Printing',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('SPRAYHEAD','SprayHead',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('MARKET','Packaging',1)"
End If
' V12: 状态变更日志表
conn.Execute "SELECT TOP 1 1 FROM PurchaseOrderStatusLog"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PurchaseOrderStatusLog (LogID INT IDENTITY(1,1) PRIMARY KEY, PurchaseID INT NOT NULL, FromStatus NVARCHAR(30), ToStatus NVARCHAR(30) NOT NULL, ChangedBy NVARCHAR(50), ChangedAt DATETIME DEFAULT GETDATE(), Remarks NVARCHAR(200))"
End If
On Error GoTo 0

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

' ========== 生成采购单号：PO-YYYYMMDD-NNN ==========
Function GeneratePurchaseNo()
    Dim today, prefix, sql, countNum, suffix
    On Error Resume Next
    today = Date()
    prefix = "PO-" & Year(today) & Right("0" & Month(today), 2) & Right("0" & Day(today), 2) & "-"
    
    ' 查询今日已有多少订单
    sql = "SELECT COUNT(*) FROM PurchaseOrders WHERE PurchaseNo LIKE '" & prefix & "%'"
    countNum = SafeNum(GetScalar(sql))
    
    ' 生成3位序号
    If Err.Number = 0 And IsNumeric(countNum) Then
        suffix = Right("000" & (countNum + 1), 3)
    Else
        suffix = "001"
        Err.Clear
    End If
    On Error GoTo 0
    GeneratePurchaseNo = prefix & suffix
End Function

' ========== 获取状态中文名称 ==========
Function GetStatusName(statusCode)
    Dim sc
    sc = UCase(Trim(statusCode & ""))
    Select Case sc
        Case "DRAFT"
            GetStatusName = "草稿"
        Case "PENDING"
            GetStatusName = "待处理"
        Case "SUBMITTED"
            GetStatusName = "已提交"
        Case "APPROVED"
            GetStatusName = "已审批"
        Case "FINANCEAPPROVED"
            GetStatusName = "财务已审批"
        Case "ORDERED"
            GetStatusName = "已下单"
        Case "PARTIALRECEIVED"
            GetStatusName = "部分收货"
        Case "RECEIVED"
            GetStatusName = "已收货"
        Case "COMPLETE", "COMPLETED"
            GetStatusName = "已完成"
        Case "COMPLETERECEIVED"
            GetStatusName = "已完成"
        Case "REJECTED"
            GetStatusName = "已拒绝"
        Case "CANCELLED"
            GetStatusName = "已取消"
        Case Else
            ' 处理未知拼接状态：优先匹配包含的关键词
            If InStr(sc, "COMPLETE") > 0 Then
                GetStatusName = "已完成"
            ElseIf InStr(sc, "RECEIVED") > 0 Then
                GetStatusName = "已收货"
            ElseIf InStr(sc, "APPROVED") > 0 Then
                GetStatusName = "已审批"
            ElseIf InStr(sc, "SUBMITTED") > 0 Then
                GetStatusName = "已提交"
            Else
                GetStatusName = statusCode
            End If
    End Select
End Function

' ========== 获取状态样式类 ==========
Function GetStatusClass(statusCode)
    Dim sc
    sc = UCase(Trim(statusCode & ""))
    Select Case sc
        Case "DRAFT"
            GetStatusClass = "status-draft"
        Case "PENDING"
            GetStatusClass = "status-submitted"
        Case "SUBMITTED"
            GetStatusClass = "status-submitted"
        Case "APPROVED", "FINANCEAPPROVED"
            GetStatusClass = "status-approved"
        Case "ORDERED"
            GetStatusClass = "status-ordered"
        Case "PARTIALRECEIVED"
            GetStatusClass = "status-partial"
        Case "RECEIVED"
            GetStatusClass = "status-received"
        Case "COMPLETE", "COMPLETED", "COMPLETERECEIVED"
            GetStatusClass = "status-completed"
        Case "REJECTED"
            GetStatusClass = "status-rejected"
        Case "CANCELLED"
            GetStatusClass = "status-rejected"
        Case Else
            If InStr(sc, "COMPLETE") > 0 Then
                GetStatusClass = "status-completed"
            ElseIf InStr(sc, "RECEIVED") > 0 Then
                GetStatusClass = "status-received"
            ElseIf InStr(sc, "APPROVED") > 0 Then
                GetStatusClass = "status-approved"
            Else
                GetStatusClass = "status-draft"
            End If
    End Select
End Function

' ========== 获取分类名称 ==========
Function GetCategoryName(catCode)
    If catCode = "RAW" Then
        GetCategoryName = "原材料"
    ElseIf catCode = "BASE" Then
        GetCategoryName = "基香原料"
    ElseIf catCode = "PACK" Then
        GetCategoryName = "包装材料"
    ElseIf catCode = "BOTTLE" Then
        GetCategoryName = "瓶子包装"
    ElseIf catCode = "PRINTING" Then
        GetCategoryName = "印刷品"
    ElseIf catCode = "SPRAYHEAD" Then
        GetCategoryName = "喷头配件"
    ElseIf catCode = "MARKET" Then
        GetCategoryName = "营销物料"
    Else
        GetCategoryName = catCode
    End If
End Function

' ========== 根据分类代码获取采购类型 ==========
Function GetOrderTypeByCategory(catCode)
    If catCode = "RAW" Then
        GetOrderTypeByCategory = "RawMaterial"
    ElseIf catCode = "BASE" Then
        GetOrderTypeByCategory = "RawMaterial"
    ElseIf catCode = "PACK" Then
        GetOrderTypeByCategory = "Packaging"
    ElseIf catCode = "BOTTLE" Then
        GetOrderTypeByCategory = "Bottle"
    ElseIf catCode = "PRINTING" Then
        GetOrderTypeByCategory = "Printing"
    ElseIf catCode = "SPRAYHEAD" Then
        GetOrderTypeByCategory = "SprayHead"
    ElseIf catCode = "MARKET" Then
        GetOrderTypeByCategory = "Packaging"
    Else
        GetOrderTypeByCategory = "RawMaterial"
    End If
End Function

' ========== 获取供应商名称（带错误保护） ==========
Function GetSupplierName(supplierID)
    Dim sql, rs, result
    result = "未知供应商"
    On Error Resume Next
    sql = "SELECT SupplierName FROM Suppliers WHERE SupplierID=" & SafeNum(supplierID)
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs.EOF Then
            result = CStr(rs("SupplierName"))
        End If
        rs.Close
        Set rs = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
    GetSupplierName = result
End Function

' ========== 处理POST请求 ==========
Dim action, message, messageType
action = Trim(Request.Form("action"))
message = ""
messageType = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "create" Then
        ' 创建采购订单
        Dim newPurchaseNo, newSupplierID, newCategory, newExpectedDate, newRemarks
        newPurchaseNo = GeneratePurchaseNo()
        newSupplierID = SafeNum(Request.Form("supplier_id"))
        newCategory = Trim(Request.Form("category_code"))
        newOrderType = Trim(Request.Form("order_type"))
        If newOrderType = "" Then newOrderType = "RawMaterial"
        newExpectedDate = Trim(Request.Form("expected_date"))
        newRemarks = Trim(Request.Form("remarks"))
        
        If newSupplierID = 0 Then
            message = "请选择供应商"
            messageType = "error"
        ElseIf newCategory = "" Then
            message = "请选择采购分类"
            messageType = "error"
        Else
            Dim sqlCreate, newPurchaseID
            sqlCreate = "INSERT INTO PurchaseOrders (PurchaseNo, SupplierID, CategoryCode, OrderType, OrderDate, ExpectedDate, TotalAmount, Status, Remarks, CreatedBy, CreatedAt, UpdatedAt) VALUES ('" & _
                newPurchaseNo & "', " & newSupplierID & ", '" & newCategory & "', '" & newOrderType & "', GETDATE(), "
            
            If newExpectedDate <> "" Then
                sqlCreate = sqlCreate & "'" & newExpectedDate & "'"
            Else
                sqlCreate = sqlCreate & "Null"
            End If
            
            sqlCreate = sqlCreate & ", 0, 'Draft', "
            
            If newRemarks <> "" Then
                sqlCreate = sqlCreate & "'" & Replace(newRemarks, "'", "''") & "'"
            Else
                sqlCreate = sqlCreate & "Null"
            End If
            
            sqlCreate = sqlCreate & ", " & SafeNum(Session("AdminID")) & ", GETDATE(), GETDATE())"
            
            On Error Resume Next
            Err.Clear  ' 清除残留错误状态
            conn.Execute sqlCreate
            
            ' 检查ADONET错误和VBScript错误
            Dim adoErrMsg
            adoErrMsg = ""
            If conn.Errors.Count > 0 Then
                Dim ae
                For Each ae In conn.Errors
                    adoErrMsg = adoErrMsg & "[" & ae.Number & "]" & ae.Description & " "
                Next
                conn.Errors.Clear
            End If
            
            If Err.Number = 0 And adoErrMsg = "" Then
                ' 验证：确认订单已写入数据库
                Dim verifySQL, verifyRs, orderExists
                orderExists = False
                verifySQL = "SELECT COUNT(*) FROM PurchaseOrders WHERE PurchaseNo='" & Replace(newPurchaseNo, "'", "''") & "'"
                Set verifyRs = conn.Execute(verifySQL)
                If Not verifyRs.EOF Then
                    If CLng(verifyRs(0)) > 0 Then
                        orderExists = True
                    End If
                End If
                verifyRs.Close
                Set verifyRs = Nothing
                
                If Not orderExists Then
                    ' 插入成功但验证失败 - 记录调试信息
                    Session("LastDBError") = "插入后验证失败: PurchaseNo=" & newPurchaseNo & " | SQL=" & sqlCreate
                    message = "创建失败：数据库写入验证未通过"
                    messageType = "error"
                    Err.Clear
                Else
                    ' 获取新插入的ID
                    Dim rsNewID
                    Set rsNewID = conn.Execute("SELECT SCOPE_IDENTITY()")
                    If Not rsNewID.EOF Then
                        newPurchaseID = rsNewID(0)
                        If IsNull(newPurchaseID) Then newPurchaseID = 0
                    Else
                        newPurchaseID = 0
                    End If
                    rsNewID.Close
                    Set rsNewID = Nothing
                    
                    ' 处理明细行
                    Dim itemCount, i, itemName, itemCode, spec, unit, qty, price, lineTotal
                    itemCount = SafeNum(Request.Form("item_count"))
                    
                    For i = 1 To itemCount
                        itemName = Trim(Request.Form("item_name_" & i))
                        If itemName <> "" Then
                            itemCode = Trim(Request.Form("item_code_" & i))
                            spec = Trim(Request.Form("spec_" & i))
                            unit = Trim(Request.Form("unit_" & i))
                            qty = SafeNum(Request.Form("qty_" & i))
                            price = SafeNum(Request.Form("price_" & i))
                            lineTotal = qty * price
                            
                            Dim sqlDetail
                            sqlDetail = "INSERT INTO PurchaseOrderDetails (PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, UnitPrice, TotalPrice, ReceivedQty) VALUES (" & _
                                newPurchaseID & ", '" & Replace(itemName, "'", "''") & "', "
                            
                            If itemCode <> "" Then
                                sqlDetail = sqlDetail & "'" & Replace(itemCode, "'", "''") & "'"
                            Else
                                sqlDetail = sqlDetail & "Null"
                            End If
                            
                            sqlDetail = sqlDetail & ", "
                            
                            If spec <> "" Then
                                sqlDetail = sqlDetail & "'" & Replace(spec, "'", "''") & "'"
                            Else
                                sqlDetail = sqlDetail & "Null"
                            End If
                            
                            sqlDetail = sqlDetail & ", "
                            
                            If unit <> "" Then
                                sqlDetail = sqlDetail & "'" & Replace(unit, "'", "''") & "'"
                            Else
                                sqlDetail = sqlDetail & "Null"
                            End If
                            
                            sqlDetail = sqlDetail & ", " & qty & ", " & price & ", " & lineTotal & ", 0)"
                            
                            conn.Execute sqlDetail
                        End If
                    Next
                    
                    ' 更新订单总金额（使用VBScript计算SUM，避免Access子查询不可更新问题）
                    If newPurchaseID > 0 Then
                        Dim rsTotalAmt, calcTotal
                        calcTotal = 0
                        Set rsTotalAmt = conn.Execute("SELECT SUM(CAST(ISNULL(TotalPrice,0) AS FLOAT)) as OrderTotal FROM PurchaseOrderDetails WHERE PurchaseID=" & newPurchaseID)
                        If Not rsTotalAmt Is Nothing Then
                            If Not rsTotalAmt.EOF Then
                                If Not IsNull(rsTotalAmt("OrderTotal")) Then
                                    calcTotal = CDbl(rsTotalAmt("OrderTotal"))
                                End If
                            End If
                            rsTotalAmt.Close
                        End If
                        Set rsTotalAmt = Nothing
                        conn.Execute "UPDATE PurchaseOrders SET TotalAmount=" & calcTotal & " WHERE PurchaseID=" & newPurchaseID
                    End If
                    
                    ' 记录成功订单号供调试
                    Session("LastPurchaseNo") = newPurchaseNo
                    
                    ' V12: 记录初始状态
                    If newPurchaseID > 0 Then
                        conn.Execute "INSERT INTO PurchaseOrderStatusLog (PurchaseID, FromStatus, ToStatus, ChangedBy, ChangedAt, Remarks) VALUES (" & newPurchaseID & ", NULL, 'Draft', '" & CStr(Session("AdminUsername") & "") & "', GETDATE(), '创建采购订单')"
                    End If
                    
                    Response.Redirect "purchase_orders.asp?msg=created"
                    Response.End
                End If
            Else
                Dim errDetail
                errDetail = Err.Description
                If adoErrMsg <> "" Then
                    errDetail = errDetail & " | ADO错误: " & adoErrMsg
                End If
                ' 保存SQL到Session用于调试
                Session("LastDBError") = "创建订单失败: " & errDetail & " | SQL=" & sqlCreate
                message = "创建失败：" & errDetail
                messageType = "error"
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
    ElseIf action = "update" Then
        ' 更新采购订单
        Dim editPurchaseID, editSupplierID, editCategory, editExpectedDate, editRemarks
        editPurchaseID = SafeNum(Request.Form("purchase_id"))
        editSupplierID = SafeNum(Request.Form("supplier_id"))
        editCategory = Trim(Request.Form("category_code"))
        editOrderType = Trim(Request.Form("order_type"))
        If editOrderType = "" Then editOrderType = "RawMaterial"
        editExpectedDate = Trim(Request.Form("expected_date"))
        editRemarks = Trim(Request.Form("remarks"))
        
        If editPurchaseID = 0 Then
            message = "订单ID无效"
            messageType = "error"
        ElseIf editSupplierID = 0 Then
            message = "请选择供应商"
            messageType = "error"
        ElseIf editCategory = "" Then
            message = "请选择采购分类"
            messageType = "error"
        Else
            ' 检查状态是否为Draft
            Dim checkStatus
            checkStatus = ""
            On Error Resume Next
            Set rsCheck = conn.Execute("SELECT Status FROM PurchaseOrders WHERE PurchaseID=" & editPurchaseID)
            If Err.Number = 0 Then
                If Not rsCheck.EOF Then
                    checkStatus = CStr(rsCheck("Status"))
                End If
                rsCheck.Close
                Set rsCheck = Nothing
            End If
            On Error GoTo 0
            
            If checkStatus <> "Draft" Then
                message = "只有草稿状态的订单可以编辑"
                messageType = "error"
            Else
                Dim sqlUpdate
                sqlUpdate = "UPDATE PurchaseOrders SET SupplierID=" & editSupplierID & ", CategoryCode='" & editCategory & "', OrderType='" & editOrderType & "', ExpectedDate="
                
                If editExpectedDate <> "" Then
                    sqlUpdate = sqlUpdate & "'" & editExpectedDate & "'"
                Else
                    sqlUpdate = sqlUpdate & "Null"
                End If
                
                sqlUpdate = sqlUpdate & ", Remarks="
                
                If editRemarks <> "" Then
                    sqlUpdate = sqlUpdate & "'" & Replace(editRemarks, "'", "''") & "'"
                Else
                    sqlUpdate = sqlUpdate & "Null"
                End If
                
                sqlUpdate = sqlUpdate & ", UpdatedAt= GETDATE() WHERE PurchaseID=" & editPurchaseID
                
                On Error Resume Next
                conn.Execute sqlUpdate
                
                If Err.Number = 0 Then
                    ' 删除旧明细
                    conn.Execute "DELETE FROM PurchaseOrderDetails WHERE PurchaseID=" & editPurchaseID
                    
                    ' 插入新明细
                    Dim editItemCount, j
                    editItemCount = SafeNum(Request.Form("item_count"))
                    
                    For j = 1 To editItemCount
                        itemName = Trim(Request.Form("item_name_" & j))
                        If itemName <> "" Then
                            itemCode = Trim(Request.Form("item_code_" & j))
                            spec = Trim(Request.Form("spec_" & j))
                            unit = Trim(Request.Form("unit_" & j))
                            qty = SafeNum(Request.Form("qty_" & j))
                            price = SafeNum(Request.Form("price_" & j))
                            lineTotal = qty * price
                            
                            Dim sqlDetailUpdate
                            sqlDetailUpdate = "INSERT INTO PurchaseOrderDetails (PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, UnitPrice, TotalPrice, ReceivedQty) VALUES (" & _
                                editPurchaseID & ", '" & Replace(itemName, "'", "''") & "', "
                            
                            If itemCode <> "" Then
                                sqlDetailUpdate = sqlDetailUpdate & "'" & Replace(itemCode, "'", "''") & "'"
                            Else
                                sqlDetailUpdate = sqlDetailUpdate & "Null"
                            End If
                            
                            sqlDetailUpdate = sqlDetailUpdate & ", "
                            
                            If spec <> "" Then
                                sqlDetailUpdate = sqlDetailUpdate & "'" & Replace(spec, "'", "''") & "'"
                            Else
                                sqlDetailUpdate = sqlDetailUpdate & "Null"
                            End If
                            
                            sqlDetailUpdate = sqlDetailUpdate & ", "
                            
                            If unit <> "" Then
                                sqlDetailUpdate = sqlDetailUpdate & "'" & Replace(unit, "'", "''") & "'"
                            Else
                                sqlDetailUpdate = sqlDetailUpdate & "Null"
                            End If
                            
                            sqlDetailUpdate = sqlDetailUpdate & ", " & qty & ", " & price & ", " & lineTotal & ", 0)"
                            
                            conn.Execute sqlDetailUpdate
                        End If
                    Next
                    
                    ' 更新订单总金额（使用VBScript计算SUM，避免Access子查询不可更新问题）
                    Dim rsEditTotal, calcEditTotal
                    calcEditTotal = 0
                    Set rsEditTotal = conn.Execute("SELECT SUM(CAST(ISNULL(TotalPrice,0) AS FLOAT)) as OrderTotal FROM PurchaseOrderDetails WHERE PurchaseID=" & editPurchaseID)
                    If Not rsEditTotal Is Nothing Then
                        If Not rsEditTotal.EOF Then
                            If Not IsNull(rsEditTotal("OrderTotal")) Then
                                calcEditTotal = CDbl(rsEditTotal("OrderTotal"))
                            End If
                        End If
                        rsEditTotal.Close
                    End If
                    Set rsEditTotal = Nothing
                    conn.Execute "UPDATE PurchaseOrders SET TotalAmount=" & calcEditTotal & " WHERE PurchaseID=" & editPurchaseID
                    
                    Response.Redirect "purchase_orders.asp?msg=updated"
                    Response.End
                Else
                    message = "更新失败：" & Err.Description
                    messageType = "error"
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
        
    ElseIf action = "changestatus" Then
        ' 状态流转
        Dim statusPurchaseID, newStatus
        statusPurchaseID = SafeNum(Request.Form("purchase_id"))
        newStatus = Trim(Request.Form("new_status"))
        
        If statusPurchaseID > 0 And newStatus <> "" Then
            Dim canChange
            canChange = False
            
            ' 检查当前状态是否允许流转
            Dim currentStatus
            currentStatus = ""
            On Error Resume Next
            Set rsStatus = conn.Execute("SELECT Status FROM PurchaseOrders WHERE PurchaseID=" & statusPurchaseID)
            If Err.Number = 0 Then
                If Not rsStatus.EOF Then
                    currentStatus = CStr(rsStatus("Status"))
                End If
                rsStatus.Close
                Set rsStatus = Nothing
            End If
            On Error GoTo 0
            
            ' 验证状态流转规则
            If newStatus = "Submitted" And currentStatus = "Draft" Then
                canChange = True  ' 任何人都可以提交草稿
            ElseIf newStatus = "FinanceApproved" And currentStatus = "Submitted" Then
                canChange = True  ' 提交后可以财务审批
            ElseIf newStatus = "Ordered" And currentStatus = "FinanceApproved" Then
                canChange = True  ' 已审批后可以下单
            ElseIf newStatus = "PartialReceived" And currentStatus = "Ordered" Then
                canChange = True  ' 已下单后可以部分收货（通过收货页面操作）
            ElseIf newStatus = "Received" And (currentStatus = "Ordered" Or currentStatus = "PartialReceived") Then
                canChange = True  ' 已下单或部分收货后可以收货
            ElseIf newStatus = "Completed" And currentStatus = "Received" Then
                ' 只有经理可以完成
                If isManager Then
                    canChange = True
                End If
            End If
            
            If canChange Then
                Dim sqlStatus
                sqlStatus = "UPDATE PurchaseOrders SET Status='" & newStatus & "', UpdatedAt= GETDATE()"
                
                If newStatus = "Completed" Then
                    sqlStatus = sqlStatus & ", ApprovedBy=" & SafeNum(Session("AdminID")) & ", ApprovedAt= GETDATE()"
                End If
                
                sqlStatus = sqlStatus & " WHERE PurchaseID=" & statusPurchaseID
                
                On Error Resume Next
                conn.Execute sqlStatus
                If Err.Number = 0 Then
                    ' V12: 记录状态变更日志
                    Dim logChangedBy : logChangedBy = CStr(Session("AdminUsername") & "")
                    If logChangedBy = "" Then logChangedBy = "System"
                    conn.Execute "INSERT INTO PurchaseOrderStatusLog (PurchaseID, FromStatus, ToStatus, ChangedBy, ChangedAt, Remarks) VALUES (" & statusPurchaseID & ", '" & currentStatus & "', '" & newStatus & "', '" & logChangedBy & "', GETDATE(), '状态变更: " & GetStatusName(currentStatus) & " → " & GetStatusName(newStatus) & "')"
                    Response.Redirect "purchase_orders.asp?msg=statuschanged"
                    Response.End
                Else
                    message = "状态更新失败：" & Err.Description
                    messageType = "error"
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                message = "当前状态不允许此操作"
                messageType = "error"
            End If
        End If
    End If
End If

' V11: 批量操作处理
Dim batchAction : batchAction = Trim(Request.Form("batch_action"))
If batchAction <> "" Then
    Dim batchIDs : batchIDs = Trim(Request.Form("batch_ids"))
    If batchIDs <> "" Then
        Dim batchNewStatus
        Select Case batchAction
            Case "approve" : batchNewStatus = "FinanceApproved"
            Case "order" : batchNewStatus = "Ordered"
            Case "receive" : batchNewStatus = "Received"
        End Select
        
        If batchNewStatus <> "" Then
            On Error Resume Next
            Dim batchArr, bi, bid, batchCurrentStatus
            batchArr = Split(batchIDs, ",")
            Dim batchSuccess : batchSuccess = 0
            Dim batchFail : batchFail = 0
            
            For bi = 0 To UBound(batchArr)
                bid = Trim(batchArr(bi))
                If IsNumeric(bid) Then
                    ' Validate status transition
                    batchCurrentStatus = ""
                    Dim rsBC : Set rsBC = conn.Execute("SELECT Status FROM PurchaseOrders WHERE PurchaseID=" & CLng(bid))
                    If Not rsBC Is Nothing Then
                        If Not rsBC.EOF Then batchCurrentStatus = CStr(rsBC("Status"))
                        rsBC.Close : Set rsBC = Nothing
                    End If
                    
                    Dim canBatch : canBatch = False
                    If batchAction = "approve" And batchCurrentStatus = "Submitted" Then canBatch = True
                    If batchAction = "order" And batchCurrentStatus = "FinanceApproved" Then canBatch = True
                    If batchAction = "receive" And (batchCurrentStatus = "Ordered" Or batchCurrentStatus = "PartialReceived") Then canBatch = True
                    
                    If canBatch Then
                        conn.Execute "UPDATE PurchaseOrders SET Status='" & batchNewStatus & "', UpdatedAt=GETDATE() WHERE PurchaseID=" & CLng(bid)
                        If Err.Number = 0 Then batchSuccess = batchSuccess + 1 Else Err.Clear : batchFail = batchFail + 1
                    Else
                        batchFail = batchFail + 1
                    End If
                End If
            Next
            On Error GoTo 0
            
            If batchFail = 0 Then
                Response.Redirect "purchase_orders.asp?msg=batchdone&count=" & batchSuccess
            Else
                message = "批量操作: " & batchSuccess & " 成功, " & batchFail & " 失败"
                messageType = IIf(batchSuccess > 0, "success", "error")
            End If
        End If
    End If
End If

' V11: 复制订单处理
If action = "copy" Then
    Dim copyID : copyID = SafeNum(Request.Form("purchase_id"))
    If copyID > 0 Then
        ' Get source order
        Dim rsCopy : Set rsCopy = conn.Execute("SELECT * FROM PurchaseOrders WHERE PurchaseID=" & copyID)
        If Not rsCopy Is Nothing And Not rsCopy.EOF Then
            Dim newPNo : newPNo = GeneratePurchaseNo()
            Dim catCode : catCode = CStr(rsCopy("CategoryCode") & "")
            Dim otCode : otCode = CStr(rsCopy("OrderType") & "")
            If catCode = "" Then
                If otCode = "Packaging" Then catCode = "PACK"
                If otCode = "Bottle" Then catCode = "BOTTLE"
                If otCode = "Printing" Then catCode = "PRINTING"
                If otCode = "SprayHead" Then catCode = "SPRAYHEAD"
            End If
            
            conn.Execute "INSERT INTO PurchaseOrders (PurchaseNo, SupplierID, CategoryCode, OrderType, OrderDate, ExpectedDate, TotalAmount, Status, Remarks, CreatedBy, CreatedAt) VALUES ('" & newPNo & "', " & SafeNum(rsCopy("SupplierID")) & ", '" & SafeSQL(catCode) & "', '" & SafeSQL(otCode) & "', GETDATE(), " & IIf(IsNull(rsCopy("ExpectedDate")), "DATEADD(DAY,14,GETDATE())", "'" & FormatDateTime(rsCopy("ExpectedDate"),2) & "'") & ", " & SafeNum(rsCopy("TotalAmount")) & ", 'Draft', '复制自 " & rsCopy("PurchaseNo") & "', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
            
            ' Get new order ID
            Set rsNewID = conn.Execute("SELECT SCOPE_IDENTITY()")
            Dim newOrderID : newOrderID = 0
            If Not rsNewID Is Nothing And Not rsNewID.EOF Then newOrderID = CLng(rsNewID(0))
            rsNewID.Close : Set rsNewID = Nothing
            rsCopy.Close : Set rsCopy = Nothing
            
            ' Copy details
            If newOrderID > 0 Then
                Dim rsCopyDet : Set rsCopyDet = conn.Execute("SELECT * FROM PurchaseOrderDetails WHERE PurchaseID=" & copyID)
                If Not rsCopyDet Is Nothing Then
                    Do While Not rsCopyDet.EOF
                        conn.Execute "INSERT INTO PurchaseOrderDetails (PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, UnitPrice, TotalPrice, ReceivedQty) VALUES (" & newOrderID & ", '" & SafeSQL(rsCopyDet("ItemName") & "") & "', '" & SafeSQL(rsCopyDet("ItemCode") & "") & "', '" & SafeSQL(rsCopyDet("Specification") & "") & "', '" & SafeSQL(rsCopyDet("Unit") & "") & "', " & SafeNum(rsCopyDet("Quantity")) & ", " & SafeNum(rsCopyDet("UnitPrice")) & ", " & SafeNum(rsCopyDet("TotalPrice")) & ", 0)"
                        rsCopyDet.MoveNext
                    Loop
                    rsCopyDet.Close
                End If
                Set rsCopyDet = Nothing
                
                ' Update total
                Dim rsCopyTotal : Set rsCopyTotal = conn.Execute("SELECT SUM(CAST(ISNULL(TotalPrice,0) AS FLOAT)) FROM PurchaseOrderDetails WHERE PurchaseID=" & newOrderID)
                Dim copyTotal : copyTotal = 0
                If Not rsCopyTotal Is Nothing Then
                    If Not rsCopyTotal.EOF And Not IsNull(rsCopyTotal(0)) Then copyTotal = CDbl(rsCopyTotal(0))
                    rsCopyTotal.Close
                End If
                Set rsCopyTotal = Nothing
                conn.Execute "UPDATE PurchaseOrders SET TotalAmount=" & copyTotal & " WHERE PurchaseID=" & newOrderID
                
                ' V12: 记录复制创建状态
                conn.Execute "INSERT INTO PurchaseOrderStatusLog (PurchaseID, FromStatus, ToStatus, ChangedBy, ChangedAt, Remarks) VALUES (" & newOrderID & ", NULL, 'Draft', '" & CStr(Session("AdminUsername") & "") & "', GETDATE(), '复制自订单 " & rsCopy("PurchaseNo") & "')"
                
                Response.Redirect "purchase_orders.asp?msg=copied&id=" & newOrderID
                Response.End
            End If
        End If
    End If
End If

' ========== 处理消息提示 ==========
Dim queryMsg
queryMsg = Trim(Request.QueryString("msg"))
If queryMsg = "created" Then
    message = "采购订单创建成功"
    messageType = "success"
    
    ' 验证最后创建的订单是否确实存在
    Dim lastPNo
    lastPNo = Session("LastPurchaseNo")
    If lastPNo <> "" Then
        Dim verifyCreated
        On Error Resume Next
        verifyCreated = SafeNum(GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE PurchaseNo='" & Replace(lastPNo, "'", "''") & "'"))
        If Err.Number <> 0 Or verifyCreated = 0 Then
            message = "采购订单可能未成功保存（调试：" & lastPNo & " 未在数据库中查到）"
            messageType = "error"
            Err.Clear
        End If
        On Error GoTo 0
        Session("LastPurchaseNo") = ""  ' 清除调试记录
    End If
ElseIf queryMsg = "updated" Then
    message = "采购订单更新成功"
    messageType = "success"
ElseIf queryMsg = "statuschanged" Then
    message = "状态更新成功"
    messageType = "success"
ElseIf queryMsg = "batchdone" Then
    message = "批量操作完成：" & Request.QueryString("count") & " 个订单处理成功"
    messageType = "success"
ElseIf queryMsg = "copied" Then
    message = "订单已成功复制，新订单号已生成"
    messageType = "success"
End If

' ========== 获取筛选参数 ==========
Dim filterStatus, filterCategory, filterStartDate, filterEndDate, filterKeyword
filterStatus = Trim(Request.QueryString("status"))
filterCategory = Trim(Request.QueryString("category"))
filterStartDate = Trim(Request.QueryString("start_date"))
filterEndDate = Trim(Request.QueryString("end_date"))
filterKeyword = Trim(Request.QueryString("keyword"))

' V8 新增：OrderType 筛选
Dim filterOrderType
filterOrderType = Trim(Request.QueryString("order_type"))

' ========== 获取分页参数 ==========
Dim page, pageSize, offset
page = SafeNum(Request.QueryString("page"))
If page < 1 Then page = 1
pageSize = 20
offset = (page - 1) * pageSize

' ========== 构建查询条件 ==========
Dim whereClause
whereClause = "WHERE 1=1"

If filterStatus <> "" Then
    whereClause = whereClause & " AND Status='" & filterStatus & "'"
End If

If filterCategory <> "" Then
    whereClause = whereClause & " AND CategoryCode='" & filterCategory & "'"
End If

If filterStartDate <> "" Then
    whereClause = whereClause & " AND OrderDate>=#" & filterStartDate & "'"
End If

If filterEndDate <> "" Then
    whereClause = whereClause & " AND OrderDate<=#" & filterEndDate & "'"
End If

If filterKeyword <> "" Then
    whereClause = whereClause & " AND (PurchaseNo LIKE '%" & filterKeyword & "%' OR Remarks LIKE '%" & filterKeyword & "%')"
End If

If filterOrderType <> "" Then
    whereClause = whereClause & " AND OrderType='" & filterOrderType & "'"
End If

' ========== 获取总记录数 ==========
Dim totalRecords, totalPages
On Error Resume Next
totalRecords = SafeNum(GetScalar("SELECT COUNT(*) FROM PurchaseOrders " & whereClause))
If Err.Number <> 0 Then
    totalRecords = 0
    Err.Clear
End If
On Error GoTo 0

totalPages = Int((totalRecords + pageSize - 1) / pageSize)
If totalPages < 1 Then totalPages = 1
If page > totalPages Then page = totalPages

' ========== 获取采购订单列表 ==========
Dim rsOrders
Dim listSQL
On Error Resume Next
If offset > 0 Then
    listSQL = "SELECT TOP " & pageSize & " PurchaseID, PurchaseNo, SupplierID, CategoryCode, OrderType, OrderDate, ExpectedDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status, Remarks, CreatedBy, CreatedAt, UpdatedAt FROM PurchaseOrders " & whereClause & " AND PurchaseID NOT IN (SELECT TOP " & offset & " PurchaseID FROM PurchaseOrders " & whereClause & " ORDER BY OrderDate DESC) ORDER BY OrderDate DESC"
Else
    ' 第一页无需排除子查询
    listSQL = "SELECT TOP " & pageSize & " PurchaseID, PurchaseNo, SupplierID, CategoryCode, OrderType, OrderDate, ExpectedDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status, Remarks, CreatedBy, CreatedAt, UpdatedAt FROM PurchaseOrders " & whereClause & " ORDER BY OrderDate DESC"
End If
Set rsOrders = conn.Execute(listSQL)
If Err.Number <> 0 Then
    Set rsOrders = Nothing
    Err.Clear
End If
On Error GoTo 0

' ========== 获取供应商列表 ==========
Dim rsSuppliers
On Error Resume Next
Set rsSuppliers = conn.Execute("SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive=1 ORDER BY SupplierName")
If Err.Number <> 0 Then
    Set rsSuppliers = Nothing
    Err.Clear
End If
On Error GoTo 0

' ========== 获取基香列表（V9：用于快捷选择） ==========
Dim rsBaseNotes
On Error Resume Next
Set rsBaseNotes = conn.Execute("SELECT BaseNoteID, BaseNoteName, CAST(ISNULL(UnitPrice,0) AS FLOAT) as UnitPrice, Description FROM BaseNotes WHERE IsActive=1 ORDER BY BaseNoteName")
If Err.Number <> 0 Then
    Set rsBaseNotes = Nothing
    Err.Clear
End If
On Error GoTo 0

' ========== 获取编辑的订单信息 ==========
Dim editMode, editOrderID, editOrderData
editMode = False
editOrderID = SafeNum(Request.QueryString("edit"))

If editOrderID > 0 Then
    On Error Resume Next
    Set rsEdit = conn.Execute("SELECT PurchaseID, PurchaseNo, SupplierID, CategoryCode, OrderType, OrderDate, ExpectedDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status, Remarks, CreatedBy, CreatedAt, UpdatedAt FROM PurchaseOrders WHERE PurchaseID=" & editOrderID)
    If Err.Number = 0 Then
        If Not rsEdit.EOF Then
            ' 检查状态是否为Draft
            If CStr(rsEdit("Status")) = "Draft" Then
                editMode = True
                Set editOrderData = rsEdit
            Else
                rsEdit.Close
                Set rsEdit = Nothing
            End If
        Else
            rsEdit.Close
            Set rsEdit = Nothing
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0
End If

' ========== 获取编辑订单的明细 ==========
Dim rsEditDetails
If editMode Then
    On Error Resume Next
    Set rsEditDetails = conn.Execute("SELECT DetailID, PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, CAST(ISNULL(UnitPrice,0) AS FLOAT) as UnitPrice, CAST(ISNULL(TotalPrice,0) AS FLOAT) as TotalPrice, ReceivedQty FROM PurchaseOrderDetails WHERE PurchaseID=" & editOrderID & " ORDER BY DetailID")
    If Err.Number <> 0 Then
        Set rsEditDetails = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End If

' ========== 获取查看的订单信息 ==========
Dim viewMode, viewOrderID, viewOrderData
viewMode = False
viewOrderID = SafeNum(Request.QueryString("view"))

If viewOrderID > 0 Then
    On Error Resume Next
    Set rsView = conn.Execute("SELECT PurchaseID, PurchaseNo, SupplierID, CategoryCode, OrderType, OrderDate, ExpectedDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status, Remarks, CreatedBy, CreatedAt, UpdatedAt FROM PurchaseOrders WHERE PurchaseID=" & viewOrderID)
    If Err.Number = 0 Then
        If Not rsView.EOF Then
            viewMode = True
            Set viewOrderData = rsView
        Else
            rsView.Close
            Set rsView = Nothing
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0
End If

' ========== 获取查看订单的明细 ==========
Dim rsViewDetails
If viewMode Then
    On Error Resume Next
    Set rsViewDetails = conn.Execute("SELECT DetailID, PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, CAST(ISNULL(UnitPrice,0) AS FLOAT) as UnitPrice, CAST(ISNULL(TotalPrice,0) AS FLOAT) as TotalPrice, ReceivedQty FROM PurchaseOrderDetails WHERE PurchaseID=" & viewOrderID & " ORDER BY DetailID")
    If Err.Number <> 0 Then
        Set rsViewDetails = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End If

' ========== V12: 获取订单状态变更时间线 ==========
Dim rsStatusLog
If viewMode Then
    On Error Resume Next
    Set rsStatusLog = conn.Execute("SELECT LogID, FromStatus, ToStatus, ChangedBy, ChangedAt, Remarks FROM PurchaseOrderStatusLog WHERE PurchaseID=" & viewOrderID & " ORDER BY ChangedAt ASC")
    If Err.Number <> 0 Then
        Set rsStatusLog = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>采购订单管理 - 采购管理中心</title>
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
        
        /* 消息提示 */
        .message {
            padding: 12px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
        }
        .message.success {
            background: rgba(76,175,80,0.15);
            border: 1px solid rgba(76,175,80,0.3);
            color: #4CAF50;
        }
        .message.error {
            background: rgba(244,67,54,0.15);
            border: 1px solid rgba(244,67,54,0.3);
            color: #F44336;
        }
        .message i {
            margin-right: 10px;
            font-size: 18px;
        }
        
        /* 筛选栏 */
        .filter-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .filter-row {
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            align-items: flex-end;
        }
        .filter-group {
            display: flex;
            flex-direction: column;
        }
        .filter-group label {
            font-size: 12px;
            color: #888;
            margin-bottom: 5px;
            text-transform: uppercase;
        }
        .filter-group select,
        .filter-group input {
            padding: 8px 12px;
            border-radius: 6px;
            border: 1px solid #3a3a3a;
            background: #252538;
            color: #e0e0e0;
            font-size: 14px;
            min-width: 120px;
        }
        /* 表格样式 */
        .data-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .data-table {
            width: 100%;
            border-collapse: collapse;
        }
        .data-table th,
        .data-table td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }
        .data-table th {
            font-size: 12px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 600;
        }
        .data-table td {
            font-size: 14px;
            color: #e0e0e0;
        }
        .data-table tr:hover td {
            background: rgba(255,255,255,0.02);
        }
        
        /* 状态标签 */
        .status-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 500;
        }
        .status-draft { background: rgba(158,158,158,0.2); color: #9E9E9E; }
        .status-submitted { background: rgba(255,152,0,0.2); color: #FF9800; }
        .status-approved { background: rgba(33,150,243,0.2); color: #2196F3; }
        .status-ordered { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .status-partial { background: rgba(255,193,7,0.2); color: #FFC107; }
        .status-received { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-completed { background: rgba(0,150,136,0.2); color: #009688; }
        .status-rejected { background: rgba(244,67,54,0.2); color: #F44336; }
        
        /* 采购类型标签 */
        .type-badge {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 10px;
            font-size: 11px;
            font-weight: 500;
        }
        .type-badge.type-raw { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .type-badge.type-packaging { background: rgba(33,150,243,0.2); color: #2196F3; }
        .type-badge.type-bottle { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .type-badge.type-printing { background: rgba(0,188,212,0.2); color: #00BCD4; }
        .type-badge.type-sprayhead { background: rgba(255,87,34,0.2); color: #FF5722; }
        
        /* 分页 */
        .pagination {
            display: flex;
            justify-content: center;
            gap: 5px;
            margin-top: 20px;
        }
        .pagination a,
        .pagination span {
            padding: 8px 12px;
            border-radius: 6px;
            text-decoration: none;
            font-size: 14px;
        }
        .pagination a {
            background: #252538;
            color: #e0e0e0;
        }
        .pagination a:hover {
            background: #3a3a3a;
        }
        .pagination .current {
            background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%);
            color: white;
        }
        
        /* 表单区域 */
        .form-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .form-section h3 {
            margin-top: 0;
            margin-bottom: 20px;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .form-section h3 i {
            color: #FF9800;
        }
        .form-row {
            display: flex;
            gap: 20px;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }
        .form-group {
            flex: 1;
            min-width: 200px;
        }
        .form-group label {
            display: block;
            font-size: 13px;
            color: #888;
            margin-bottom: 5px;
        }
        .form-group input,
        .form-group select,
        .form-group textarea {
            width: 100%;
            padding: 10px 12px;
            border-radius: 6px;
            border: 1px solid #3a3a3a;
            background: #252538;
            color: #e0e0e0;
            font-size: 14px;
            box-sizing: border-box;
        }
        .form-group input:focus,
        .form-group select:focus,
        .form-group textarea:focus {
            outline: none;
            border-color: #FF9800;
        }
        
        /* 明细表格 */
        .details-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }
        .details-table th,
        .details-table td {
            padding: 10px;
            border: 1px solid #3a3a3a;
            text-align: left;
        }
        .details-table th {
            background: #252538;
            font-size: 12px;
            color: #888;
        }
        .details-table input {
            width: 100%;
            padding: 6px 8px;
            border: 1px solid #3a3a3a;
            background: #1e1e32;
            color: #e0e0e0;
            border-radius: 4px;
            box-sizing: border-box;
        }
        .details-table .num-input {
            text-align: right;
        }
        .details-table .row-total {
            text-align: right;
            font-weight: 500;
        }
        
        /* 查看详情 */
        .view-section {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 20px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .view-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 20px;
            padding-bottom: 20px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }
        .view-info-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
            margin-bottom: 20px;
        }
        .view-info-item {
            background: rgba(255,255,255,0.02);
            padding: 12px 15px;
            border-radius: 8px;
        }
        .view-info-item label {
            font-size: 11px;
            color: #888;
            display: block;
            margin-bottom: 4px;
        }
        .view-info-item value {
            font-size: 14px;
            color: #fff;
            font-weight: 500;
        }
        
        /* 操作按钮组 */
        .action-btns {
            display: flex;
            gap: 8px;
        }
        
        /* 空数据 */
        .empty-row {
            text-align: center;
            color: #666;
            padding: 40px;
        }
        
        /* 响应式 */
        @media (max-width: 768px) {
            .filter-row {
                flex-direction: column;
            }
            .filter-group {
                width: 100%;
            }
            .view-info-grid {
                grid-template-columns: 1fr;
            }
        }
        /* V9: 基香选择模态框 */
        .modal-overlay {
            display: none;
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0,0,0,0.6);
            z-index: 9999;
            justify-content: center;
            align-items: center;
        }
        .modal-overlay.active { display: flex; }
        .modal-dialog {
            background: #2d2d44;
            border-radius: 12px;
            border: 1px solid rgba(255,255,255,0.1);
            width: 650px;
            max-height: 80vh;
            display: flex;
            flex-direction: column;
        }
        .modal-header {
            padding: 16px 20px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .modal-header h4 { margin: 0; color: #fff; }
        .modal-close {
            background: none; border: none; color: #888; font-size: 20px;
            cursor: pointer; padding: 4px 8px;
        }
        .modal-close:hover { color: #F44336; }
        .modal-body {
            padding: 16px 20px;
            overflow-y: auto;
            flex: 1;
        }
        .modal-search {
            width: 100%;
            padding: 8px 12px;
            border-radius: 6px;
            border: 1px solid #3a3a3a;
            background: #252538;
            color: #e0e0e0;
            font-size: 14px;
            margin-bottom: 12px;
            box-sizing: border-box;
        }
        .base-note-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 8px;
        }
        .base-note-card {
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.06);
            border-radius: 8px;
            padding: 12px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .base-note-card:hover {
            background: rgba(255,152,0,0.1);
            border-color: rgba(255,152,0,0.3);
        }
        .base-note-card .bn-name {
            color: #FF9800;
            font-weight: 500;
            font-size: 13px;
            margin-bottom: 4px;
        }
        .base-note-card .bn-desc {
            color: #888;
            font-size: 11px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .base-note-card .bn-price {
            color: #4CAF50;
            font-size: 11px;
            margin-top: 4px;
        }
        .base-note-card .bn-price .no-price {
            color: #F44336;
            font-size: 10px;
        }
        .btn-select-base {
            background: rgba(255,152,0,0.15);
            color: #FF9800;
            border: 1px dashed rgba(255,152,0,0.3);
            padding: 6px 14px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 12px;
            transition: all 0.2s;
            white-space: nowrap;
        }
        .btn-select-base:hover {
            background: rgba(255,152,0,0.25);
            border-color: #FF9800;
        }
    
        /* V12: 操作时间线 */
        .timeline {
            position: relative;
            padding-left: 30px;
        }
        .timeline::before {
            content: '';
            position: absolute;
            left: 10px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: rgba(255,255,255,0.08);
        }
        .timeline-item {
            position: relative;
            margin-bottom: 20px;
        }
        .timeline-item:last-child {
            margin-bottom: 0;
        }
        .timeline-dot {
            position: absolute;
            left: -24px;
            top: 4px;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: rgba(255,255,255,0.15);
            border: 2px solid rgba(255,255,255,0.1);
        }
        .timeline-dot.active {
            background: #FF9800;
            border-color: #FF9800;
            box-shadow: 0 0 8px rgba(255,152,0,0.4);
        }
        .timeline-content {
            background: rgba(255,255,255,0.02);
            padding: 12px 15px;
            border-radius: 8px;
            border: 1px solid rgba(255,255,255,0.04);
        }
        .timeline-header {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 6px;
        }
        .timeline-time {
            font-size: 11px;
            color: #888;
        }
        .timeline-desc {
            font-size: 13px;
            color: #ccc;
        }
        .timeline-actor {
            font-size: 11px;
            color: #888;
            margin-top: 4px;
        }
</style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-file-invoice"></i> 采购订单管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>采购订单</span>
            </div>
        </div>
        
        <% If message <> "" Then %>
        <div class="message <%= messageType %>">
            <i class="fas fa-<%= IIf(messageType="success", "check-circle", "exclamation-circle") %>"></i>
            <%= message %>
        </div>
        <% End If %>
        
        <% 
        ' 调试信息：显示数据库错误详情（仅管理员可见）
        Dim dbgErr
        dbgErr = Session("LastDBError")
        If dbgErr <> "" Then
        %>
        <div class="message error">
            <i class="fas fa-bug"></i>
            <span style="font-size:12px;">调试: <%= Server.HTMLEncode(dbgErr) %></span>
        </div>
        <% 
            Session("LastDBError") = ""
        End If
        %>
        
        <% If viewMode Then %>
        ' ========== 查看订单详情 ==========
        <div class="view-section">
            <div class="view-header">
                <div>
                    <h3 style="margin:0 0 10px 0;">
                        <i class="fas fa-eye" style="color:#FF9800;"></i> 
                        订单详情：<%= Server.HTMLEncode(CStr(viewOrderData("PurchaseNo"))) %>
                    </h3>
                    <span class="status-badge <%= GetStatusClass(CStr(viewOrderData("Status"))) %>">
                        <%= GetStatusName(CStr(viewOrderData("Status"))) %>
                    </span>
                </div>
                <div>
                    <a href="purchase_orders.asp" class="btn btn-secondary">
                        <i class="fas fa-arrow-left"></i> 返回列表
                    </a>
                    <button type="button" class="btn btn-secondary" style="margin-left:8px;" onclick="copyOrderFromView(<%= viewOrderID %>)">
                        <i class="fas fa-copy"></i> 复制订单
                    </button>
                </div>
            </div>
            
            <div class="view-info-grid">
                <div class="view-info-item">
                    <label>供应商</label>
                    <value><%= Server.HTMLEncode(GetSupplierName(viewOrderData("SupplierID"))) %></value>
                </div>
                <div class="view-info-item">
                    <label>采购分类</label>
                    <value><%= GetCategoryName(CStr(viewOrderData("CategoryCode"))) %></value>
                </div>
                <div class="view-info-item">
                    <label>订单日期</label>
                    <value><% If IsDate(viewOrderData("OrderDate")) Then Response.Write FormatDateTime(viewOrderData("OrderDate"), 2) End If %></value>
                </div>
                <div class="view-info-item">
                    <label>期望交期</label>
                    <value>
                        <% If IsNull(viewOrderData("ExpectedDate")) Then %>
                            未设置
                        <% Else %>
                            <% If IsDate(viewOrderData("ExpectedDate")) Then Response.Write FormatDateTime(viewOrderData("ExpectedDate"), 2) End If %>
                        <% End If %>
                    </value>
                </div>
                <div class="view-info-item">
                    <label>订单金额</label>
                    <value style="color:#FF9800;font-size:18px;">¥<%= FormatNumber(SafeNum(viewOrderData("TotalAmount")), 2) %></value>
                </div>
                <div class="view-info-item">
                    <label>创建人</label>
                    <value><%= SafeNum(viewOrderData("CreatedBy")) %></value>
                </div>
            </div>
            
            <% Dim remarksVal : remarksVal = viewOrderData("Remarks") & "" : If remarksVal <> "" Then %>
            <div style="margin-bottom:20px;">
                <label style="font-size:12px;color:#888;display:block;margin-bottom:5px;">备注</label>
                <div style="background:rgba(255,255,255,0.02);padding:12px 15px;border-radius:8px;">
                    <%= Server.HTMLEncode(remarksVal) %>
                </div>
            </div>
            <% End If %>
            
            <h4 style="margin:20px 0 15px 0;color:#fff;"><i class="fas fa-list"></i> 采购明细</h4>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>物料名称</th>
                        <th>物料编码</th>
                        <th>规格</th>
                        <th>单位</th>
                        <th style="text-align:right;">数量</th>
                        <th style="text-align:right;">单价</th>
                        <th style="text-align:right;">小计</th>
                        <th style="text-align:right;">已收货</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    If rsViewDetails Is Nothing Then
                    %>
                    <tr>
                        <td colspan="7" class="empty-row">暂无明细</td>
                    </tr>
                    <% 
                    Else
                        If rsViewDetails.EOF Then
                    %>
                    <tr>
                        <td colspan="7" class="empty-row">暂无明细</td>
                    </tr>
                    <% 
                        Else
                            Do While Not rsViewDetails.EOF
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(CStr(rsViewDetails("ItemName"))) %></td>
                        <td>
                            <% If IsNull(rsViewDetails("ItemCode")) Then %>
                                -
                            <% Else %>
                                <%= Server.HTMLEncode(CStr(rsViewDetails("ItemCode"))) %>
                            <% End If %>
                        </td>
                        <td>
                            <% If IsNull(rsViewDetails("Specification")) Then %>
                                -
                            <% Else %>
                                <%= Server.HTMLEncode(CStr(rsViewDetails("Specification"))) %>
                            <% End If %>
                        </td>
                        <td>
                            <% If IsNull(rsViewDetails("Unit")) Then %>
                                -
                            <% Else %>
                                <%= Server.HTMLEncode(CStr(rsViewDetails("Unit"))) %>
                            <% End If %>
                        </td>
                        <td style="text-align:right;"><%= SafeNum(rsViewDetails("Quantity")) %></td>
                        <td style="text-align:right;">¥<%= FormatNumber(SafeNum(rsViewDetails("UnitPrice")), 2) %></td>
                        <td style="text-align:right;">¥<%= FormatNumber(SafeNum(rsViewDetails("TotalPrice")), 2) %></td>
                        <td style="text-align:right;"><%= SafeNum(rsViewDetails("ReceivedQty")) %></td>
                    </tr>
                    <% 
                                rsViewDetails.MoveNext
                            Loop
                            rsViewDetails.Close
                            Set rsViewDetails = Nothing
                        End If
                    End If
                    %>
                </tbody>
            </table>
            
            
            <% ' ========== V12: 操作时间线 ========== %>
            <h4 style="margin:20px 0 15px 0;color:#fff;"><i class="fas fa-history"></i> 操作时间线</h4>
            <div class="timeline">
                <%
                If Not rsStatusLog Is Nothing Then
                    If Not rsStatusLog.EOF Then
                        Dim tlIdx : tlIdx = 0
                        Do While Not rsStatusLog.EOF
                            Dim tlFromStatus : tlFromStatus = CStr(rsStatusLog("FromStatus") & "")
                            Dim tlToStatus : tlToStatus = CStr(rsStatusLog("ToStatus"))
                            Dim tlLogTime : tlLogTime = rsStatusLog("ChangedAt")
                            Dim tlChanger : tlChanger = CStr(rsStatusLog("ChangedBy") & "")
                            Dim tlRemark : tlRemark = CStr(rsStatusLog("Remarks") & "")
                %>
                <div class="timeline-item">
                    <div class="timeline-dot <%= IIf(tlIdx=0, "active", "") %>"></div>
                    <div class="timeline-content">
                        <div class="timeline-header">
                            <span class="status-badge <%= GetStatusClass(tlToStatus) %>"><%= GetStatusName(tlToStatus) %></span>
                            <span class="timeline-time"><% If IsDate(tlLogTime) Then Response.Write FormatDateTime(tlLogTime, 2) & " " & FormatDateTime(tlLogTime, 4) End If %></span>
                        </div>
                        <div class="timeline-desc">
                            <% If tlFromStatus <> "" Then %>
                                <%= GetStatusName(tlFromStatus) %> &rarr; <%= GetStatusName(tlToStatus) %>
                            <% Else %>
                                创建订单（初始状态：<%= GetStatusName(tlToStatus) %>）
                            <% End If %>
                            <% If tlRemark <> "" Then %> &mdash; <em><%= Server.HTMLEncode(tlRemark) %></em><% End If %>
                        </div>
                        <% If tlChanger <> "" Then %>
                        <div class="timeline-actor"><i class="fas fa-user"></i> <%= Server.HTMLEncode(tlChanger) %></div>
                        <% End If %>
                    </div>
                </div>
                <%
                            tlIdx = tlIdx + 1
                            rsStatusLog.MoveNext
                        Loop
                        rsStatusLog.Close
                        Set rsStatusLog = Nothing
                    Else
                        rsStatusLog.Close
                        Set rsStatusLog = Nothing
                %>
                <div class="timeline-item">
                    <div class="timeline-dot"></div>
                    <div class="timeline-content">
                        <div class="timeline-desc" style="color:#666;">暂无状态变更记录</div>
                    </div>
                </div>
                <%
                    End If
                Else
                %>
                <div class="timeline-item">
                    <div class="timeline-dot"></div>
                    <div class="timeline-content">
                        <div class="timeline-desc" style="color:#666;">暂无状态变更记录</div>
                    </div>
                </div>
                <% End If %>
            </div>
<% ' ========== 状态操作按钮 ========== %>
            <div style="margin-top:25px;padding-top:20px;border-top:1px solid rgba(255,255,255,0.05);">
                <h4 style="margin:0 0 15px 0;color:#fff;"><i class="fas fa-exchange-alt"></i> 状态操作</h4>
                <div class="action-btns">
                    <% 
                    Dim viewStatus
                    viewStatus = CStr(viewOrderData("Status"))
                    
                    If viewStatus = "Draft" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Submitted">
                        <button type="submit" class="btn btn-success" onclick="return confirm('确定提交审批吗？');">
                            <i class="fas fa-paper-plane"></i> 提交审批
                        </button>
                    </form>
                    <a href="purchase_orders.asp?edit=<%= viewOrderID %>" class="btn btn-primary">
                        <i class="fas fa-edit"></i> 编辑
                    </a>
                    <% 
                    ElseIf viewStatus = "Submitted" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="FinanceApproved">
                        <button type="submit" class="btn btn-primary" onclick="return confirm('确定通过财务审批吗？');">
                            <i class="fas fa-check"></i> 财务审批
                        </button>
                    </form>
                    <% 
                    ElseIf viewStatus = "FinanceApproved" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Ordered">
                        <button type="submit" class="btn btn-primary" onclick="return confirm('确定标记为已下单吗？');">
                            <i class="fas fa-shopping-cart"></i> 确认下单
                        </button>
                    </form>
                    <% 
                    ElseIf viewStatus = "Ordered" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Received">
                        <button type="submit" class="btn btn-success" onclick="return confirm('确定标记为已收货吗？');">
                            <i class="fas fa-box"></i> 确认收货
                        </button>
                    </form>
                    <% 
                    ElseIf viewStatus = "Received" And isManager Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Completed">
                        <button type="submit" class="btn btn-success" onclick="return confirm('确定完成订单吗？');">
                            <i class="fas fa-check-circle"></i> 完成订单
                        </button>
                    </form>
                    <% End If %>
                </div>
            </div>
        </div>
        <% 
        viewOrderData.Close
        Set viewOrderData = Nothing
        End If 
        %>
        
        <%
        ' ========== 新建订单预选类型 ==========
        Dim preselectCategory, preselectOrderType
        preselectCategory = ""
        preselectOrderType = ""
        If Request.QueryString("new") = "1" Then
            Dim qsOT : qsOT = Trim(Request.QueryString("order_type"))
            If qsOT <> "" Then
                preselectOrderType = qsOT
                Select Case qsOT
                    Case "RawMaterial" : preselectCategory = "RAW"
                    Case "Packaging" : preselectCategory = "PACK"
                    Case "Bottle" : preselectCategory = "BOTTLE"
                    Case "Printing" : preselectCategory = "PRINTING"
                    Case "SprayHead" : preselectCategory = "SPRAYHEAD"
                End Select
            End If
        End If
        
        ' ========== 预计算采购分类选中状态 ==========
        Dim catSelRAW, catSelBASE, catSelPACK, catSelMARKET, catSelBOTTLE, catSelPRINTING, catSelSPRAYHEAD
        catSelRAW = "" : catSelBASE = "" : catSelPACK = "" : catSelMARKET = "" : catSelBOTTLE = "" : catSelPRINTING = "" : catSelSPRAYHEAD = ""
        If editMode Then
            Dim edCatVal : edCatVal = CStr(editOrderData("CategoryCode"))
            If edCatVal = "RAW" Then catSelRAW = "selected"
            If edCatVal = "BASE" Then catSelBASE = "selected"
            If edCatVal = "PACK" Then catSelPACK = "selected"
            If edCatVal = "MARKET" Then catSelMARKET = "selected"
            If edCatVal = "BOTTLE" Then catSelBOTTLE = "selected"
            If edCatVal = "PRINTING" Then catSelPRINTING = "selected"
            If edCatVal = "SPRAYHEAD" Then catSelSPRAYHEAD = "selected"
        Else
            If preselectCategory = "RAW" Then catSelRAW = "selected"
            If preselectCategory = "BASE" Then catSelBASE = "selected"
            If preselectCategory = "PACK" Then catSelPACK = "selected"
            If preselectCategory = "MARKET" Then catSelMARKET = "selected"
            If preselectCategory = "BOTTLE" Then catSelBOTTLE = "selected"
            If preselectCategory = "PRINTING" Then catSelPRINTING = "selected"
            If preselectCategory = "SPRAYHEAD" Then catSelSPRAYHEAD = "selected"
        End If
        
        ' ========== 预计算采购类型选中状态 ==========
        Dim otSelRAW, otSelPACK, otSelBOTTLE, otSelPRINT, otSelSPRAY
        otSelRAW = "" : otSelPACK = "" : otSelBOTTLE = "" : otSelPRINT = "" : otSelSPRAY = ""
        If editMode Then
            Dim edOTVal : edOTVal = CStr(editOrderData("OrderType"))
            If edOTVal = "RawMaterial" Then otSelRAW = "selected"
            If edOTVal = "Packaging" Then otSelPACK = "selected"
            If edOTVal = "Bottle" Then otSelBOTTLE = "selected"
            If edOTVal = "Printing" Then otSelPRINT = "selected"
            If edOTVal = "SprayHead" Then otSelSPRAY = "selected"
        Else
            If preselectOrderType = "RawMaterial" Or preselectOrderType = "" Then otSelRAW = "selected"
            If preselectOrderType = "Packaging" Then otSelPACK = "selected"
            If preselectOrderType = "Bottle" Then otSelBOTTLE = "selected"
            If preselectOrderType = "Printing" Then otSelPRINT = "selected"
            If preselectOrderType = "SprayHead" Then otSelSPRAY = "selected"
        End If
        
        ' ========== 基香按钮可见性 ==========
        Dim showBaseNoteSection
        showBaseNoteSection = False
        If editMode Then
            If CStr(editOrderData("CategoryCode")) = "BASE" Or CStr(editOrderData("OrderType")) = "RawMaterial" Then
                showBaseNoteSection = True
            End If
        Else
            If preselectCategory = "BASE" Or preselectOrderType = "RawMaterial" Or preselectOrderType = "" Then
                showBaseNoteSection = True
            End If
        End If
        %>
        
        <% If editMode Or Request.QueryString("new") = "1" Then %>
        ' ========== 创建/编辑表单 ==========
        <div class="form-section">
            <h3>
                <% If editMode Then %>
                <i class="fas fa-edit"></i> 编辑采购订单
                <% Else %>
                <i class="fas fa-plus-circle"></i> 新建采购订单
                <% End If %>
            </h3>
            
            <form method="post" id="purchaseForm">
                <input type="hidden" name="action" value="<%= IIf(editMode, "update", "create") %>">
                <% If editMode Then %>
                <input type="hidden" name="purchase_id" value="<%= editOrderID %>">
                <% End If %>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>采购单号</label>
                        <% If editMode Then %>
                        <input type="text" value="<%= Server.HTMLEncode(CStr(editOrderData("PurchaseNo"))) %>" readonly style="background:#1a1a2e;">
                        <% Else %>
                        <input type="text" value="<%= GeneratePurchaseNo() %>（自动生成）" readonly style="background:#1a1a2e;">
                        <% End If %>
                    </div>
                    <div class="form-group">
                        <label>供应商 <span style="color:#F44336;">*</span></label>
                        <select name="supplier_id" required>
                            <option value="">请选择供应商</option>
                            <% 
                            If Not rsSuppliers Is Nothing Then
                                Do While Not rsSuppliers.EOF
                                    Dim selected
                                    selected = ""
                                    If editMode Then
                                        If SafeNum(editOrderData("SupplierID")) = SafeNum(rsSuppliers("SupplierID")) Then
                                            selected = "selected"
                                        End If
                                    End If
                            %>
                            <option value="<%= rsSuppliers("SupplierID") %>" <%= selected %>><%= Server.HTMLEncode(CStr(IIf(IsNull(rsSuppliers("SupplierName")), "", rsSuppliers("SupplierName")))) %></option>
                            <% 
                                    rsSuppliers.MoveNext
                                Loop
                                rsSuppliers.Close
                                Set rsSuppliers = Nothing
                            End If
                            %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>采购分类 <span style="color:#F44336;">*</span></label>
                        <select name="category_code" required onchange="updateOrderTypeByCategory()">
                            <option value="">请选择分类</option>
                            <option value="RAW" <%= catSelRAW %>>原材料</option>
                            <option value="BASE" <%= catSelBASE %>>基香原料</option>
                            <option value="PACK" <%= catSelPACK %>>包装材料</option>
                            <option value="MARKET" <%= catSelMARKET %>>营销物料</option>
                            <option value="BOTTLE" <%= catSelBOTTLE %>>瓶子包装</option>
                            <option value="PRINTING" <%= catSelPRINTING %>>印刷品</option>
                            <option value="SPRAYHEAD" <%= catSelSPRAYHEAD %>>喷头配件</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>采购类型 <span style="color:#F44336;">*</span></label>
                        <select name="order_type" required onchange="updateCategoryByOrderType()">
                            <option value="">请选择类型</option>
                            <option value="RawMaterial" <%= otSelRAW %>>原料采购</option>
                            <option value="Packaging" <%= otSelPACK %>>包装物采购</option>
                            <option value="Bottle" <%= otSelBOTTLE %>>瓶子采购</option>
                            <option value="Printing" <%= otSelPRINT %>>印刷品采购</option>
                            <option value="SprayHead" <%= otSelSPRAY %>>喷头采购</option>
                        </select>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>期望交期</label>
                        <input type="date" name="expected_date" 
                            <% If editMode Then %>
                            <% If Not IsNull(editOrderData("ExpectedDate")) And IsDate(editOrderData("ExpectedDate")) Then %>
                            value="<% If IsDate(editOrderData("ExpectedDate")) Then Response.Write FormatDateTime(editOrderData("ExpectedDate"), 2) End If %>"
                            <% End If %>
                            <% End If %>>
                    </div>
                    <div class="form-group" style="flex:2;">
                        <label>备注</label>
                        <input type="text" name="remarks" maxlength="500" placeholder="可选填"
                            <% If editMode Then %>
                            <% If Not IsNull(editOrderData("Remarks")) Then %>
                            value="<%= Server.HTMLEncode(CStr(editOrderData("Remarks"))) %>"
                            <% End If %>
                            <% End If %>>
                    </div>
                </div>
                
                <h4 style="margin:25px 0 15px 0;color:#fff;"><i class="fas fa-list"></i> 采购明细</h4>
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
                    <div id="baseNoteSection" style="display:<% If showBaseNoteSection Then Response.Write "flex" Else Response.Write "none" End If %>;align-items:center;gap:10px;">
                        <button type="button" class="btn-select-base" onclick="openBaseNoteModal()">
                            <i class="fas fa-flask"></i> 选择基香原料
                        </button>
                        <span style="font-size:11px;color:#888;">快速选择系统已录入的基香，自动填充物料信息（编码BN-XXX、单价等）</span>
                    </div>
                    <button type="button" class="btn-select-base" onclick="openHistoryModal()" style="margin-left:auto;">
                        <i class="fas fa-history"></i> 历史产品快速选择
                    </button>
                </div>
                <table class="details-table" id="detailsTable">
                    <thead>
                        <tr>
                            <th style="width:25%;">物料名称 <span style="color:#F44336;">*</span></th>
                            <th style="width:12%;">物料编码</th>
                            <th style="width:15%;">规格</th>
                            <th style="width:10%;">单位</th>
                            <th style="width:10%;">数量</th>
                            <th style="width:12%;">单价</th>
                            <th style="width:12%;">小计</th>
                            <th style="width:4%;"></th>
                        </tr>
                    </thead>
                    <tbody id="detailsBody">
                        <% 
                        Dim rowCount
                        rowCount = 0
                        
                        If editMode Then
                            If Not rsEditDetails Is Nothing Then
                                Do While Not rsEditDetails.EOF
                                    rowCount = rowCount + 1
                        %>
                        <tr>
                            <td><input type="text" name="item_name_<%= rowCount %>" value="<%= Server.HTMLEncode(CStr(rsEditDetails("ItemName"))) %>" required></td>
                            <td><input type="text" name="item_code_<%= rowCount %>" value="<% If Not IsNull(rsEditDetails("ItemCode")) Then Response.Write Server.HTMLEncode(CStr(rsEditDetails("ItemCode"))) %>"></td>
                            <td><input type="text" name="spec_<%= rowCount %>" value="<% If Not IsNull(rsEditDetails("Specification")) Then Response.Write Server.HTMLEncode(CStr(rsEditDetails("Specification"))) %>"></td>
                            <td><input type="text" name="unit_<%= rowCount %>" value="<% If Not IsNull(rsEditDetails("Unit")) Then Response.Write Server.HTMLEncode(CStr(rsEditDetails("Unit"))) %>"></td>
                            <td><input type="number" name="qty_<%= rowCount %>" class="num-input qty" value="<%= SafeNum(rsEditDetails("Quantity")) %>" min="0" step="0.01" onchange="calculateRow(this)"></td>
                            <td><input type="number" name="price_<%= rowCount %>" class="num-input price" value="<%= SafeNum(rsEditDetails("UnitPrice")) %>" min="0" step="0.01" onchange="calculateRow(this)"></td>
                            <td class="row-total">¥<%= FormatNumber(SafeNum(rsEditDetails("TotalPrice")), 2) %></td>
                            <td><button type="button" class="btn btn-danger btn-sm" onclick="removeRow(this)"><i class="fas fa-trash"></i></button></td>
                        </tr>
                        <% 
                                    rsEditDetails.MoveNext
                                Loop
                                rsEditDetails.Close
                                Set rsEditDetails = Nothing
                            End If
                        End If
                        %>
                    </tbody>
                    <tfoot>
                        <tr>
                            <td colspan="8" style="text-align:center;">
                                <button type="button" class="btn btn-secondary" onclick="addRow()">
                                    <i class="fas fa-plus"></i> 添加明细行
                                </button>
                            </td>
                        </tr>
                    </tfoot>
                </table>
                
                <input type="hidden" name="item_count" id="itemCount" value="<%= rowCount %>">
                
                <div style="margin-top:25px;display:flex;gap:10px;justify-content:flex-end;">
                    <a href="purchase_orders.asp" class="btn btn-secondary">取消</a>
                    <button type="submit" class="btn btn-primary">
                        <i class="fas fa-save"></i> <%= IIf(editMode, "保存修改", "创建订单") %>
                    </button>
                </div>
            </form>
        </div>
        <% 
        If editMode Then
            editOrderData.Close
            Set editOrderData = Nothing
        End If
        End If 
        %>
        
        <% If Not viewMode And Request.QueryString("new") <> "1" And Not editMode Then %>
        ' ========== 筛选栏 ==========
        <div class="filter-section">
            <form method="get" class="filter-row">
                <div class="filter-group">
                    <label>状态</label>
                    <select name="status">
                        <option value="">全部</option>
                        <option value="Draft" <% If filterStatus="Draft" Then Response.Write "selected" %>>草稿</option>
                        <option value="Submitted" <% If filterStatus="Submitted" Then Response.Write "selected" %>>待审批</option>
                        <option value="FinanceApproved" <% If filterStatus="FinanceApproved" Then Response.Write "selected" %>>已审批</option>
                        <option value="Ordered" <% If filterStatus="Ordered" Then Response.Write "selected" %>>已下单</option>
                        <option value="PartialReceived" <% If filterStatus="PartialReceived" Then Response.Write "selected" %>>部分收货</option>
                        <option value="Received" <% If filterStatus="Received" Then Response.Write "selected" %>>已收货</option>
                        <option value="Completed" <% If filterStatus="Completed" Then Response.Write "selected" %>>已完成</option>
                        <option value="Rejected" <% If filterStatus="Rejected" Then Response.Write "selected" %>>已拒绝</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>采购分类</label>
                    <select name="category">
                        <option value="">全部</option>
                        <option value="RAW" <% If filterCategory="RAW" Then Response.Write "selected" %>>原材料</option>
                        <option value="BASE" <% If filterCategory="BASE" Then Response.Write "selected" %>>基香原料</option>
                        <option value="PACK" <% If filterCategory="PACK" Then Response.Write "selected" %>>包装材料</option>
                        <option value="MARKET" <% If filterCategory="MARKET" Then Response.Write "selected" %>>营销物料</option>
                        <option value="BOTTLE" <% If filterCategory="BOTTLE" Then Response.Write "selected" %>>瓶子包装</option>
                        <option value="PRINTING" <% If filterCategory="PRINTING" Then Response.Write "selected" %>>印刷品</option>
                        <option value="SPRAYHEAD" <% If filterCategory="SPRAYHEAD" Then Response.Write "selected" %>>喷头配件</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>采购类型</label>
                    <select name="order_type">
                        <option value="">全部</option>
                        <option value="RawMaterial" <% If filterOrderType="RawMaterial" Then Response.Write "selected" %>>原料采购</option>
                        <option value="Packaging" <% If filterOrderType="Packaging" Then Response.Write "selected" %>>包装物采购</option>
                        <option value="Bottle" <% If filterOrderType="Bottle" Then Response.Write "selected" %>>瓶子采购</option>
                        <option value="Printing" <% If filterOrderType="Printing" Then Response.Write "selected" %>>印刷品采购</option>
                        <option value="SprayHead" <% If filterOrderType="SprayHead" Then Response.Write "selected" %>>喷头采购</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>开始日期</label>
                    <input type="date" name="start_date" value="<%= filterStartDate %>">
                </div>
                <div class="filter-group">
                    <label>结束日期</label>
                    <input type="date" name="end_date" value="<%= filterEndDate %>">
                </div>
                <div class="filter-group">
                    <label>搜索</label>
                    <input type="text" name="keyword" value="<%= Server.HTMLEncode(filterKeyword) %>" placeholder="单号/备注">
                </div>
                <div class="filter-group">
                    <button type="submit" class="btn btn-secondary">
                        <i class="fas fa-filter"></i> 筛选
                    </button>
                </div>
                <div class="filter-group" style="margin-left:auto;">
                    <a href="purchase_orders.asp?new=1<% If filterOrderType <> "" Then %>&order_type=<%= filterOrderType %><% End If %>" class="btn btn-primary">
                        <i class="fas fa-plus"></i> 新建订单
                    </a>
                </div>
            </form>
        </div>
        
        ' ========== 订单列表 ==========
        <div class="data-section">
            <!-- V11: 批量操作工具栏 -->
            <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap;">
                <button type="button" class="btn btn-sm" style="background:#2196F3;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;" onclick="batchAction('approve')" title="批量财务审批（仅待审批状态）"><i class="fas fa-check"></i> 批量审批</button>
                <button type="button" class="btn btn-sm" style="background:#FF9800;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;" onclick="batchAction('order')" title="批量确认下单（仅已审批状态）"><i class="fas fa-shopping-cart"></i> 批量下单</button>
                <button type="button" class="btn btn-sm" style="background:#4CAF50;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;" onclick="batchAction('receive')" title="批量确认收货（仅已下单状态）"><i class="fas fa-box"></i> 批量收货</button>
                <span id="batchCount" style="color:#888;font-size:12px;margin-left:10px;"></span>
            </div>
            
            <table class="data-table">
                <thead>
                    <tr>
                        <th style="width:40px;"><input type="checkbox" id="selectAll" onchange="toggleSelectAll(this)" title="全选"></th>
                        <th>采购单号</th>
                        <th>供应商</th>
                        <th>类型</th>
                        <th style="text-align:right;">金额</th>
                        <th>状态</th>
                        <th>订单日期</th>
                        <th style="text-align:center;">操作</th>
                    </tr>
                </thead>
                <tbody>
                    <% 
                    If rsOrders Is Nothing Then
                    %>
                    <tr>
                        <td colspan="8" class="empty-row">
                            <i class="fas fa-inbox" style="font-size: 24px; margin-bottom: 10px; display: block;"></i>
                            暂无采购订单数据
                        </td>
                    </tr>
                    <% 
                    Else
                        If rsOrders.EOF Then
                    %>
                    <tr>
                        <td colspan="8" class="empty-row">
                            <i class="fas fa-inbox" style="font-size: 24px; margin-bottom: 10px; display: block;"></i>
                            暂无符合条件的订单
                        </td>
                    </tr>
                    <% 
                        Else
                            Do While Not rsOrders.EOF
                    %>
                    <tr>
                        <td><input type="checkbox" class="row-check" value="<%= rsOrders("PurchaseID") %>" onchange="updateBatchCount()"></td>
                        <td><%= Server.HTMLEncode(CStr(rsOrders("PurchaseNo"))) %></td>
                        <td><%= Server.HTMLEncode(GetSupplierName(rsOrders("SupplierID"))) %></td>
                        <td>
                            <%
                            Dim oType : oType = CStr(rsOrders("OrderType") & "")
                            If oType = "Packaging" Then
                                Response.Write "<span class='type-badge type-packaging'>包装物</span>"
                            ElseIf oType = "Bottle" Then
                                Response.Write "<span class='type-badge type-bottle'>瓶子</span>"
                            ElseIf oType = "Printing" Then
                                Response.Write "<span class='type-badge type-printing'>印刷品</span>"
                            ElseIf oType = "SprayHead" Then
                                Response.Write "<span class='type-badge type-sprayhead'>喷头</span>"
                            Else
                                Response.Write "<span class='type-badge type-raw'>原料</span>"
                            End If
                            %>
                        </td>
                        <td style="text-align:right;">¥<%= FormatNumber(SafeNum(rsOrders("TotalAmount")), 2) %></td>
                        <td>
                            <span class="status-badge <%= GetStatusClass(CStr(rsOrders("Status"))) %>">
                                <%= GetStatusName(CStr(rsOrders("Status"))) %>
                            </span>
                        </td>
                        <td><%
                            Dim od : od = rsOrders("OrderDate")
                            If Not IsNull(od) And IsDate(od) Then
                                Response.Write FormatDateTime(od, 2)
                            End If
                        %></td>
                        <td style="text-align:center;">
                            <div class="action-btns" style="justify-content:center;">
                                <a href="purchase_orders.asp?view=<%= rsOrders("PurchaseID") %>" class="btn btn-secondary btn-sm" title="查看">
                                    <i class="fas fa-eye"></i>
                                </a>
                                <% If CStr(rsOrders("Status")) = "Draft" Then %>
                                <a href="purchase_orders.asp?edit=<%= rsOrders("PurchaseID") %>" class="btn btn-primary btn-sm" title="编辑">
                                    <i class="fas fa-edit"></i>
                                </a>
                                <% End If %>
                                <% ' V11: 复制订单 %>
                                <form method="post" style="display:inline;" onsubmit="return confirm('确定复制该订单吗？将创建一个新的草稿订单。')">
                                    <input type="hidden" name="action" value="copy">
                                    <input type="hidden" name="purchase_id" value="<%= rsOrders("PurchaseID") %>">
                                    <button type="submit" class="btn btn-sm" style="background:#9C27B0;color:#fff;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;" title="复制订单">
                                        <i class="fas fa-copy"></i>
                                    </button>
                                </form>
                            </div>
                        </td>
                    </tr>
                    <% 
                                rsOrders.MoveNext
                            Loop
                            rsOrders.Close
                            Set rsOrders = Nothing
                        End If
                    End If
                    %>
                </tbody>
            </table>
            
            ' ========== 分页 ==========
            <% If totalPages > 1 Then %>
            <div class="pagination">
                <% If page > 1 Then %>
                <a href="purchase_orders.asp?page=<%= page-1 %>&status=<%= filterStatus %>&category=<%= filterCategory %>&start_date=<%= filterStartDate %>&end_date=<%= filterEndDate %>&keyword=<%= Server.URLEncode(filterKeyword) %>"><i class="fas fa-chevron-left"></i></a>
                <% End If %>
                
                <% 
                Dim p
                For p = 1 To totalPages
                    If p = page Then
                %>
                <span class="current"><%= p %></span>
                <% Else %>
                <a href="purchase_orders.asp?page=<%= p %>&status=<%= filterStatus %>&category=<%= filterCategory %>&start_date=<%= filterStartDate %>&end_date=<%= filterEndDate %>&keyword=<%= Server.URLEncode(filterKeyword) %>"><%= p %></a>
                <% 
                    End If
                Next 
                %>
                
                <% If page < totalPages Then %>
                <a href="purchase_orders.asp?page=<%= page+1 %>&status=<%= filterStatus %>&category=<%= filterCategory %>&start_date=<%= filterStartDate %>&end_date=<%= filterEndDate %>&keyword=<%= Server.URLEncode(filterKeyword) %>"><i class="fas fa-chevron-right"></i></a>
                <% End If %>
            </div>
            <% End If %>
        </div>
        <% End If %>
    </div>
    
    ' V9: 基香选择模态框
    %>
    <div class="modal-overlay" id="baseNoteModal">
        <div class="modal-dialog">
            <div class="modal-header">
                <h4><i class="fas fa-flask" style="color:#FF9800;"></i> 选择基香原料</h4>
                <button type="button" class="modal-close" onclick="closeBaseNoteModal()">&times;</button>
            </div>
            <div class="modal-body">
                <input type="text" class="modal-search" id="baseNoteSearch" placeholder="搜索基香名称..." oninput="filterBaseNotes()">
                <div class="base-note-grid" id="baseNoteGrid">
                    <div style="grid-column:1/-1;text-align:center;color:#666;padding:20px;">
                        加载中...
                    </div>
                </div>
            </div>
        </div>
    </div>
    <%
    %>
    <div class="modal-overlay" id="historyModal">
        <div class="modal-dialog" style="width:750px;">
            <div class="modal-header">
                <h4><i class="fas fa-history" style="color:#FF9800;"></i> 历史采购产品快速选择</h4>
                <button type="button" class="modal-close" onclick="closeHistoryModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div style="display:flex;gap:10px;margin-bottom:12px;">
                    <input type="text" class="modal-search" id="historySearch" placeholder="搜索产品名称或编码..." oninput="searchHistory()" style="flex:2;">
                    <select id="historySupplierFilter" onchange="searchHistory()" style="flex:1;padding:8px 12px;border-radius:6px;border:1px solid #3a3a3a;background:#252538;color:#e0e0e0;font-size:14px;">
                        <option value="">全部供应商</option>
                    </select>
                </div>
                <div style="max-height:400px;overflow-y:auto;" id="historyGrid">
                    <div style="text-align:center;color:#666;padding:20px;">加载中...</div>
                </div>
            </div>
        </div>
    </div>
    <%
    %>
    <script>
        // V9: 基香数据（从服务端渲染）
        var baseNotesData = [
        <%
        If Not rsBaseNotes Is Nothing Then
            Dim bnFirst : bnFirst = True
            Do While Not rsBaseNotes.EOF
                If Not bnFirst Then Response.Write ","
                bnFirst = False
                Dim bnDesc : bnDesc = CStr(rsBaseNotes("Description") & "")
                Dim bnName : bnName = CStr(rsBaseNotes("BaseNoteName") & "")
                Dim bnNameJS : bnNameJS = Replace(bnName, Chr(34), "\" & Chr(34))
                Dim bnDescJS : bnDescJS = Replace(bnDesc, Chr(34), "\" & Chr(34))
        %>
            {id:<%=rsBaseNotes("BaseNoteID")%>, name:"<%=bnNameJS%>", desc:"<%=bnDescJS%>", price:<%=SafeNum(rsBaseNotes("UnitPrice"))%>}
        <%
                rsBaseNotes.MoveNext
            Loop
            rsBaseNotes.Close
            Set rsBaseNotes = Nothing
        End If
        %>
        ];
        
        // 渲染基香网格
        function renderBaseNoteGrid(filterText) {
            var grid = document.getElementById('baseNoteGrid');
            filterText = (filterText || '').toLowerCase();
            var html = '';
            var count = 0;
            for (var i = 0; i < baseNotesData.length; i++) {
                var bn = baseNotesData[i];
                if (filterText && bn.name.toLowerCase().indexOf(filterText) === -1 && bn.desc.toLowerCase().indexOf(filterText) === -1) continue;
                count++;
                var priceHtml = bn.price > 0 ? '¥' + bn.price.toFixed(4) + '/ml' : '<span class="no-price">未设单价</span>';
                var descHtml = bn.desc ? bn.desc : '暂无描述';
                html += '<div class="base-note-card" onclick="selectBaseNote(' + bn.id + ')">';
                html += '<div class="bn-name">' + bn.name + '</div>';
                html += '<div class="bn-desc">' + descHtml + '</div>';
                html += '<div class="bn-price">' + priceHtml + '</div>';
                html += '</div>';
            }
            if (count === 0) {
                html = '<div style="grid-column:1/-1;text-align:center;color:#666;padding:20px;">暂无匹配的基香</div>';
            }
            grid.innerHTML = html;
        }
        
        // 搜索过滤基香
        function filterBaseNotes() {
            var searchText = document.getElementById('baseNoteSearch').value;
            renderBaseNoteGrid(searchText);
        }
        
        // 打开基香选择模态框
        function openBaseNoteModal() {
            document.getElementById('baseNoteModal').classList.add('active');
            document.getElementById('baseNoteSearch').value = '';
            renderBaseNoteGrid('');
            document.getElementById('baseNoteSearch').focus();
        }
        
        // 关闭基香选择模态框
        function closeBaseNoteModal() {
            document.getElementById('baseNoteModal').classList.remove('active');
        }
        
        // 点击选择基香后自动填充明细行
        function selectBaseNote(baseNoteId) {
            var bn = null;
            for (var i = 0; i < baseNotesData.length; i++) {
                if (baseNotesData[i].id === baseNoteId) { bn = baseNotesData[i]; break; }
            }
            if (!bn) return;
            
            // 调用addRow填充数据
            addRowWithData(bn.name, 'BN-' + bn.id, '', 'ml', 100000, bn.price);
            closeBaseNoteModal();
            
            // 自动滚动到底部
            var tbody = document.getElementById('detailsBody');
            tbody.lastElementChild.scrollIntoView({behavior:'smooth'});
        }
        
        // 带数据添加明细行
        function addRowWithData(itemName, itemCode, spec, unit, qty, price) {
            var tbody = document.getElementById('detailsBody');
            var rowCount = tbody.children.length + 1;
            
            var row = document.createElement('tr');
            row.innerHTML = 
                '<td><input type="text" name="item_name_' + rowCount + '" value="' + escapeHtml(itemName) + '" required></td>' +
                '<td><input type="text" name="item_code_' + rowCount + '" value="' + escapeHtml(itemCode) + '"></td>' +
                '<td><input type="text" name="spec_' + rowCount + '" value="' + escapeHtml(spec) + '"></td>' +
                '<td><input type="text" name="unit_' + rowCount + '" value="' + escapeHtml(unit) + '"></td>' +
                '<td><input type="number" name="qty_' + rowCount + '" class="num-input qty" value="' + qty + '" min="0" step="0.01" onchange="calculateRow(this)"></td>' +
                '<td><input type="number" name="price_' + rowCount + '" class="num-input price" value="' + price.toFixed(4) + '" min="0" step="0.0001" onchange="calculateRow(this)"></td>' +
                '<td class="row-total">¥' + (qty * price).toFixed(2) + '</td>' +
                '<td><button type="button" class="btn btn-danger btn-sm" onclick="removeRow(this)"><i class="fas fa-trash"></i></button></td>';
            
            tbody.appendChild(row);
            renumberRows();
        }
        
        // HTML转义
        function escapeHtml(str) {
            if (!str) return '';
            return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&#39;');
        }
        
        // ========== 采购分类与采购类型联动 ==========
        var categoryToOrderType = {
            'RAW': 'RawMaterial',
            'BASE': 'RawMaterial',
            'PACK': 'Packaging',
            'MARKET': 'Packaging',
            'BOTTLE': 'Bottle',
            'PRINTING': 'Printing',
            'SPRAYHEAD': 'SprayHead'
        };
        
        var orderTypeToCategory = {
            'RawMaterial': 'RAW',
            'Packaging': 'PACK',
            'Bottle': 'BOTTLE',
            'Printing': 'PRINTING',
            'SprayHead': 'SPRAYHEAD'
        };
        
        // 根据采购分类自动选择采购类型
        function updateOrderTypeByCategory() {
            var catSelect = document.querySelector('select[name="category_code"]');
            var otSelect = document.querySelector('select[name="order_type"]');
            if (!catSelect || !otSelect) return;
            var catCode = catSelect.value;
            if (catCode && categoryToOrderType[catCode]) {
                otSelect.value = categoryToOrderType[catCode];
            }
            updateBaseNoteVisibility();
        }
        
        // 根据采购类型自动选择采购分类
        function updateCategoryByOrderType() {
            var catSelect = document.querySelector('select[name="category_code"]');
            var otSelect = document.querySelector('select[name="order_type"]');
            if (!catSelect || !otSelect) return;
            var otCode = otSelect.value;
            if (otCode && orderTypeToCategory[otCode]) {
                catSelect.value = orderTypeToCategory[otCode];
            }
            updateBaseNoteVisibility();
        }
        
        // 控制基香按钮可见性（联动时实时切换）
        function updateBaseNoteVisibility() {
            var bnSection = document.getElementById('baseNoteSection');
            if (!bnSection) return;
            var catSelect = document.querySelector('select[name="category_code"]');
            var otSelect = document.querySelector('select[name="order_type"]');
            if (!catSelect || !otSelect) return;
            var show = (catSelect.value === 'BASE') || (otSelect.value === 'RawMaterial');
            bnSection.style.display = show ? 'flex' : 'none';
        }
        
        // 点击遮罩关闭模态框
        document.addEventListener('DOMContentLoaded', function() {
            var modal = document.getElementById('baseNoteModal');
            if (modal) {
                modal.addEventListener('click', function(e) {
                    if (e.target === modal) closeBaseNoteModal();
                });
            }
            // ESC关闭模态框
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape') closeBaseNoteModal();
            });
        });
        
        // 添加明细行
        function addRow() {
            var tbody = document.getElementById('detailsBody');
            var rowCount = tbody.children.length + 1;
            
            var row = document.createElement('tr');
            row.innerHTML = 
                '<td><input type="text" name="item_name_' + rowCount + '" required></td>' +
                '<td><input type="text" name="item_code_' + rowCount + '"></td>' +
                '<td><input type="text" name="spec_' + rowCount + '"></td>' +
                '<td><input type="text" name="unit_' + rowCount + '"></td>' +
                '<td><input type="number" name="qty_' + rowCount + '" class="num-input qty" value="0" min="0" step="0.01" onchange="calculateRow(this)"></td>' +
                '<td><input type="number" name="price_' + rowCount + '" class="num-input price" value="0" min="0" step="0.01" onchange="calculateRow(this)"></td>' +
                '<td class="row-total">¥0.00</td>' +
                '<td><button type="button" class="btn btn-danger btn-sm" onclick="removeRow(this)"><i class="fas fa-trash"></i></button></td>';
            
            tbody.appendChild(row);
            renumberRows();
        }
        
        // 删除明细行
        function removeRow(btn) {
            var tbody = document.getElementById('detailsBody');
            if (tbody.children.length <= 1) {
                alert('至少保留一行明细');
                return;
            }
            btn.closest('tr').remove();
            renumberRows();
        }
        
        // 重新编号行
        function renumberRows() {
            var tbody = document.getElementById('detailsBody');
            var rows = tbody.children;
            
            for (var i = 0; i < rows.length; i++) {
                var rowNum = i + 1;
                var inputs = rows[i].querySelectorAll('input');
                inputs[0].name = 'item_name_' + rowNum;
                inputs[1].name = 'item_code_' + rowNum;
                inputs[2].name = 'spec_' + rowNum;
                inputs[3].name = 'unit_' + rowNum;
                inputs[4].name = 'qty_' + rowNum;
                inputs[5].name = 'price_' + rowNum;
            }
            
            document.getElementById('itemCount').value = rows.length;
        }
        
        // 计算行小计
        function calculateRow(input) {
            var row = input.closest('tr');
            var qty = parseFloat(row.querySelector('.qty').value) || 0;
            var price = parseFloat(row.querySelector('.price').value) || 0;
            var total = qty * price;
            row.querySelector('.row-total').textContent = '¥' + total.toFixed(2);
        }
        
        // 页面加载时确保至少有一行
        window.onload = function() {
            var tbody = document.getElementById('detailsBody');
            if (tbody && tbody.children.length === 0) {
                addRow();
            }
        };
        
        // ========== V11: 历史产品快速选择 ==========
        var historyDataCache = {};
        
        function openHistoryModal() {
            document.getElementById('historyModal').classList.add('active');
            document.getElementById('historySearch').value = '';
            loadHistoryData();
        }
        
        function closeHistoryModal() {
            document.getElementById('historyModal').classList.remove('active');
        }
        
        function getCurrentOrderType() {
            var otSelect = document.querySelector('select[name="order_type"]');
            return otSelect ? otSelect.value : 'RawMaterial';
        }
        
        function loadHistoryData() {
            var grid = document.getElementById('historyGrid');
            grid.innerHTML = '<div style="text-align:center;color:#666;padding:20px;"><i class="fas fa-spinner fa-pulse"></i> 加载中...</div>';
            
            var orderType = getCurrentOrderType();
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'ajax_product_history.asp?ordertype=' + encodeURIComponent(orderType) + '&limit=100', true);
            xhr.onload = function() {
                if (xhr.status === 200) {
                    try {
                        var data = JSON.parse(xhr.responseText);
                        historyDataCache[orderType] = data;
                        renderHistoryGrid(data);
                    } catch(e) {
                        grid.innerHTML = '<div style="text-align:center;color:#F44336;padding:20px;">加载失败</div>';
                    }
                }
            };
            xhr.onerror = function() {
                grid.innerHTML = '<div style="text-align:center;color:#F44336;padding:20px;">网络错误</div>';
            };
            xhr.send();
        }
        
        function searchHistory() {
            var orderType = getCurrentOrderType();
            var data = historyDataCache[orderType] || [];
            var searchText = (document.getElementById('historySearch').value || '').toLowerCase();
            var supplierFilter = document.getElementById('historySupplierFilter').value;
            
            if (data.length === 0) {
                loadHistoryData();
                return;
            }
            
            var filtered = data.filter(function(item) {
                if (searchText && item.itemname.toLowerCase().indexOf(searchText) === -1 && item.itemcode.toLowerCase().indexOf(searchText) === -1) return false;
                if (supplierFilter && item.lastsupplierid != supplierFilter) return false;
                return true;
            });
            renderHistoryGrid(filtered);
        }
        
        function renderHistoryGrid(items) {
            var grid = document.getElementById('historyGrid');
            if (!items || items.length === 0) {
                grid.innerHTML = '<div style="text-align:center;color:#666;padding:20px;">暂无匹配的历史采购记录</div>';
                return;
            }
            
            // Build supplier filter options
            var suppliers = {};
            items.forEach(function(item) {
                if (item.lastsupplier && item.lastsupplierid > 0) {
                    suppliers[item.lastsupplierid] = item.lastsupplier;
                }
            });
            var supplierHtml = '<option value="">全部供应商</option>';
            for (var sid in suppliers) {
                supplierHtml += '<option value="' + sid + '">' + escapeHtml(suppliers[sid]) + '</option>';
            }
            document.getElementById('historySupplierFilter').innerHTML = supplierHtml;
            
            var html = '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
            html += '<thead><tr style="background:rgba(255,255,255,0.03);"><th style="padding:8px;text-align:left;color:#aaa;">产品名称</th><th style="padding:8px;text-align:left;color:#aaa;">编码</th><th style="padding:8px;text-align:left;color:#aaa;">规格</th><th style="padding:8px;text-align:left;color:#aaa;">单位</th><th style="padding:8px;text-align:right;color:#aaa;">最近价格</th><th style="padding:8px;text-align:left;color:#aaa;">供应商</th><th style="padding:8px;text-align:center;color:#aaa;">采购次数</th><th style="padding:8px;"></th></tr></thead><tbody>';
            
            items.forEach(function(item) {
                html += '<tr style="border-bottom:1px solid rgba(255,255,255,0.04);cursor:pointer;" onmouseover="this.style.background=\'rgba(255,152,0,0.08)\'" onmouseout="this.style.background=\'\'">';
                html += '<td style="padding:8px;"><strong>' + escapeHtml(item.itemname) + '</strong></td>';
                html += '<td style="padding:8px;color:#888;">' + escapeHtml(item.itemcode) + '</td>';
                html += '<td style="padding:8px;color:#888;">' + escapeHtml(item.spec) + '</td>';
                html += '<td style="padding:8px;color:#888;">' + escapeHtml(item.unit) + '</td>';
                html += '<td style="padding:8px;text-align:right;color:#4CAF50;">¥' + parseFloat(item.lastprice).toFixed(4) + '</td>';
                html += '<td style="padding:8px;">' + escapeHtml(item.lastsupplier) + '</td>';
                html += '<td style="padding:8px;text-align:center;">' + item.purchasecount + '次</td>';
                html += '<td style="padding:8px;"><button type="button" class="btn-select-base" onclick="selectHistoryItem(\'' + escapeHtml(item.itemname) + '\',\'' + escapeHtml(item.itemcode) + '\',\'' + escapeHtml(item.spec) + '\',\'' + escapeHtml(item.unit) + '\',' + parseFloat(item.lastprice).toFixed(4) + ')">选择</button></td>';
                html += '</tr>';
            });
            html += '</tbody></table>';
            grid.innerHTML = html;
        }
        
        function selectHistoryItem(itemName, itemCode, spec, unit, price) {
            addRowWithData(itemName, itemCode, spec, unit, 0, price);
            closeHistoryModal();
        }
        
        // 点击遮罩关闭
        document.addEventListener('DOMContentLoaded', function() {
            var hModal = document.getElementById('historyModal');
            if (hModal) {
                hModal.addEventListener('click', function(e) {
                    if (e.target === hModal) closeHistoryModal();
                });
            }
            
            // V11: 退出确认 - 未保存内容时提示
            var formSection = document.querySelector('.form-section');
            if (formSection) {
                window.addEventListener('beforeunload', function(e) {
                    var tbody = document.getElementById('detailsBody');
                    if (tbody && tbody.children.length > 0) {
                        var hasContent = false;
                        var inputs = tbody.querySelectorAll('input[name^="item_name_"]');
                        inputs.forEach(function(inp) {
                            if (inp.value.trim() !== '') hasContent = true;
                        });
                        if (hasContent) {
                            e.preventDefault();
                            e.returnValue = '您有未保存的采购明细，确定离开吗？';
                            return e.returnValue;
                        }
                    }
                });
            }
        });
        
        // ========== V11: 批量操作 ==========
        function toggleSelectAll(cb) {
            var checks = document.querySelectorAll('.row-check');
            checks.forEach(function(c) { c.checked = cb.checked; });
            updateBatchCount();
        }
        
        function updateBatchCount() {
            var checked = document.querySelectorAll('.row-check:checked');
            var countEl = document.getElementById('batchCount');
            if (checked.length > 0) {
                countEl.textContent = '已选择 ' + checked.length + ' 个订单';
                countEl.style.color = '#FF9800';
            } else {
                countEl.textContent = '';
            }
        }
        
        function batchAction(action) {
            var checked = document.querySelectorAll('.row-check:checked');
            if (checked.length === 0) {
                alert('请至少选择一个订单');
                return;
            }
            
            var ids = [];
            checked.forEach(function(c) { ids.push(c.value); });
            
            var actionLabel = action === 'approve' ? '审批' : (action === 'order' ? '下单' : '收货');
            if (!confirm('确定要复制此订单吗？将创建一个新的草稿订单。')) return;
            
            // Create and submit form
            var form = document.createElement('form');
            form.method = 'POST';
            form.style.display = 'none';
            
            var inputAction = document.createElement('input');
            inputAction.name = 'batch_action';
            inputAction.value = action;
            form.appendChild(inputAction);
            
            var inputIds = document.createElement('input');
            inputIds.name = 'batch_ids';
            inputIds.value = ids.join(',');
            form.appendChild(inputIds);
            
            document.body.appendChild(form);
            form.submit();
        }
        function copyOrderFromView(orderId) {
            if (!confirm('确定要复制此订单吗？将创建一个新的草稿订单。')) return;
            var form = document.createElement('form');
            form.method = 'POST';
            form.style.display = 'none';
            var input1 = document.createElement('input');
            input1.name = 'action';
            input1.value = 'copy';
            form.appendChild(input1);
            var input2 = document.createElement('input');
            input2.name = 'purchase_id';
            input2.value = orderId;
            form.appendChild(input2);
            document.body.appendChild(form);
            form.submit();
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
