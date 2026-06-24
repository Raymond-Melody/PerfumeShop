<%
' ============================================
' V14.6 采购订单 - 消息处理 + 数据加载器
' 从 purchase_orders.asp 提取
' 包含：消息映射、筛选参数、分页查询、编辑/查看数据加载
' ============================================

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
