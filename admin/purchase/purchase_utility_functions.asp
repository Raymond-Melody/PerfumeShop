<%
' ============================================
' V14.6 采购订单 - 业务工具函数
' 从 purchase_orders.asp 提取
' 注意: SafeNum/SafeDiv 已移至 common_utils.asp
' ============================================

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
%>
