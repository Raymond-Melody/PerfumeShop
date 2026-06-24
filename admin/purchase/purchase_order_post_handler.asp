<%
' ============================================
' V14.6 采购订单 - POST请求处理器
' 从 purchase_orders.asp 提取
' 包含：创建/更新/状态变更/批量操作/复制订单
' ============================================

' 消息变量初始化
Dim action, message, messageType
action = Request.Form("action") & ""
If action = "" Then action = Request.QueryString("action") & ""
message = ""
messageType = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "create" Then
        ' 创建采购订单
        Dim newPurchaseNo, newSupplierID, newCategory, newOrderType, newExpectedDate, newRemarks
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
        Dim editPurchaseID, editSupplierID, editCategory, editOrderType, editExpectedDate, editRemarks
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
%>