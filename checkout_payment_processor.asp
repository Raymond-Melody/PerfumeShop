<%
' ============================================
' V14.6 结算页 - 支付处理器
' 从 checkout.asp 提取
' ============================================

' 处理支付请求
Dim paymentMethod, orderId, paymentResult, realName, phone, address, selectedAddressId, paymentError, cartSnapshotSql, rsCartSnapshot, productDetails, itemCount, itemDesc, noteInfo

' 初始化变量
paymentResult = False
realName = ""
phone = ""
address = ""
orderId = 0
paymentMethod = Request.Form("payment_method")
If paymentMethod = "" Then
    paymentMethod = Request.QueryString("payment_method")
End If

If paymentMethod <> "" And IsNumeric(paymentMethod) Then
    paymentMethod = CLng(paymentMethod)
End If

selectedAddressId = Request.Form("selectedAddress")

If paymentMethod <> "" Then
    ' 计算订单总金额
    Dim finalAmount
    If discountedGrandTotal >= FREE_SHIPPING_AMOUNT Then
        finalAmount = discountedGrandTotal
    Else
        finalAmount = discountedGrandTotal + SHIPPING_FEE
    End If
    
    ' 判断是否是已存在订单模式
    If isExistingOrder Then
        orderId = existingOrderId
        realName = rsExistingOrder("ShippingName")
        phone = rsExistingOrder("ShippingPhone")
        address = rsExistingOrder("ShippingAddress")
        rsExistingOrder.Close
        Set rsExistingOrder = Nothing
    Else
        ' 新订单模式
    
    ' 检查是否选择了现有地址
    If selectedAddressId <> "" And selectedAddressId <> "new" And IsNumeric(selectedAddressId) Then
        Dim rsSelectedAddress
        Set rsSelectedAddress = ExecuteQuery("SELECT * FROM UserAddresses WHERE AddressID = " & selectedAddressId & " AND UserID = " & userId)
            
        If Not rsSelectedAddress Is Nothing And Not rsSelectedAddress.EOF Then
            realName = rsSelectedAddress("Consignee")
            phone = rsSelectedAddress("Phone")
            address = BuildFullAddress(rsSelectedAddress("Province"), rsSelectedAddress("City"), rsSelectedAddress("District"), rsSelectedAddress("Address"))
                        
            ' 重新查询购物车数据以获取商品信息，用于订单快照
            productDetails = ""
                        
            cartSnapshotSql = "SELECT c.*, p.ProductName, p.ImageURL, p.ProductType, " & _
                "tn.NoteName AS TopNoteName, mn.NoteName AS MiddleNoteName, bn.NoteName AS BaseNoteName, " & _
                "v.VolumeName, v.VolumeML, b.BottleName, " & _
                "c.Quantity * c.UnitPrice AS SubTotal " & _
                "FROM ((((((Cart c " & _
                "LEFT JOIN Products p ON c.ProductID = p.ProductID) " & _
                "LEFT JOIN FragranceNotes tn ON c.TopNoteID = tn.NoteID) " & _
                "LEFT JOIN FragranceNotes mn ON c.MiddleNoteID = mn.NoteID) " & _
                "LEFT JOIN FragranceNotes bn ON c.BaseNoteID = bn.NoteID) " & _
                "LEFT JOIN Volumes v ON c.VolumeID = v.VolumeID) " & _
                "LEFT JOIN BottleStyles b ON c.BottleID = b.BottleID) " & _
                "WHERE " & whereClause & " ORDER BY c.CreatedAt DESC"
            
            Set rsCartSnapshot = ExecuteQuery(cartSnapshotSql)
            
            If Not rsCartSnapshot Is Nothing Then
                itemCount = 0
                Do While Not rsCartSnapshot.EOF
                    If productDetails <> "" Then productDetails = productDetails & "|"
                    itemDesc = rsCartSnapshot("ProductName") & " x" & rsCartSnapshot("Quantity") & " (" & FormatMoney(rsCartSnapshot("SubTotal")) & ")"
                    
                    Dim snapshotPT1
                    snapshotPT1 = LCase(rsCartSnapshot("ProductType") & "")
                    noteInfo = ""
                    If snapshotPT1 = "custom" Then
                        If Not IsNull(rsCartSnapshot("TopNoteName")) And rsCartSnapshot("TopNoteName") <> "" Then
                            noteInfo = noteInfo & "前调:" & rsCartSnapshot("TopNoteName")
                        End If
                        If Not IsNull(rsCartSnapshot("MiddleNoteName")) And rsCartSnapshot("MiddleNoteName") <> "" Then
                            If noteInfo <> "" Then noteInfo = noteInfo & ", "
                            noteInfo = noteInfo & "中调:" & rsCartSnapshot("MiddleNoteName")
                        End If
                        If Not IsNull(rsCartSnapshot("BaseNoteName")) And rsCartSnapshot("BaseNoteName") <> "" Then
                            If noteInfo <> "" Then noteInfo = noteInfo & ", "
                            noteInfo = noteInfo & "后调:" & rsCartSnapshot("BaseNoteName")
                        End If
                    End If
                    
                    If Not IsNull(rsCartSnapshot("VolumeName")) And rsCartSnapshot("VolumeName") <> "" Then
                        If noteInfo <> "" Then noteInfo = noteInfo & ", "
                        noteInfo = noteInfo & rsCartSnapshot("VolumeName")
                    End If
                    If Not IsNull(rsCartSnapshot("VolumeML")) And rsCartSnapshot("VolumeML") <> "" Then
                        If noteInfo <> "" Then noteInfo = noteInfo & ", "
                        noteInfo = noteInfo & rsCartSnapshot("VolumeML") & "ml"
                    End If
                    If Not IsNull(rsCartSnapshot("BottleName")) And rsCartSnapshot("BottleName") <> "" Then
                        If noteInfo <> "" Then noteInfo = noteInfo & ", "
                        noteInfo = noteInfo & "瓶身:" & rsCartSnapshot("BottleName")
                    End If
                    
                    If noteInfo <> "" Then
                        itemDesc = itemDesc & " [" & noteInfo & "]"
                    End If
                    
                    productDetails = productDetails & itemDesc
                    itemCount = itemCount + 1
                    rsCartSnapshot.MoveNext
                Loop
                rsCartSnapshot.Close
                Set rsCartSnapshot = Nothing
            End If
            
            If productDetails = "" Then
                productDetails = "香水订单 (" & cartCount & " 件商品, 总计: " & FormatMoney(cartTotal) & ")"
            Else
                productDetails = "详情: " & productDetails
            End If
                
            orderId = SafeCreatePaymentOrder(userId, finalAmount, productDetails, paymentMethod, realName, phone, address)
                            
            If orderId > 0 Then
                Dim updateUserInfoSql
                updateUserInfoSql = "UPDATE Users SET FullName = '" & SafeSQL(realName) & "', Phone = '" & SafeSQL(phone) & "', Address = '" & SafeSQL(address) & "' WHERE UserID = " & userId
                Call ExecuteNonQuery(updateUserInfoSql)
                
                Call SyncOrderDetailsAndIngredients(orderId, userId, whereClause)
                Call CE_UpdateOrderCosts(orderId)
            Else
                If Session("ErrorMessage") = "" Then
                    Session("ErrorMessage") = "订单创建失败，请检查购物车和支付方式。Payment: '" & paymentMethod & "'"
                End If
            End If
        Else
            Session("ErrorMessage") = "未找到选择的地址"
        End If
            
        If Not rsSelectedAddress Is Nothing Then
            rsSelectedAddress.Close
            Set rsSelectedAddress = Nothing
        End If
    Else
        ' 使用新地址
        realName = SafeSQL(Request.Form("realName"))
        phone = SafeSQL(Request.Form("phone"))
        provinceName = SafeSQL(Trim(Request.Form("province")))
        cityName = SafeSQL(Trim(Request.Form("city")))
        districtName = SafeSQL(Trim(Request.Form("district")))
        detailAddress = SafeSQL(Request.Form("address"))
            
        address = BuildFullAddress(provinceName, cityName, districtName, detailAddress)
            
        If realName = "" Or phone = "" Or provinceName = "" Or cityName = "" Or districtName = "" Or detailAddress = "" Then
            Session("ErrorMessage") = "请填写完整的收货信息"
        Else
            productDetails = ""
            
            cartSnapshotSql = "SELECT c.*, p.ProductName, p.ImageURL, p.ProductType, " & _
                "tn.NoteName AS TopNoteName, mn.NoteName AS MiddleNoteName, bn.NoteName AS BaseNoteName, " & _
                "v.VolumeName, v.VolumeML, b.BottleName, " & _
                "c.Quantity * c.UnitPrice AS SubTotal " & _
                "FROM ((((((Cart c " & _
                "LEFT JOIN Products p ON c.ProductID = p.ProductID) " & _
                "LEFT JOIN FragranceNotes tn ON c.TopNoteID = tn.NoteID) " & _
                "LEFT JOIN FragranceNotes mn ON c.MiddleNoteID = mn.NoteID) " & _
                "LEFT JOIN FragranceNotes bn ON c.BaseNoteID = bn.NoteID) " & _
                "LEFT JOIN Volumes v ON c.VolumeID = v.VolumeID) " & _
                "LEFT JOIN BottleStyles b ON c.BottleID = b.BottleID) " & _
                "WHERE " & whereClause & " ORDER BY c.CreatedAt DESC"
            
            Set rsCartSnapshot = ExecuteQuery(cartSnapshotSql)
            
            If Not rsCartSnapshot Is Nothing Then
                itemCount = 0
                Do While Not rsCartSnapshot.EOF
                    If productDetails <> "" Then productDetails = productDetails & "|"
                    itemDesc = rsCartSnapshot("ProductName") & " x" & rsCartSnapshot("Quantity") & " (" & FormatMoney(rsCartSnapshot("SubTotal")) & ")"
                    
                    Dim snapshotPT2
                    snapshotPT2 = LCase(rsCartSnapshot("ProductType") & "")
                    noteInfo = ""
                    If snapshotPT2 = "custom" Then
                        If Not IsNull(rsCartSnapshot("TopNoteName")) And rsCartSnapshot("TopNoteName") <> "" Then
                            noteInfo = noteInfo & "前调:" & rsCartSnapshot("TopNoteName")
                        End If
                        If Not IsNull(rsCartSnapshot("MiddleNoteName")) And rsCartSnapshot("MiddleNoteName") <> "" Then
                            If noteInfo <> "" Then noteInfo = noteInfo & ", "
                            noteInfo = noteInfo & "中调:" & rsCartSnapshot("MiddleNoteName")
                        End If
                        If Not IsNull(rsCartSnapshot("BaseNoteName")) And rsCartSnapshot("BaseNoteName") <> "" Then
                            If noteInfo <> "" Then noteInfo = noteInfo & ", "
                            noteInfo = noteInfo & "后调:" & rsCartSnapshot("BaseNoteName")
                        End If
                    End If
                    
                    If Not IsNull(rsCartSnapshot("VolumeName")) And rsCartSnapshot("VolumeName") <> "" Then
                        If noteInfo <> "" Then noteInfo = noteInfo & ", "
                        noteInfo = noteInfo & rsCartSnapshot("VolumeName")
                    End If
                    If Not IsNull(rsCartSnapshot("VolumeML")) And rsCartSnapshot("VolumeML") <> "" Then
                        If noteInfo <> "" Then noteInfo = noteInfo & ", "
                        noteInfo = noteInfo & rsCartSnapshot("VolumeML") & "ml"
                    End If
                    If Not IsNull(rsCartSnapshot("BottleName")) And rsCartSnapshot("BottleName") <> "" Then
                        If noteInfo <> "" Then noteInfo = noteInfo & ", "
                        noteInfo = noteInfo & "瓶身:" & rsCartSnapshot("BottleName")
                    End If
                    
                    If noteInfo <> "" Then
                        itemDesc = itemDesc & " [" & noteInfo & "]"
                    End If
                    
                    productDetails = productDetails & itemDesc
                    itemCount = itemCount + 1
                    rsCartSnapshot.MoveNext
                Loop
                rsCartSnapshot.Close
                Set rsCartSnapshot = Nothing
            End If
            
            If productDetails = "" Then
                productDetails = "Perfume Order (" & cartCount & " items, Total: " & FormatMoney(cartTotal) & ")"
            Else
                productDetails = "Details: " & productDetails
            End If
            
            orderId = SafeCreatePaymentOrder(userId, finalAmount, productDetails, paymentMethod, realName, phone, address)
            
            If orderId > 0 Then
                updateUserInfoSql = "UPDATE Users SET FullName = '" & SafeSQL(realName) & "', Phone = '" & SafeSQL(phone) & "', Address = '" & SafeSQL(address) & "' WHERE UserID = " & userId
                Call ExecuteNonQuery(updateUserInfoSql)
                
                Call SyncOrderDetailsAndIngredients(orderId, userId, whereClause)
                Call CE_UpdateOrderCosts(orderId)
            Else
                If Session("ErrorMessage") = "" Then
                    Session("ErrorMessage") = "订单创建失败，请检查购物车和支付方式。Payment: '" & paymentMethod & "'"
                End If
            End If
        End If
    End If
    End If  ' 结束新订单模式的判断
    
    ' 支付处理
    If orderId > 0 Then
                
                Select Case paymentMethod
                    Case PAYMENT_METHOD_WECHAT
                        On Error Resume Next
                        paymentResult = ProcessWeChatPay(orderId)
                        If Err.Number <> 0 Then
                            Session("ErrorMessage") = "WeChat pay error: " & Err.Description
                            paymentResult = False
                        End If
                        On Error Goto 0
                    Case PAYMENT_METHOD_ALIPAY
                        On Error Resume Next
                        paymentResult = ProcessAlipay(orderId)
                        If Err.Number <> 0 Then
                            Session("ErrorMessage") = "Alipay error: " & Err.Description
                            paymentResult = False
                        End If
                        On Error Goto 0
                    Case PAYMENT_METHOD_PAYPAL
                        On Error Resume Next
                        paymentResult = ProcessPayPal(orderId)
                        If Err.Number <> 0 Then
                            Session("ErrorMessage") = "PayPal error: " & Err.Description
                            paymentResult = False
                        End If
                        On Error Goto 0
                    Case PAYMENT_METHOD_COD
                        On Error Resume Next
                        paymentResult = ProcessCashOnDelivery(orderId)
                        If Err.Number <> 0 Then
                            Session("ErrorMessage") = "COD error: " & Err.Description
                            paymentResult = False
                        End If
                        On Error Goto 0
                End Select
                
                If paymentResult Then
                    If Not isExistingOrder Then
                        Dim actualDeleteClause2
                        If IsEmpty(deleteClause) Then
                            actualDeleteClause2 = whereClause
                            actualDeleteClause2 = Replace(actualDeleteClause2, "c.", "")
                        Else
                            actualDeleteClause2 = deleteClause
                        End If
                        Dim clearSql, clearResult
                        clearSql = "DELETE FROM Cart WHERE " & actualDeleteClause2
                        Session("DebugCartClearSQL") = clearSql
                        clearResult = ExecuteNonQuery(clearSql)
                        Session("DebugCartClearResult") = clearResult
                        If Not clearResult Then
                            Session("ErrorMessage") = "Cart clear failed: " & Session("LastDBError")
                        End If
                    End If
                    
                    Response.Redirect "/order_success.asp?order_id=" & orderId
                    Response.End
                Else
                    If Session("ErrorMessage") = "" Then
                        Session("ErrorMessage") = "Payment processing failed"
                    End If
                End If
    End If  ' 结束If orderId > 0 Then
End If  ' 结束If paymentMethod <> "" Then
%>