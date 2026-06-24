<%
' ============================================
' V14.6 结算页 - 购物车/订单数据加载
' 从 checkout.asp 提取
' ============================================

' 查询支付方式启用状态
Dim enableAlipay, enableWechat, enablePaypal, enableCOD, enableBankTransfer
enableAlipay = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableAlipay'")
enableWechat = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableWechatPay'")
enablePaypal = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnablePayPal'")
enableCOD = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableCOD'")
enableBankTransfer = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableBankTransfer'")
' 默认值处理
If IsNull(enableAlipay) Or enableAlipay = "" Then enableAlipay = "1"
If IsNull(enableWechat) Or enableWechat = "" Then enableWechat = "1"
If IsNull(enablePaypal) Or enablePaypal = "" Then enablePaypal = "1"
If IsNull(enableCOD) Or enableCOD = "" Then enableCOD = "1"
If IsNull(enableBankTransfer) Or enableBankTransfer = "" Then enableBankTransfer = "1"

Dim userId
userId = Session("UserID")

' 检查是否是从"立即支付"按钮跳转过来的（已存在的未支付订单）
Dim existingOrderId, isExistingOrder
existingOrderId = Trim(Request.QueryString("order_id"))
isExistingOrder = False

If existingOrderId <> "" And IsNumeric(existingOrderId) Then
    ' 验证订单归属和状态
    Dim rsExistingOrder
    Set rsExistingOrder = ExecuteQuery("SELECT OrderID, OrderNo, TotalAmount, Status, PaymentMethod, ShippingName, ShippingPhone, ShippingAddress FROM Orders WHERE OrderID = " & existingOrderId & " AND UserID = " & userId)

    If Not rsExistingOrder Is Nothing And Not rsExistingOrder.EOF Then
        If rsExistingOrder("Status") = "Pending" Then
            isExistingOrder = True
        Else
            rsExistingOrder.Close
            Set rsExistingOrder = Nothing
            Response.Redirect "/user/order_detail.asp?order_id=" & existingOrderId
            Response.End
        End If
    Else
        If Not rsExistingOrder Is Nothing Then
            rsExistingOrder.Close
            Set rsExistingOrder = Nothing
        End If
        Response.Redirect "/user/orders.asp"
        Response.End
    End If
End If

' 获取购物车信息（只在非已存在订单模式下）
Dim sessionId, whereClause, rsCart, cartTotal, cartCount, cartIds, cartIdList
Dim deleteClause
sessionId = Session.SessionID

If Not isExistingOrder Then
    ' 获取从购物车页面传递的选中商品ID
    cartIds = Trim(Request.QueryString("cart_ids"))
    If cartIds = "" Then
        cartIds = Trim(Request.Form("cart_ids"))
    End If

    If userId <> "" Then
        whereClause = "UserID = " & userId
    Else
        whereClause = "SessionID = '" & SafeSQL(sessionId) & "'"
    End If

    ' 如果有指定的商品ID，只查询这些商品
    If cartIds <> "" Then
        Dim cleanCartIds, idArr, i, tempId
        cleanCartIds = ""
        idArr = Split(cartIds, ",")
        For i = 0 To UBound(idArr)
            tempId = Trim(idArr(i))
            If IsNumeric(tempId) And tempId <> "" Then
                If cleanCartIds <> "" Then cleanCartIds = cleanCartIds & ","
                cleanCartIds = cleanCartIds & CLng(tempId)
            End If
        Next

        If cleanCartIds <> "" Then
            whereClause = whereClause & " AND c.CartID IN (" & cleanCartIds & ")"
            deleteClause = whereClause
            deleteClause = Replace(deleteClause, "c.CartID", "CartID")
            deleteClause = Replace(deleteClause, "UserID = ", "UserID = ")
            deleteClause = Replace(deleteClause, "SessionID = ", "SessionID = ")
            cartIdList = cleanCartIds
        End If
    End If

    Set rsCart = ExecuteQuery("SELECT c.*, p.ProductName, p.ImageURL, p.EngravingPrice, " & _
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
        "WHERE " & whereClause & " ORDER BY c.CreatedAt DESC")

    cartTotal = 0
    cartCount = 0
    Dim totalEngravingFee, itemEngravingPrice
    totalEngravingFee = 0

    If Not rsCart Is Nothing Then
        Do While Not rsCart.EOF
            cartTotal = cartTotal + CDbl(rsCart("SubTotal"))
            cartCount = cartCount + 1
            itemEngravingPrice = 0
            On Error Resume Next
            itemEngravingPrice = CDbl(rsCart("EngravingPrice"))
            If Err.Number <> 0 Then itemEngravingPrice = 0
            On Error GoTo 0
            If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" And itemEngravingPrice > 0 Then
                totalEngravingFee = totalEngravingFee + (itemEngravingPrice * rsCart("Quantity"))
            End If
            rsCart.MoveNext
        Loop
        rsCart.Close
        Set rsCart = Nothing
    End If

    ' 计算应付总额
    Dim grandTotal
    grandTotal = cartTotal + totalEngravingFee

    ' 计算会员折扣
    Dim memberLevel, memberDiscount, memberDiscountAmount, discountedGrandTotal
    memberLevel = MU_CalcUserLevel(userId)
    memberDiscount = MU_GetLevelDiscount(memberLevel)
    If memberDiscount < 1.0 Then
        memberDiscountAmount = grandTotal * (1 - memberDiscount)
        discountedGrandTotal = grandTotal - memberDiscountAmount
    Else
        memberDiscountAmount = 0
        discountedGrandTotal = grandTotal
    End If

    ' 如果购物车为空，跳转回购物车页面
    If cartCount = 0 Then
        Response.Redirect "/cart.asp"
        Response.End
    End If
Else
    ' 已存在订单模式
    cartTotal = CDbl(rsExistingOrder("TotalAmount"))
    cartCount = 1
    grandTotal = cartTotal
    memberLevel = MU_CalcUserLevel(userId)
    memberDiscount = MU_GetLevelDiscount(memberLevel)
    If memberDiscount < 1.0 Then
        memberDiscountAmount = grandTotal * (1 - memberDiscount)
        discountedGrandTotal = grandTotal - memberDiscountAmount
    Else
        memberDiscountAmount = 0
        discountedGrandTotal = grandTotal
    End If
End If

' 获取用户地址信息
Dim rsUser, userAddress, userPhone, userFullName
userAddress = ""
userPhone = ""
userFullName = ""

If userId <> "" Then
    Set rsUser = ExecuteQuery("SELECT FullName, Phone, Address FROM Users WHERE UserID = " & userId)
    If Not rsUser Is Nothing Then
        If Not rsUser.EOF Then
            userFullName = rsUser("FullName")
            userPhone = rsUser("Phone")
            userAddress = rsUser("Address")
        End If
        rsUser.Close
        Set rsUser = Nothing
    End If
End If
%>
