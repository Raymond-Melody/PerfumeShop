<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/payment_handler.asp"-->
<!--#include file="includes/cost_engine.asp"-->
<%
Call OpenConnection()

' 查询支付方式启用状态（在打开任何Recordset之前执行，避免Access MARS问题）
Dim enableAlipay, enableWechat, enablePaypal, enableCOD, enableBankTransfer
enableAlipay = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableAlipay'")
enableWechat = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableWechatPay'")
enablePaypal = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnablePayPal'")
enableCOD = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableCOD'")
enableBankTransfer = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey='EnableBankTransfer'")
' 默认值处理：如果配置项不存在，默认启用
If IsNull(enableAlipay) Or enableAlipay = "" Then enableAlipay = "1"
If IsNull(enableWechat) Or enableWechat = "" Then enableWechat = "1"
If IsNull(enablePaypal) Or enablePaypal = "" Then enablePaypal = "1"
If IsNull(enableCOD) Or enableCOD = "" Then enableCOD = "1"
If IsNull(enableBankTransfer) Or enableBankTransfer = "" Then enableBankTransfer = "1"

' 调试日志函数
Sub DebugLog(msg)
    Dim fso, logFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile(Server.MapPath("/debug_checkout.log"), 8, True)
    logFile.WriteLine Now() & " - " & msg
    logFile.Close
    Set logFile = Nothing
    Set fso = Nothing
End Sub

' 通用成分分割函数 - 支持逗号、空格、换行符等多种分隔符
Function SplitIngredientsUniversal(rawStr)
    Dim result, arr, item, i
    Set result = CreateObject("Scripting.Dictionary")
    
    If rawStr = "" Then
        Set SplitIngredientsUniversal = result
        Exit Function
    End If
    
    ' 统一将所有分隔符转换为英文逗号
    rawStr = Replace(rawStr, "，", ",")      ' 中文逗号
    rawStr = Replace(rawStr, vbCrLf, ",")   ' 回车换行
    rawStr = Replace(rawStr, vbLf, ",")     ' 换行符
    rawStr = Replace(rawStr, vbCr, ",")     ' 回车符
    rawStr = Replace(rawStr, Chr(160), ",") ' NBSP
    rawStr = Replace(rawStr, "　", ",")     ' 全角空格
    
    ' 清理连续逗号
    Do While InStr(rawStr, ",,") > 0
        rawStr = Replace(rawStr, ",,", ",")
    Loop
    
    ' 用逗号分割
    arr = Split(rawStr, ",")
    For i = 0 To UBound(arr)
        item = Trim(arr(i))
        If item <> "" And Not result.Exists(item) Then
            result.Add item, True
        End If
    Next
    
    Set SplitIngredientsUniversal = result
End Function

' 检查用户是否登录
If Session("UserID") = "" Then
    Dim returnFullUrl
    returnFullUrl = Request.ServerVariables("SCRIPT_NAME")
    If Request.ServerVariables("QUERY_STRING") <> "" Then
        returnFullUrl = returnFullUrl & "?" & Request.ServerVariables("QUERY_STRING")
    End If
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(returnFullUrl)
    Response.End
End If

Dim userId
userId = Session("UserID")

' 检查是否是从“立即支付”按钮跳转过来的（已存在的未支付订单）
Dim existingOrderId, isExistingOrder
existingOrderId = Trim(Request.QueryString("order_id"))
isExistingOrder = False

If existingOrderId <> "" And IsNumeric(existingOrderId) Then
    ' 验证订单归属和状态
    Dim rsExistingOrder
    Set rsExistingOrder = ExecuteQuery("SELECT OrderID, OrderNo, TotalAmount, Status, PaymentMethod, ShippingName, ShippingPhone, ShippingAddress FROM Orders WHERE OrderID = " & existingOrderId & " AND UserID = " & userId)
    
    If Not rsExistingOrder Is Nothing And Not rsExistingOrder.EOF Then
        If rsExistingOrder("Status") = "Pending" Then
            ' 这是一个有效的未支付订单，标记为已存在订单模式
            isExistingOrder = True
        Else
            ' 订单已支付或其他状态，跳转到订单详情
            rsExistingOrder.Close
            Set rsExistingOrder = Nothing
            Response.Redirect "/user/order_detail.asp?order_id=" & existingOrderId
            Response.End
        End If
    Else
        ' 订单不存在或无权访问，跳转回订单列表
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
sessionId = Session.SessionID

If Not isExistingOrder Then
    ' 获取从购物车页面传递的选中商品ID（支持GET和POST）
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
        ' 验证并清理cartIds，防止SQL注入
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
            ' 用于DELETE语句的条件（不带别名）
            deleteClause = whereClause
            deleteClause = Replace(deleteClause, "c.CartID", "CartID")
            deleteClause = Replace(deleteClause, "UserID = ", "UserID = ")  ' 确保UserID部分正确
            deleteClause = Replace(deleteClause, "SessionID = ", "SessionID = ")  ' 确保SessionID部分正确
            ' 保存清理后的ID列表供后续使用
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
            ' 计算刻字费用
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
    
    ' 计算应付总额（商品金额 + 刻字费用）
    Dim grandTotal
    grandTotal = cartTotal + totalEngravingFee
    
    ' 如果购物车为空，跳转回购物车页面
    If cartCount = 0 Then
        Response.Redirect "/cart.asp"
        Response.End
    End If
Else
    ' 已存在订单模式，从订单中获取信息
    cartTotal = CDbl(rsExistingOrder("TotalAmount"))
    cartCount = 1  ' 订单已存在，不需要计数
    ' 保持rsExistingOrder打开，后面需要用到其中的信息
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()

' CSRF验证 - 对所有POST请求进行验证
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        Response.Write "<script>alert('安全验证失败，请刷新页面重试'); history.back();</script>"
        Response.End
    End If
End If

' 处理表单提交 - 添加地址
If Request.Form("action") = "add" Then
    Dim consignee, phoneNum, provinceName, cityName, districtName, detailAddress, isDefaultAddr
    consignee = SafeSQL(Trim(Request.Form("realName")))
    phoneNum = SafeSQL(Trim(Request.Form("phone")))
    provinceName = SafeSQL(Trim(Request.Form("province")))
    cityName = SafeSQL(Trim(Request.Form("city")))
    districtName = SafeSQL(Trim(Request.Form("district")))
    detailAddress = SafeSQL(Trim(Request.Form("address")))
    isDefaultAddr = Request.Form("isDefault")

    If isDefaultAddr <> "" And isDefaultAddr <> "0" Then
        isDefaultAddr = 1
    Else
        isDefaultAddr = 0
    End If
    
    ' 验证收货信息
    If consignee = "" Or phoneNum = "" Or provinceName = "" Or cityName = "" Or districtName = "" Or detailAddress = "" Then
        Session("ErrorMessage") = "请填写完整的收货信息"
    Else
        ' 如果设为默认地址，先取消其他默认地址
        If isDefaultAddr <> 0 Then
            Call ExecuteNonQuery("UPDATE UserAddresses SET IsDefault = 0 WHERE UserID = " & userId)
        End If
            
        Dim insertSql
        insertSql = "INSERT INTO UserAddresses (UserID, Consignee, Phone, Province, City, District, Address, IsDefault, CreatedAt) VALUES (" & userId & ", '" & consignee & "', '" & phoneNum & "', '" & provinceName & "', '" & cityName & "', '" & districtName & "', '" & detailAddress & "', " & isDefaultAddr & ", GETDATE())"
        
        If ExecuteNonQuery(insertSql) Then
            ' 重新加载页面以显示新地址（保留cart_ids和payment_method参数）
            Dim newAddressId
            newAddressId = GetLastInsertID("UserAddresses")  ' 获取新插入地址的ID
            Dim paymentMethodFromForm
            paymentMethodFromForm = Request.Form("payment_method")
            Dim redirectUrl
            redirectUrl = "checkout.asp"
            
            ' 构建查询字符串
            Dim queryString
            queryString = ""
            
            If cartIds <> "" Then
                queryString = queryString & "cart_ids=" & cartIds
            End If
            
            If paymentMethodFromForm <> "" Then
                If queryString <> "" Then queryString = queryString & "&"
                queryString = queryString & "payment_method=" & paymentMethodFromForm
            End If
            
            If isDefaultAddr <> 0 Then
                If queryString <> "" Then queryString = queryString & "&"
                queryString = queryString & "selected_address=" & newAddressId
            End If
            
            If queryString <> "" Then
                redirectUrl = redirectUrl & "?" & queryString
            End If
            
            Response.Redirect redirectUrl
            Response.End
        Else
        Session("ErrorMessage") = "地址保存失败，请重试"
        End If
    End If
End If

' 安全创建支付订单函数（替代CreatePaymentOrder）
Function SafeCreatePaymentOrder(userId, orderAmount, orderDesc, paymentMethod, shippingName, shippingPhone, shippingAddress)
    Dim orderId, orderNo, sql, insertSuccess
    orderId = 0
    
    ' 验证订单金额
    If Not IsNumeric(orderAmount) Or CDbl(orderAmount) <= 0 Then
        Session("ErrorMessage") = "订单金额无效"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' 验证支付方式
    If paymentMethod = "" Then
        Session("ErrorMessage") = "支付方式为空"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    If Not IsNumeric(paymentMethod) Then
        Session("ErrorMessage") = "Payment method is not numeric: '" & paymentMethod & "'"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' 检查支付方式是否在允许范围内
    Dim validPaymentMethod
    validPaymentMethod = (CLng(paymentMethod) = 1 Or CLng(paymentMethod) = 2 Or CLng(paymentMethod) = 3 Or CLng(paymentMethod) = 4)
    If Not validPaymentMethod Then
        Session("ErrorMessage") = "支付方式参数无效: " & paymentMethod & "。支持的支付方式: 1(微信), 2(支付宝), 3(PayPal), 4(货到付款)"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' 验证用户ID
    If Not IsNumeric(userId) Or userId <= 0 Then
        Session("ErrorMessage") = "Invalid user ID"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' 验证订单描述
    If IsNull(orderDesc) Then
        orderDesc = ""
    End If
    
    ' 验证收货人信息（可选）
    If IsNull(shippingName) Then shippingName = ""
    If IsNull(shippingPhone) Then shippingPhone = ""
    If IsNull(shippingAddress) Then shippingAddress = ""
    
    ' 生成订单号（时间戳+用户ID后几位）
    orderNo = "ORD" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
              Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & _
              Right(userId, 4)
    
    ' 验证订单号生成是否成功
    If orderNo = "" Then
        Session("ErrorMessage") = "Order number generation failed"
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    ' 构建插入订单记录的SQL语句，使用正确的字段名，包括收货人信息
    sql = "INSERT INTO Orders (OrderNo, UserID, TotalAmount, Notes, PaymentMethod, Status, ShippingName, ShippingPhone, ShippingAddress, CreatedAt) " & _
          "VALUES ('" & SafeSQL(orderNo) & "', " & CLng(userId) & ", " & CDbl(orderAmount) & ", '" & _
          SafeSQL(orderDesc) & "', " & CLng(paymentMethod) & ", 'Pending', '" & _
          SafeSQL(shippingName) & "', '" & SafeSQL(shippingPhone) & "', '" & SafeSQL(shippingAddress) & "', GETDATE())"
    
    ' 检查INSERT操作是否成功
    insertSuccess = ExecuteNonQuery(sql)
    
    If insertSuccess Then
        ' 获取新创建的订单ID
        sql = "SELECT TOP 1 OrderID FROM Orders WHERE OrderNo = '" & SafeSQL(orderNo) & "' ORDER BY CreatedAt DESC"
        Dim rs
        Set rs = ExecuteQuery(sql)
        
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                orderId = rs("OrderID")
            End If
            rs.Close
            Set rs = Nothing
        End If
    Else
        ' INSERT操作失败，记录详细错误
        Dim dbError
        dbError = Session("LastDBError")
        Session("ErrorMessage") = "Database error. SQL: " & sql & "; Error: " & dbError & "; Payment: " & paymentMethod
        SafeCreatePaymentOrder = 0
        Exit Function
    End If
    
    SafeCreatePaymentOrder = orderId
End Function

' 同步订单详情和成分数据到OrderDetails和OrderIngredients表
Sub SyncOrderDetailsAndIngredients(orderId, userId, whereClause)
    On Error Resume Next
    
    DebugLog "===== SyncOrderDetailsAndIngredients 开始 ====="
    DebugLog "orderId=" & orderId & ", userId=" & userId
    DebugLog "whereClause=" & whereClause
    
    ' 查询购物车数据
    Dim cartSql, rsCart
    cartSql = "SELECT c.*, p.ProductName, p.ImageURL, " & _
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
    
    Set rsCart = ExecuteQuery(cartSql)
    
    If rsCart Is Nothing Then
        Exit Sub
    End If
    
    Dim detailId, insertDetailSql, insertNoteSql, insertIngredientSql
    Dim productId, productName, quantity, unitPrice, subtotal
    Dim topNoteId, middleNoteId, baseNoteId, topNoteName, middleNoteName, baseNoteName
    Dim volumeId, volumeName, volumeML, bottleId, bottleName, customLabel
    Dim cartId, rsCartNotes, rsCartNotes2, processedIngredients
    Dim cNoteId, cNoteType, cPercentage, cNoteId2
    Dim rsBaseIngredients, ingredientsStr, ingredientArr, singleIngredient
    Dim rsDetailId
    Dim productRecipeId, rsRecipeCheck, rsRecipeIngredients, ingName
    
    Do While Not rsCart.EOF
        ' 获取购物车数据
        productId = rsCart("ProductID")
        productName = rsCart("ProductName") & ""
        quantity = CLng(rsCart("Quantity"))
        unitPrice = CDbl(rsCart("UnitPrice"))
        subtotal = CDbl(rsCart("SubTotal"))
        topNoteId = rsCart("TopNoteID") & ""
        middleNoteId = rsCart("MiddleNoteID") & ""
        baseNoteId = rsCart("BaseNoteID") & ""
        topNoteName = rsCart("TopNoteName") & ""
        middleNoteName = rsCart("MiddleNoteName") & ""
        baseNoteName = rsCart("BaseNoteName") & ""
        volumeId = rsCart("VolumeID") & ""
        volumeName = rsCart("VolumeName") & ""
        volumeML = rsCart("VolumeML") & ""
        bottleId = rsCart("BottleID") & ""
        bottleName = rsCart("BottleName") & ""
        customLabel = rsCart("CustomLabel") & ""
        
        ' 插入订单详情记录（包含容量、瓶身信息和刻字）
        ' 同时保存VolumeML和VolumeName
        Dim volumeMLValue
        On Error Resume Next
        volumeMLValue = CDbl(volumeML)
        If Err.Number <> 0 Then volumeMLValue = 0 : Err.Clear
        On Error GoTo 0
        insertDetailSql = "INSERT INTO OrderDetails (OrderID, ProductID, ProductName, Quantity, UnitPrice, Subtotal, VolumeName, VolumeML, BottleName, CustomLabel) VALUES (" & _
            orderId & ", " & productId & ", '" & SafeSQL(productName) & "', " & quantity & ", " & unitPrice & ", " & subtotal & ", '" & _
            SafeSQL(volumeName) & "', " & volumeMLValue & ", '" & SafeSQL(bottleName) & "', '" & SafeSQL(customLabel) & "')"
        
        ExecuteNonQuery(insertDetailSql)
        
        ' 获取刚插入的DetailID
        Set rsDetailId = ExecuteQuery("SELECT MAX(DetailID) as maxId FROM OrderDetails WHERE OrderID = " & orderId)
        If Not rsDetailId.EOF Then
            detailId = rsDetailId("maxId")
        Else
            detailId = 0
        End If
        rsDetailId.Close
        Set rsDetailId = Nothing
        
        ' 从CartNoteSelections表读取实际配比数据并插入到OrderDetailNoteSelections表
        ' 保存中文NoteType（前调/中调/后调）到数据库
        cartId = rsCart("CartID")
        DebugLog "  获取CartID: " & cartId
        Set rsCartNotes = ExecuteQuery("SELECT * FROM CartNoteSelections WHERE CartID = " & cartId)
        If Not rsCartNotes Is Nothing Then
            Do While Not rsCartNotes.EOF
                cNoteId = rsCartNotes("NoteID")
                cNoteType = rsCartNotes("NoteType") & ""
                cPercentage = CDbl(rsCartNotes("Percentage"))
                
                ' 直接保存中文NoteType到数据库
                insertNoteSql = "INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage) VALUES (" & _
                    detailId & ", " & cNoteId & ", '" & cNoteType & "', " & CLng(cPercentage) & ")"
                ExecuteNonQuery(insertNoteSql)
                
                rsCartNotes.MoveNext
            Loop
            rsCartNotes.Close
            Set rsCartNotes = Nothing
        End If
        
        ' 从CartNoteSelections表读取实际香调信息并插入到OrderIngredients表
        ' 通过香调→基香→成分的层级关系获取实际成分
        Set processedIngredients = CreateObject("Scripting.Dictionary")
        
        ' 调试日志
        DebugLog "处理购物车 CartID=" & cartId & ", DetailID=" & detailId
        
        ' 新增：检查产品是否有关联配方
        productRecipeId = 0
        Set rsRecipeCheck = ExecuteQuery("SELECT RecipeID FROM Products WHERE ProductID = " & productId)
        If Not rsRecipeCheck Is Nothing Then
            If Not rsRecipeCheck.EOF Then
                If Not IsNull(rsRecipeCheck("RecipeID")) Then
                    productRecipeId = rsRecipeCheck("RecipeID")
                End If
            End If
            rsRecipeCheck.Close
            Set rsRecipeCheck = Nothing
        End If
        
        If productRecipeId > 0 Then
            ' 从 RecipeIngredients 直接读取成分（高效路径）
            Set rsRecipeIngredients = ExecuteQuery("SELECT IngredientName FROM RecipeIngredients WHERE RecipeID = " & CInt(productRecipeId))
            If Not rsRecipeIngredients Is Nothing Then
                Do While Not rsRecipeIngredients.EOF
                    ingName = Trim(rsRecipeIngredients("IngredientName") & "")
                    If ingName <> "" And Not processedIngredients.Exists(ingName) Then
                        processedIngredients.Add ingName, True
                        insertIngredientSql = "INSERT INTO OrderIngredients (OrderID, DetailID, IngredientName, CreatedAt) VALUES (" & _
                            orderId & ", " & detailId & ", '" & SafeSQL(ingName) & "', GETDATE())"
                        DebugLog "    保存配方成分: " & ingName
                        ExecuteNonQuery(insertIngredientSql)
                    End If
                    rsRecipeIngredients.MoveNext
                Loop
                rsRecipeIngredients.Close
                Set rsRecipeIngredients = Nothing
            End If
        Else
            ' 保持现有逻辑：从 NoteIngredients→BaseNotes 逐层查询
            Set rsCartNotes2 = ExecuteQuery("SELECT c.NoteID, c.NoteType, c.Percentage, n.NoteName FROM CartNoteSelections c LEFT JOIN FragranceNotes n ON c.NoteID = n.NoteID WHERE c.CartID = " & cartId)
            If Not rsCartNotes2 Is Nothing Then
                DebugLog "  找到 " & rsCartNotes2.RecordCount & " 个香调选择"
                
                Do While Not rsCartNotes2.EOF
                    cNoteId2 = rsCartNotes2("NoteID")
                    
                    DebugLog "  处理香调 NoteID=" & cNoteId2
                    
                    ' 查询该香调关联的所有基香的成分
                    Set rsBaseIngredients = ExecuteQuery("SELECT b.Ingredients FROM NoteIngredients ni LEFT JOIN BaseNotes b ON ni.BaseNoteID = b.BaseNoteID WHERE ni.NoteID = " & cNoteId2 & " AND b.Ingredients IS NOT NULL AND b.Ingredients <> ''")
                    
                    If Not rsBaseIngredients Is Nothing Then
                        If Not rsBaseIngredients.EOF Then
                            DebugLog "    找到 " & rsBaseIngredients.RecordCount & " 个基香"
                            
                            Do While Not rsBaseIngredients.EOF
                                ingredientsStr = rsBaseIngredients("Ingredients") & ""
                                DebugLog "    基香成分: " & ingredientsStr
                                
                                If ingredientsStr <> "" Then
                                    ' 使用通用成分分割函数，支持多种分隔符（逗号、空格、换行符等）
                                    Dim ingredientDict, ingKey
                                    Set ingredientDict = SplitIngredientsUniversal(ingredientsStr)
                                    
                                    For Each ingKey In ingredientDict.Keys
                                        If Not processedIngredients.Exists(ingKey) Then
                                            processedIngredients.Add ingKey, True
                                            insertIngredientSql = "INSERT INTO OrderIngredients (OrderID, DetailID, IngredientName, CreatedAt) VALUES (" & _
                                                orderId & ", " & detailId & ", '" & SafeSQL(ingKey) & "', GETDATE())"
                                            DebugLog "    执行SQL: " & insertIngredientSql
                                            ExecuteNonQuery(insertIngredientSql)
                                        End If
                                    Next
                                    Set ingredientDict = Nothing
                                End If
                                rsBaseIngredients.MoveNext
                            Loop
                        Else
                            DebugLog "    该香调没有基香成分"
                        End If
                        rsBaseIngredients.Close
                        Set rsBaseIngredients = Nothing
                    Else
                        DebugLog "    查询基香失败"
                    End If
                    
                    rsCartNotes2.MoveNext
                Loop
                
                DebugLog "  完成处理，共保存 " & processedIngredients.Count & " 个成分"
                
                rsCartNotes2.Close
                Set rsCartNotes2 = Nothing
            Else
                DebugLog "  没有找到香调选择"
            End If
        End If
        
        ' ===== 品牌定香商品的BaseIngredients处理 =====
        ' 检查商品是否有BaseIngredients字段（品牌定香商品）
        Dim rsProductIngr, productBaseIngr, fixedIngDict, fixedIngKey
        Set rsProductIngr = ExecuteQuery("SELECT BaseIngredients, ProductType FROM Products WHERE ProductID = " & productId)
        If Not rsProductIngr Is Nothing Then
            If Not rsProductIngr.EOF Then
                productBaseIngr = Trim(rsProductIngr("BaseIngredients") & "")
                If productBaseIngr <> "" Then
                    DebugLog "  品牌定香商品BaseIngredients: " & productBaseIngr
                    
                    ' 初始化processedIngredients（如果之前没有创建）
                    If processedIngredients Is Nothing Then
                        Set processedIngredients = CreateObject("Scripting.Dictionary")
                    End If
                    
                    ' 使用通用成分分割函数
                    Set fixedIngDict = SplitIngredientsUniversal(productBaseIngr)
                    For Each fixedIngKey In fixedIngDict.Keys
                        If Not processedIngredients.Exists(fixedIngKey) Then
                            processedIngredients.Add fixedIngKey, True
                            insertIngredientSql = "INSERT INTO OrderIngredients (OrderID, DetailID, IngredientName, CreatedAt) VALUES (" & _
                                orderId & ", " & detailId & ", '" & SafeSQL(fixedIngKey) & "', GETDATE())"
                            DebugLog "    保存品牌定香成分: " & fixedIngKey
                            ExecuteNonQuery(insertIngredientSql)
                        End If
                    Next
                    Set fixedIngDict = Nothing
                    DebugLog "  品牌定香成分处理完成"
                End If
            End If
            rsProductIngr.Close
            Set rsProductIngr = Nothing
        End If
        
        ' 检查是否有错误
        If Err.Number <> 0 Then
            DebugLog "  错误: " & Err.Description & " (" & Err.Number & ")"
            Err.Clear
        End If
        
        ' 注意：基础溶剂成分（乙醇、蒸馏水）应该通过基香配置来添加
        ' 而不是硬编码，这样可以灵活配置不同产品的基础成分
        
        ' ==================== 库存扣减 ====================
        ' 检查是否启用库存管理
        Dim enableInventoryCheck
        enableInventoryCheck = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableInventoryCheck'")
        If IsNull(enableInventoryCheck) Then enableInventoryCheck = "1"
        
        If enableInventoryCheck = "1" Then
            ' 扣减香调库存
            Call DeductNoteInventory(topNoteId, quantity, orderId, productName)
            Call DeductNoteInventory(middleNoteId, quantity, orderId, productName)
            Call DeductNoteInventory(baseNoteId, quantity, orderId, productName)
        End If
        ' ==================== 库存扣减结束 ====================
        
        rsCart.MoveNext
    Loop
    
    rsCart.Close
    Set rsCart = Nothing
    
    DebugLog "===== SyncOrderDetailsAndIngredients 结束 ====="
    
    On Error GoTo 0
End Sub

' 扣减香调库存的辅助函数
Sub DeductNoteInventory(noteIdList, quantity, orderId, productName)
    On Error Resume Next
    
    If noteIdList = "" Then Exit Sub
    
    Dim arr, i, nId
    arr = Split(noteIdList & "", ",")
    
    For i = 0 To UBound(arr)
        nId = Trim(arr(i))
        If IsNumeric(nId) Then
            ' 扣减库存
            Dim deductSql
            deductSql = "UPDATE NoteInventory SET " & _
                "StockQuantity = StockQuantity - " & CInt(quantity) & ", " & _
                "UpdatedAt = GETDATE() " & _
                "WHERE NoteID = " & CInt(nId) & " AND StockQuantity >= " & CInt(quantity)
            ExecuteNonQuery(deductSql)
            
            ' 记录库存变动
            Dim transSql
            transSql = "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, ReferenceOrderID, ReferenceType, Notes, CreatedAt) VALUES (" & _
                CInt(nId) & ", -" & CInt(quantity) & ", '订单消耗', " & orderId & ", '订单" & orderId & "-" & SafeSQL(productName) & "', '', GETDATE())"
            ExecuteNonQuery(transSql)
        End If
    Next
    
    On Error GoTo 0
End Sub

' 根据ID获取地区名称的函数
Function GetAreaNameById(areaId)
    If Not IsNumeric(areaId) Or areaId = "" Then
        GetAreaNameById = ""
        Exit Function
    End If
    
    Dim sql, rs
    sql = "SELECT AreaName FROM Areas WHERE AreaID = " & areaId
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            GetAreaNameById = rs("AreaName")
        Else
            GetAreaNameById = ""
        End If
        rs.Close
        Set rs = Nothing
    Else
        GetAreaNameById = ""
    End If
End Function

' 处理支付请求
Dim paymentMethod, orderId, paymentResult, realName, phone, address, selectedAddressId, paymentError, cartSnapshotSql, rsCartSnapshot, productDetails, itemCount, itemDesc, noteInfo

' 初始化变量
paymentResult = False
realName = ""
phone = ""
address = ""
orderId = 0
' 首先尝试从表单获取，如果为空则从 URL参数获取
paymentMethod = Request.Form("payment_method")
If paymentMethod = "" Then
    paymentMethod = Request.QueryString("payment_method")
End If

' 将支付方式转换为整数类型，以便与常量比较
If paymentMethod <> "" And IsNumeric(paymentMethod) Then
    paymentMethod = CLng(paymentMethod)
End If

selectedAddressId = Request.Form("selectedAddress")

' 初始化变量
realName = ""
phone = ""
address = ""
orderId = 0

If paymentMethod <> "" Then
    ' 计算订单总金额（商品金额 + 刻字费用 + 运费）
    Dim finalAmount
    If grandTotal >= FREE_SHIPPING_AMOUNT Then
        finalAmount = grandTotal
    Else
        finalAmount = grandTotal + SHIPPING_FEE
    End If
    
    ' 判断是否是已存在订单模式
    If isExistingOrder Then
        ' 已存在订单模式，直接使用现有订单ID
        orderId = existingOrderId
        ' 从订单中获取收货信息
        realName = rsExistingOrder("ShippingName")
        phone = rsExistingOrder("ShippingPhone")
        address = rsExistingOrder("ShippingAddress")
        ' 关闭记录集
        rsExistingOrder.Close
        Set rsExistingOrder = Nothing
    Else
        ' 新订单模式，需要创建订单
    
    ' 检查是否选择了现有地址
    If selectedAddressId <> "" And selectedAddressId <> "new" And IsNumeric(selectedAddressId) Then
        ' 使用选择的地址
        Dim rsSelectedAddress
        Set rsSelectedAddress = ExecuteQuery("SELECT * FROM UserAddresses WHERE AddressID = " & selectedAddressId & " AND UserID = " & userId)
            
        If Not rsSelectedAddress Is Nothing And Not rsSelectedAddress.EOF Then
            realName = rsSelectedAddress("Consignee")
            phone = rsSelectedAddress("Phone")
            address = rsSelectedAddress("Province") & rsSelectedAddress("City") & rsSelectedAddress("District") & rsSelectedAddress("Address")
                        
            ' 重新查询购物车数据以获取商品信息，用于订单快照
            productDetails = ""
                        
            cartSnapshotSql = "SELECT c.*, p.ProductName, p.ImageURL, " & _
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
                    ' 构建包含完整定制信息的商品描述
                    itemDesc = rsCartSnapshot("ProductName") & " x" & rsCartSnapshot("Quantity") & " (" & FormatMoney(rsCartSnapshot("SubTotal")) & ")"
                    
                    ' 添加香水定制信息
                    noteInfo = ""
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
                    
                    ' 添加容量和瓶身信息
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
                    
                    ' 组合完整商品描述
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
            
            ' 如果没有商品信息，使用默认描述
            If productDetails = "" Then
                productDetails = "香水订单 (" & cartCount & " 件商品, 总计: " & FormatMoney(cartTotal) & ")"
            Else
                productDetails = "详情: " & productDetails
            End If
                
            ' 创建订单
            orderId = SafeCreatePaymentOrder(userId, finalAmount, productDetails, paymentMethod, realName, phone, address)
                            
            If orderId > 0 Then
                ' 更新用户收货信息
                Dim updateUserInfoSql
                updateUserInfoSql = "UPDATE Users SET FullName = '" & SafeSQL(realName) & "', Phone = '" & SafeSQL(phone) & "', Address = '" & SafeSQL(address) & "' WHERE UserID = " & userId
                Call ExecuteNonQuery(updateUserInfoSql)
                
                ' 同步订单详情和成分数据到OrderDetails和OrderIngredients表
                Call SyncOrderDetailsAndIngredients(orderId, userId, whereClause)
                
                ' 成本自动传导：更新订单中各产品的BOM成本和订单总利润
                Call CE_UpdateOrderCosts(orderId)
            Else
                ' 订单创建失败 - 提供更详细的错误信息
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
        ' 使用新地址 - 从弹窗表单获取数据
        realName = SafeSQL(Request.Form("realName"))
        phone = SafeSQL(Request.Form("phone"))
        provinceName = SafeSQL(Trim(Request.Form("province")))
        cityName = SafeSQL(Trim(Request.Form("city")))
        districtName = SafeSQL(Trim(Request.Form("district")))
        detailAddress = SafeSQL(Request.Form("address"))
            
        address = provinceName & cityName & districtName & detailAddress
            
        ' 验证收货信息
        If realName = "" Or phone = "" Or provinceName = "" Or cityName = "" Or districtName = "" Or detailAddress = "" Then
            Session("ErrorMessage") = "请填写完整的收货信息"
        Else
            ' 重新查询购物车数据以获取商品信息，用于订单快照
            productDetails = ""
            
            cartSnapshotSql = "SELECT c.*, p.ProductName, p.ImageURL, " & _
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
                    ' 构建包含完整定制信息的商品描述
                    itemDesc = rsCartSnapshot("ProductName") & " x" & rsCartSnapshot("Quantity") & " (" & FormatMoney(rsCartSnapshot("SubTotal")) & ")"
                    
                    ' 添加香水定制信息
                    noteInfo = ""
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
                    
                    ' 添加容量和瓶身信息
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
                    
                    ' 组合完整商品描述
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
            
            ' 如果没有商品信息，使用默认描述
            If productDetails = "" Then
                productDetails = "Perfume Order (" & cartCount & " items, Total: " & FormatMoney(cartTotal) & ")"
            Else
                productDetails = "Details: " & productDetails
            End If
            
            ' 创建订单
            orderId = SafeCreatePaymentOrder(userId, finalAmount, productDetails, paymentMethod, realName, phone, address)
            
            If orderId > 0 Then
                ' 更新用户收货信息
                updateUserInfoSql = "UPDATE Users SET FullName = '" & SafeSQL(realName) & "', Phone = '" & SafeSQL(phone) & "', Address = '" & SafeSQL(address) & "' WHERE UserID = " & userId
                Call ExecuteNonQuery(updateUserInfoSql)
                
                ' 同步订单详情和成分数据到OrderDetails和OrderIngredients表
                Call SyncOrderDetailsAndIngredients(orderId, userId, whereClause)
                
                ' 成本自动传导：更新订单中各产品的BOM成本和订单总利润
                Call CE_UpdateOrderCosts(orderId)
            Else
                ' 订单创建失败
                If Session("ErrorMessage") = "" Then
                    Session("ErrorMessage") = "订单创建失败，请检查购物车和支付方式。Payment: '" & paymentMethod & "'"
                End If
            End If
        End If
    End If
    End If  ' 结束新订单模式的判断
    
    ' 支付处理（无论是新订单还是已存在订单）
    If orderId > 0 Then
                
                ' 支付方式处理
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
                        ' 货到付款，使用专门的处理函数
                        paymentResult = ProcessCashOnDelivery(orderId)
                        If Err.Number <> 0 Then
                            Session("ErrorMessage") = "COD error: " & Err.Description
                            paymentResult = False
                        End If
                        On Error Goto 0
                End Select
                
                If paymentResult Then
                    ' 清空购物车（只在新订单模式下）
                    If Not isExistingOrder Then
                        ' 使用专门构建的deleteClause（不带别名）
                        Dim actualDeleteClause2
                        If IsEmpty(deleteClause) Then
                            actualDeleteClause2 = whereClause
                            actualDeleteClause2 = Replace(actualDeleteClause2, "c.", "")
                        Else
                            actualDeleteClause2 = deleteClause
                        End If
                        clearSql = "DELETE FROM Cart WHERE " & actualDeleteClause2
                        Session("DebugCartClearSQL") = clearSql
                        clearResult = ExecuteNonQuery(clearSql)
                        Session("DebugCartClearResult") = clearResult
                        If Not clearResult Then
                            Session("ErrorMessage") = "Cart clear failed: " & Session("LastDBError")
                        End If
                    End If
                    
                    ' 跳转到订单成功页面
                    Response.Redirect "/order_success.asp?order_id=" & orderId
                    Response.End
                Else
                    ' 支付处理失败
                    If Session("ErrorMessage") = "" Then
                        Session("ErrorMessage") = "Payment processing failed"
                    End If
                End If
    End If  ' 结束If orderId > 0 Then
End If  ' 结束If paymentMethod <> "" Then

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
<!--#include file="includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/cart.asp">购物车</a>
        <span class="separator">/</span>
        <span>结算</span>
    </div>
</div>

<div class="container">
    <div class="checkout-page">
        <h1 class="page-title"><i class="fas fa-credit-card"></i> 订单结算</h1>
        
        <% If Session("ErrorMessage") <> "" Then %>
        <div class="alert alert-error">
            <%= Session("ErrorMessage") %>
            <% Session("ErrorMessage") = "" %>
        </div>
        <% End If %>
        
        <div class="checkout-content">
            <div class="checkout-items">
                <h3>订单商品</h3>
                
                <%
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
                
                If Not rsCart Is Nothing Then
                    Do While Not rsCart.EOF
                %>
                <div class="checkout-item">
                    <div class="item-image">
                        <img src="<%= rsCart("ImageURL") %>" alt="<%= HTMLEncode(rsCart("ProductName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                    </div>
                    <div class="item-details">
                        <h4><%= HTMLEncode(rsCart("ProductName")) %></h4>
                        <div class="item-attributes">
                            <% If Not IsNull(rsCart("TopNoteName")) Then %>
                            <span><i class="fas fa-wind"></i> 前调: <%= HTMLEncode(rsCart("TopNoteName")) %></span>
                            <% End If %>
                            <% If Not IsNull(rsCart("MiddleNoteName")) Then %>
                            <span><i class="fas fa-heart"></i> 中调: <%= HTMLEncode(rsCart("MiddleNoteName")) %></span>
                            <% End If %>
                            <% If Not IsNull(rsCart("BaseNoteName")) Then %>
                            <span><i class="fas fa-moon"></i> 后调: <%= HTMLEncode(rsCart("BaseNoteName")) %></span>
                            <% End If %>
                            <% If Not IsNull(rsCart("VolumeName")) Then %>
                            <span><i class="fas fa-tint"></i> 容量: <%= rsCart("VolumeML") %>ml (<%= HTMLEncode(rsCart("VolumeName")) %>)</span>
                            <% End If %>
                            <% If Not IsNull(rsCart("BottleName")) Then %>
                            <span><i class="fas fa-wine-bottle"></i> 瓶身: <%= HTMLEncode(rsCart("BottleName")) %></span>
                            <% End If %>
                            <% If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" Then %>
                            <span><i class="fas fa-pen-fancy"></i> 刻字: <%= HTMLEncode(rsCart("CustomLabel")) %></span>
                            <% End If %>
                            <% 
                            ' 显示刻字费用
                            Dim checkoutItemEngravingPrice
                            checkoutItemEngravingPrice = 0
                            On Error Resume Next
                            checkoutItemEngravingPrice = CDbl(rsCart("EngravingPrice"))
                            If Err.Number <> 0 Then checkoutItemEngravingPrice = 0
                            On Error GoTo 0
                            If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" And checkoutItemEngravingPrice > 0 Then 
                            %>
                            <span style="color:#e91e63;"><i class="fas fa-tag"></i> 刻字费用: <%= FormatMoney(checkoutItemEngravingPrice) %></span>
                            <% End If %>
                        </div>
                    </div>
                    <div class="item-quantity">
                        × <%= rsCart("Quantity") %>
                    </div>
                    <div class="item-price">
                        <%= FormatMoney(rsCart("SubTotal")) %>
                    </div>
                </div>
                <%
                    rsCart.MoveNext
                    Loop
                    rsCart.Close
                    Set rsCart = Nothing
                End If
                %>
            </div>
            
            <div class="checkout-summary">
                <h3>订单摘要</h3>
                
                <div class="summary-row">
                    <span>商品金额:</span>
                    <span><%= FormatMoney(cartTotal) %></span>
                </div>
                
                <% If totalEngravingFee > 0 Then %>
                <div class="summary-row">
                    <span>刻字费用:</span>
                    <span><%= FormatMoney(totalEngravingFee) %></span>
                </div>
                <% End If %>
                
                <div class="summary-row">
                    <span>运费:</span>
                    <% If grandTotal >= FREE_SHIPPING_AMOUNT Then %>
                    <span>免运费</span>
                    <% Else %>
                    <span><%= FormatMoney(SHIPPING_FEE) %></span>
                    <% End If %>
                </div>
                
                <div class="summary-divider"></div>
                
                <div class="summary-total">
                    <span>应付总额:</span>
                    <% If grandTotal >= FREE_SHIPPING_AMOUNT Then %>
                    <span class="total-amount"><%= FormatMoney(grandTotal) %></span>
                    <% Else %>
                    <span class="total-amount"><%= FormatMoney(grandTotal + SHIPPING_FEE) %></span>
                    <% End If %>
                </div>
                
                <form method="post" id="paymentForm">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="cart_ids" value="<%= cartIdList %>">
                <h3>收货信息</h3>
                
                <div class="form-group">
                    <label>选择收货地址</label>
                    <div class="address-selector">
                        <select name="selectedAddress" id="selectedAddress" class="form-control" onchange="loadAddressDetails()">
                            <option value="">-- 选择已有地址或添加新地址 --</option>
                            <% 
                            ' 获取用户地址列表
                            Dim rsUserAddresses, addrId, addrConsignee, addrPhone, addrProvince, addrCity, addrDistrict, addrDetail, addrIsDefault
                            Dim selectedAddressParam
                            selectedAddressParam = Request.QueryString("selected_address")
                            Set rsUserAddresses = ExecuteQuery("SELECT * FROM UserAddresses WHERE UserID = " & userId & " ORDER BY IsDefault DESC, CreatedAt DESC")
                            If Not rsUserAddresses Is Nothing Then
                                If Not rsUserAddresses.EOF Then
                                    Do While Not rsUserAddresses.EOF
                                        addrId = rsUserAddresses("AddressID")
                                        addrConsignee = rsUserAddresses("Consignee")
                                        addrPhone = rsUserAddresses("Phone")
                                        addrProvince = rsUserAddresses("Province")
                                        addrCity = rsUserAddresses("City")
                                        addrDistrict = rsUserAddresses("District")
                                        addrDetail = rsUserAddresses("Address")
                                        addrIsDefault = rsUserAddresses("IsDefault")
                                        %>
                                    <option value="<%= addrId %>"<% If (addrIsDefault <> 0) Or (selectedAddressParam <> "" And CLng(selectedAddressParam) = CLng(addrId)) Then Response.Write " selected" End If %>><%= HTMLEncode(addrConsignee) %> <%= HTMLEncode(addrPhone) %> <%= HTMLEncode(addrProvince) %> <%= HTMLEncode(addrCity) %> <%= HTMLEncode(addrDistrict) %> <%= HTMLEncode(addrDetail) %></option>
                                    <%
                                        rsUserAddresses.MoveNext
                                    Loop
                                End If
                                rsUserAddresses.Close
                                Set rsUserAddresses = Nothing
                            End If
                            %>
                            <option value="new"<% If selectedAddressParam <> "" And selectedAddressParam = "new" Then Response.Write " selected" End If %>>+ 添加新地址</option>
                        </select>
                        <button type="button" class="btn btn-secondary" onclick="showAddressForm()" style="margin-top: 10px;">
                            <i class="fas fa-plus"></i> 新增收货地址
                        </button>
                    </div>
                </div>
                
                <div id="selectedAddressDisplay" style="display:block;">
                    <!-- 默认显示用户信息，如果没有地址被选中 -->
                    <% If userRealName <> "" Then %>
                    <div class="selected-address-info">
                        <p><strong>当前地址：</strong><%= HTMLEncode(userRealName) %> <%= HTMLEncode(userPhone) %> <%= HTMLEncode(userAddress) %></p>
                    </div>
                    <% End If %>
                </div>
                
                <h3>支付方式</h3>
                
                    <div class="payment-methods">
                        <% If enableCOD = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_COD %>" checked>
                                <span class="radio-text">货到付款</span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enableWechat = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_WECHAT %>" <% If enableCOD <> "1" Then Response.Write "checked" End If %>>
                                <span class="radio-text">微信支付</span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enableAlipay = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_ALIPAY %>">
                                <span class="radio-text">支付宝</span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enablePaypal = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_PAYPAL %>">
                                <span class="radio-text">PayPal</span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enableBankTransfer = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="5">
                                <span class="radio-text">银行转账</span>
                            </label>
                        </div>
                        <% End If %>
                    </div>
                    
                    <% If enableCOD <> "1" And enableWechat <> "1" And enableAlipay <> "1" And enablePaypal <> "1" And enableBankTransfer <> "1" Then %>
                    <div class="alert alert-warning">
                        <i class="fas fa-exclamation-triangle"></i> 暂无可用的支付方式，请联系客服。
                    </div>
                    <% End If %>
                    
                    <button type="submit" class="btn btn-primary btn-lg btn-block">
                        <i class="fas fa-check"></i> 确认订单并支付
                    </button>
                </form>
                    
                <!-- 添加/编辑地址弹窗 -->
                <div class="modal" id="addressModal">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h3 id="modalTitle">新增收货地址</h3>
                            <span class="close" onclick="closeAddressForm()">&times;</span>
                        </div>
                        <div class="modal-body">
                            <form id="addressForm" method="post" action="checkout.asp">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" id="formAction" name="action" value="add">
                                <input type="hidden" id="formAddressId" name="addressId" value="">
                                <input type="hidden" name="cart_ids" value="<%= cartIdList %>">
                                <input type="hidden" name="payment_method" value="" id="addressFormPaymentMethod">
                                <div class="form-row">
                                    <div class="form-group">
                                        <label for="consignee">收货人姓名 *</label>
                                        <input type="text" id="consignee" name="realName" required>
                                    </div>
                                    <div class="form-group">
                                        <label for="phone">联系电话 *</label>
                                        <input type="tel" id="phone" name="phone" required>
                                    </div>
                                </div>
                                    
                                <div class="form-group">
                                    <label for="province">所在地区 *</label>
                                    <select id="province" name="province" required onchange="updateCities()">
                                        <option value="">请选择省份</option>
                                        <option value="北京市">北京市</option>
                                        <option value="上海市">上海市</option>
                                        <option value="天津市">天津市</option>
                                        <option value="重庆市">重庆市</option>
                                        <option value="河北省">河北省</option>
                                        <option value="山西省">山西省</option>
                                        <option value="辽宁省">辽宁省</option>
                                        <option value="吉林省">吉林省</option>
                                        <option value="黑龙江省">黑龙江省</option>
                                        <option value="江苏省">江苏省</option>
                                        <option value="浙江省">浙江省</option>
                                        <option value="安徽省">安徽省</option>
                                        <option value="福建省">福建省</option>
                                        <option value="江西省">江西省</option>
                                        <option value="山东省">山东省</option>
                                        <option value="河南省">河南省</option>
                                        <option value="湖北省">湖北省</option>
                                        <option value="湖南省">湖南省</option>
                                        <option value="广东省">广东省</option>
                                        <option value="海南省">海南省</option>
                                        <option value="四川省">四川省</option>
                                        <option value="贵州省">贵州省</option>
                                        <option value="云南省">云南省</option>
                                        <option value="陕西省">陕西省</option>
                                        <option value="甘肃省">甘肃省</option>
                                        <option value="青海省">青海省</option>
                                        <option value="台湾省">台湾省</option>
                                        <option value="内蒙古自治区">内蒙古自治区</option>
                                        <option value="广西壮族自治区">广西壮族自治区</option>
                                        <option value="西藏自治区">西藏自治区</option>
                                        <option value="宁夏回族自治区">宁夏回族自治区</option>
                                        <option value="新疆维吾尔自治区">新疆维吾尔自治区</option>
                                        <option value="香港特别行政区">香港特别行政区</option>
                                        <option value="澳门特别行政区">澳门特别行政区</option>
                                    </select>
                                    <select id="city" name="city" required onchange="updateDistricts()">
                                        <option value="">请选择城市</option>
                                    </select>
                                    <select id="district" name="district" required>
                                        <option value="">请选择区县</option>
                                    </select>
                                </div>
                                    
                                <div class="form-group">
                                    <label for="address">详细地址 *</label>
                                    <input type="text" id="address" name="address" placeholder="请输入详细地址，如街道、门牌号等" required>
                                </div>
                                    
                                <div class="form-group">
                                    <label class="checkbox-label">
                                        <input type="checkbox" id="isDefault" name="isDefault" value="1">
                                        设为默认地址
                                    </label>
                                </div>
                                    
                                <div class="form-actions">
                                    <button type="submit" class="btn btn-primary">保存地址</button>
                                    <button type="button" class="btn btn-text" onclick="closeAddressForm()">取消</button>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>
</div>

<script src="/js/area_data.js"></script>
<script>
// 省市区数据已移至外部文件 area_data.js

/*var addressData = {
    "北京市": ["东城区", "西城区", "朝阳区", "丰台区", "石景山区", "海淀区", "门头沟区", "房山区", "通州区", "顺义区", "昌平区", "大兴区", "怀柔区", "平谷区", "密云区", "延庆区"],
    "上海市": ["黄浦区", "徐汇区", "长宁区", "静安区", "普陀区", "虹口区", "杨浦区", "闵行区", "宝山区", "嘉定区", "浦东新区", "金山区", "松江区", "青浦区", "奉贤区", "崇明区"],
    "天津市": ["和平区", "河东区", "河西区", "南开区", "河北区", "红桥区", "东丽区", "西青区", "津南区", "北辰区", "武清区", "宝坻区", "滨海新区", "宁河区", "静海区", "蓟州区"],
    "重庆市": ["万州区", "涪陵区", "渝中区", "大渡口区", "江北区", "沙坪坝区", "九龙坡区", "南岸区", "北碚区", "綦江区", "大足区", "渝北区", "巴南区", "黔江区", "长寿区", "江津区", "合川区", "永川区", "南川区", "璧山区", "铜梁区", "潼南区", "荣昌区", "开州区", "梁平区", "武隆区"],
    "河北省": ["石家庄市", "唐山市", "秦皇岛市", "邯郸市", "邢台市", "保定市", "张家口市", "承德市", "沧州市", "廊坊市", "衡水市"],
    "山西省": ["太原市", "大同市", "阳泉市", "长治市", "晋城市", "朔州市", "晋中市", "运城市", "忻州市", "临汾市", "吕梁市"],
    "辽宁省": ["沈阳市", "大连市", "鞍山市", "抚顺市", "本溪市", "丹东市", "锦州市", "营口市", "阜新市", "辽阳市", "盘锦市", "铁岭市", "朝阳市", "葫芦岛市"],
    "吉林省": ["长春市", "吉林市", "四平市", "辽源市", "通化市", "白山市", "松原市", "白城市", "延边朝鲜族自治州"],
    "黑龙江省": ["哈尔滨市", "齐齐哈尔市", "鸡西市", "鹤岗市", "双鸭山市", "大庆市", "伊春市", "佳木斯市", "七台河市", "牡丹江市", "黑河市", "绥化市", "大兴安岭地区"],
    "江苏省": ["南京市", "无锡市", "徐州市", "常州市", "苏州市", "南通市", "连云港市", "淮安市", "盐城市", "扬州市", "镇江市", "泰州市", "宿迁市"],
    "浙江省": ["杭州市", "宁波市", "温州市", "嘉兴市", "湖州市", "绍兴市", "金华市", "衢州市", "舟山市", "台州市", "丽水市"],
    "安徽省": ["合肥市", "芜湖市", "蚌埠市", "淮南市", "马鞍山市", "淮北市", "铜陵市", "安庆市", "黄山市", "滁州市", "阜阳市", "宿州市", "六安市", "亳州市", "池州市", "宣城市"],
    "福建省": ["福州市", "厦门市", "莆田市", "三明市", "泉州市", "漳州市", "南平市", "龙岩市", "宁德市"],
    "江西省": ["南昌市", "景德镇市", "萍乡市", "九江市", "新余市", "鹰潭市", "赣州市", "吉安市", "宜春市", "抚州市", "上饶市"],
    "山东省": ["济南市", "青岛市", "淄博市", "枣庄市", "东营市", "烟台市", "潍坊市", "济宁市", "泰安市", "威海市", "日照市", "临沂市", "德州市", "聊城市", "滨州市", "菏泽市"],
    "河南省": ["郑州市", "开封市", "洛阳市", "平顶山市", "安阳市", "鹤壁市", "新乡市", "焦作市", "濮阳市", "许昌市", "漯河市", "三门峡市", "南阳市", "商丘市", "信阳市", "周口市", "驻马店市", "济源市"],
    "湖北省": ["武汉市", "黄石市", "十堰市", "宜昌市", "襄阳市", "鄂州市", "荆门市", "孝感市", "荆州市", "黄冈市", "咸宁市", "随州市", "恩施土家族苗族自治州"],
    "湖南省": ["长沙市", "株洲市", "湘潭市", "衡阳市", "邵阳市", "岳阳市", "常德市", "张家界市", "益阳市", "郴州市", "永州市", "怀化市", "娄底市", "湘西土家族苗族自治州"],
    "广东省": ["广州市", "韶关市", "深圳市", "珠海市", "汕头市", "佛山市", "江门市", "湛江市", "茂名市", "肇庆市", "惠州市", "梅州市", "汕尾市", "河源市", "阳江市", "清远市", "东莞市", "中山市", "潮州市", "揭阳市", "云浮市"],
    "海南省": ["海口市", "三亚市", "三沙市", "儋州市"],
    "四川省": ["成都市", "自贡市", "攀枝花市", "泸州市", "德阳市", "绵阳市", "广元市", "遂宁市", "内江市", "乐山市", "南充市", "眉山市", "宜宾市", "广安市", "达州市", "雅安市", "巴中市", "资阳市", "阿坝藏族羌族自治州", "甘孜藏族自治州", "凉山彝族自治州"],
    "贵州省": ["贵阳市", "六盘水市", "遵义市", "安顺市", "毕节市", "铜仁市", "黔西南布依族苗族自治州", "黔东南苗族侗族自治州", "黔南布依族苗族自治州"],
    "云南省": ["昆明市", "曲靖市", "玉溪市", "保山市", "昭通市", "丽江市", "普洱市", "临沧市", "楚雄彝族自治州", "红河哈尼族彝族自治州", "文山壮族苗族自治州", "西双版纳傣族自治州", "大理白族自治州", "德宏傣族景颇族自治州", "怒江傈僳族自治州", "迪庆藏族自治州"],
    "陕西省": ["西安市", "铜川市", "宝鸡市", "咸阳市", "渭南市", "延安市", "汉中市", "榆林市", "安康市", "商洛市"],
    "甘肃省": ["兰州市", "嘉峪关市", "金昌市", "白银市", "天水市", "武威市", "张掖市", "平凉市", "酒泉市", "庆阳市", "定西市", "陇南市", "临夏回族自治州", "甘南藏族自治州"],
    "青海省": ["西宁市", "海东市", "海北藏族自治州", "黄南藏族自治州", "海南藏族自治州", "果洛藏族自治州", "玉树藏族自治州", "海西蒙古族藏族自治州"],
    "台湾省": ["台北市", "新北市", "桃园市", "台中市", "台南市", "高雄市"],
    "内蒙古自治区": ["呼和浩特市", "包头市", "乌海市", "赤峰市", "通辽市", "鄂尔多斯市", "呼伦贝尔市", "巴彦淖尔市", "乌兰察布市", "兴安盟", "锡林郭勒盟", "阿拉善盟"],
    "广西壮族自治区": ["南宁市", "柳州市", "桂林市", "梧州市", "北海市", "防城港市", "钦州市", "贵港市", "玉林市", "百色市", "贺州市", "河池市", "来宾市", "崇左市"],
    "西藏自治区": ["拉萨市", "日喀则市", "昌都市", "林芝市", "山南市", "那曲市", "阿里地区"],
    "宁夏回族自治区": ["银川市", "石嘴山市", "吴忠市", "固原市", "中卫市"],
    "新疆维吾尔自治区": ["乌鲁木齐市", "克拉玛依市", "吐鲁番市", "哈密市", "昌吉回族自治州", "博尔塔拉蒙古自治州", "巴音郭楞蒙古自治州", "阿克苏地区", "克孜勒苏柯尔克孜自治州", "喀什地区", "和田地区", "伊犁哈萨克自治州", "塔城地区", "阿勒泰地区"],
    "香港特别行政区": ["香港岛", "九龙", "新界"],
    "澳门特别行政区": ["澳门半岛", "氹仔", "路环"]
};

// 区县数据
var districtData = {
    "北京市": ["东城区", "西城区", "朝阳区", "丰台区", "石景山区", "海淀区", "门头沟区", "房山区", "通州区", "顺义区", "昌平区", "大兴区", "怀柔区", "平谷区", "密云区", "延庆区"],
    "上海市": ["黄浦区", "徐汇区", "长宁区", "静安区", "普陀区", "虹口区", "杨浦区", "闵行区", "宝山区", "嘉定区", "浦东新区", "金山区", "松江区", "青浦区", "奉贤区", "崇明区"],
    "天津市": ["和平区", "河东区", "河西区", "南开区", "河北区", "红桥区", "东丽区", "西青区", "津南区", "北辰区", "武清区", "宝坻区", "滨海新区", "宁河区", "静海区", "蓟州区"],
    "重庆市": ["万州区", "涪陵区", "渝中区", "大渡口区", "江北区", "沙坪坝区", "九龙坡区", "南岸区", "北碚区", "綦江区", "大足区", "渝北区", "巴南区", "黔江区", "长寿区", "江津区", "合川区", "永川区", "南川区", "璧山区", "铜梁区", "潼南区", "荣昌区", "开州区", "梁平区", "武隆区"],
    "石家庄市": ["长安区", "桥西区", "新华区", "井陉矿区", "裕华区", "藁城区", "鹿泉区", "栾城区", "井陉县", "正定县", "行唐县", "灵寿县", "高邑县", "深泽县", "赞皇县", "无极县", "平山县", "元氏县", "赵县", "晋州市", "新乐市"],
    "唐山市": ["路南区", "路北区", "古冶区", "开平区", "丰南区", "丰润区", "曹妃甸区", "滦南县", "乐亭县", "迁西县", "玉田县", "遵化市", "迁安市", "滦州市"],
    "秦皇岛市": ["海港区", "山海关区", "北戴河区", "抚宁区", "青龙满族自治县", "昌黎县", "卢龙县"],
    "邯郸市": ["邯山区", "丛台区", "复兴区", "峰峰矿区", "肥乡区", "永年区", "临漳县", "成安县", "大名县", "涉县", "磁县", "邱县", "鸡泽县", "广平县", "馆陶县", "魏县", "曲周县", "武安市"],
    "广州市": ["荔湾区", "越秀区", "海珠区", "天河区", "白云区", "黄埔区", "番禺区", "花都区", "南沙区", "从化区", "增城区"],
    "深圳市": ["罗湖区", "福田区", "南山区", "宝安区", "龙岗区", "盐田区", "龙华区", "坪山区", "光明区"],
    "成都市": ["锦江区", "青羊区", "金牛区", "武侯区", "成华区", "龙泉驿区", "青白江区", "新都区", "温江区", "双流区", "郫都区", "金堂县", "大邑县", "蒲江县", "新津区", "都江堰市", "彭州市", "郛崃市", "崇州市", "简阳市"],
    "杭州市": ["上城区", "拱墅区", "西湖区", "滨江区", "萧山区", "余杭区", "临平区", "钱塘区", "富阳区", "临安区", "桐庐县", "淳安县", "建德市"],
    "南京市": ["玄武区", "秦淮区", "建邺区", "鼓楼区", "浦口区", "栖霞区", "雨花台区", "江宁区", "六合区", "溧水区", "高淳区"],
    "武汉市": ["江岸区", "江汉区", "硚口区", "汉阳区", "武昌区", "青山区", "洪山区", "东西湖区", "蔡甸区", "江夏区", "黄陂区", "新洲区"],
    "西安市": ["新城区", "碑林区", "莲湖区", "灞桥区", "未央区", "雁塔区", "阎良区", "临潼区", "长安区", "高陵区", "鄠邑区", "蓝田县", "周至县"],
    "南京市": ["玄武区", "秦淮区", "建邺区", "鼓楼区", "浦口区", "栖霞区", "雨花台区", "江宁区", "六合区", "溧水区", "高淳区"],
    "无锡市": ["锡山区", "惠山区", "滨湖区", "梁溪区", "新吴区", "江阴市", "宜兴市"],
    "徐州市": ["鼓楼区", "云龙区", "贾汪区", "泉山区", "铜山区", "丰县", "沛县", "睢宁县", "新沂市", "邳州市"],
    "常州市": ["天宁区", "钟楼区", "新北区", "武进区", "金坛区", "溧阳市"],
    "苏州市": ["姑苏区", "虎丘区", "吴中区", "相城区", "吴江区", "常熟市", "张家港市", "昆山市", "太仓市"],
    "宁波市": ["海曙区", "江北区", "镇海区", "北仑区", "鄞州区", "奉化区", "象山县", "宁海县", "余姚市", "慈溪市"],
    "温州市": ["鹿城区", "龙湾区", "瓯海区", "洞头区", "永嘉县", "平阳县", "苍南县", "文成县", "泰顺县", "瑞安市", "乐清市", "龙港市"],
    "合肥市": ["瑶海区", "庐阳区", "蜀山区", "包河区", "长丰县", "肥东县", "肥西县", "庐江县", "巢湖市"],
    "福州市": ["鼓楼区", "台江区", "仓山区", "马尾区", "晋安区", "长乐区", "闽侯县", "连江县", "罗源县", "闽清县", "永泰县", "平潭县", "福清市"],
    "厦门市": ["思明区", "海沧区", "湖里区", "集美区", "同安区", "翔安区"],
    "南昌市": ["东湖区", "西湖区", "青云谱区", "青山湖区", "新建区", "红谷滩区", "南昌县", "安义县", "进贤县"],
    "济南市": ["历下区", "市中区", "槐荫区", "天桥区", "历城区", "长清区", "章丘区", "济阳区", "莱芜区", "钢城区", "平阴县", "商河县"],
    "青岛市": ["市南区", "市北区", "黄岛区", "崂山区", "李沧区", "城阳区", "即墨区", "胶州市", "平度市", "莱西市"],
    "郑州市": ["中原区", "二七区", "管城回族区", "金水区", "上街区", "惠济区", "中牟县", "荥阳市", "新郑市", "新密市", "登封市"],
    "长沙市": ["芙蓉区", "天心区", "岳麓区", "开福区", "雨花区", "望城区", "长沙县", "浏阳市", "宁乡市"]
};*/

function updateCities() {
    var provinceSelect = document.getElementById('province');
    var citySelect = document.getElementById('city');
    var districtSelect = document.getElementById('district');
    
    var selectedProvince = provinceSelect.value;
    
    // 清空城市和区县选项
    citySelect.innerHTML = '<option value="">请选择城市</option>';
    districtSelect.innerHTML = '<option value="">请选择区县</option>';
    
    if (selectedProvince && addressData[selectedProvince]) {
        var cities = addressData[selectedProvince];
        for (var i = 0; i < cities.length; i++) {
            var option = document.createElement('option');
            option.value = cities[i];
            option.textContent = cities[i];
            citySelect.appendChild(option);
        }
    }
}

function updateDistricts() {
    var provinceSelect = document.getElementById('province');
    var citySelect = document.getElementById('city');
    var districtSelect = document.getElementById('district');
    
    var selectedCity = citySelect.value;
    var selectedProvince = provinceSelect.value;
    
    // 清空区县选项
    districtSelect.innerHTML = '<option value="">请选择区县</option>';
    
    // 对于直辖市，使用省份名称作为键
    var searchKey = selectedCity;
    if (!districtData[searchKey] && (selectedProvince === '北京市' || selectedProvince === '上海市' || selectedProvince === '天津市' || selectedProvince === '重庆市')) {
        searchKey = selectedProvince;
    }
    
    if (searchKey && districtData[searchKey]) {
        var districts = districtData[searchKey];
        for (var i = 0; i < districts.length; i++) {
            var option = document.createElement('option');
            option.value = districts[i];
            option.textContent = districts[i];
            districtSelect.appendChild(option);
        }
    }
}

// 显示地址表单弹窗
function showAddressForm() {
    document.getElementById('modalTitle').textContent = '新增收货地址';
    document.getElementById('addressForm').reset();
    document.getElementById('formAction').value = 'add';
    document.getElementById('formAddressId').value = '';
    
    // 复制当前选中的支付方式到地址表单
    var selectedPayment = document.querySelector('input[name="payment_method"]:checked');
    if (selectedPayment) {
        document.getElementById('addressFormPaymentMethod').value = selectedPayment.value;
    }
    
    document.getElementById('addressModal').classList.add('show');
    document.body.style.overflow = 'hidden'; // 防止背景滚动
}

// 关闭地址表单弹窗
function closeAddressForm() {
    document.getElementById('addressModal').classList.remove('show');
    document.body.style.overflow = 'auto'; // 恢复背景滚动
}

// 加载地址详情
function loadAddressDetails() {
    var selectElement = document.getElementById('selectedAddress');
    var selectedValue = selectElement.value;
    var selectedAddressDisplay = document.getElementById('selectedAddressDisplay');
    
    if (selectedValue === 'new') {
        // 显示地址弹窗
        showAddressForm();
        selectedAddressDisplay.style.display = 'none';
    } else if (selectedValue !== '') {
        // 如果选择了现有地址，显示地址信息
        selectedAddressDisplay.style.display = 'block';
    } else {
        // 没有选择任何地址
        selectedAddressDisplay.style.display = 'block';
    }
}

// 在页面加载时初始化
window.addEventListener('DOMContentLoaded', function() {
    // 检查URL参数中是否有预选地址
    const urlParams = new URLSearchParams(window.location.search);
    const preselectedAddress = urlParams.get('selected_address');
    
    if (preselectedAddress) {
        // 设置下拉框选中项
        const selectElement = document.getElementById('selectedAddress');
        if (selectElement) {
            selectElement.value = preselectedAddress;
        }
    }
    
    // 检查URL参数中是否有预选支付方式
    const preselectedPaymentMethod = urlParams.get('payment_method');
    if (preselectedPaymentMethod) {
        // 设置对应支付方式单选按钮为选中状态
        const paymentRadio = document.querySelector(`input[name="payment_method"][value="${preselectedPaymentMethod}"]`);
        if (paymentRadio) {
            paymentRadio.checked = true;
        }
    }
    
    // 初始化页面
    loadAddressDetails();
    
    // 绑定支付表单提交事件
    var paymentForm = document.getElementById('paymentForm');
    if(paymentForm) {
        paymentForm.addEventListener('submit', function(e) {
            var selectedPayment = document.querySelector('input[name="payment_method"]:checked');
            var selectedAddress = document.getElementById('selectedAddress').value;
            
            if (!selectedPayment) {
                e.preventDefault();
                alert('请选择支付方式');
                return false;
            }
            
            // 如果没有选择地址或选择了新地址，阻止提交
            if (!selectedAddress || selectedAddress === '') {
                e.preventDefault();
                alert('请选择收货地址');
                return false;
            }
            
            // 如果用户选择了添加新地址，阻止提交并打开地址弹窗
            if (selectedAddress === 'new') {
                e.preventDefault();
                showAddressForm();
                alert('请先添加收货地址');
                return false;
            }
            
            // 显示加载提示
            var submitBtn = this.querySelector('button[type="submit"]');
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 处理中...';
            submitBtn.disabled = true;
        });
    }
});

// ESC键关闭模态框
document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        var modal = document.getElementById('addressModal');
        if (modal && modal.classList.contains('show')) {
            modal.classList.remove('show');
            document.body.style.overflow = 'auto';
        }
    }
});

// 点击模态框外部关闭
window.onclick = function(event) {
    var modal = document.getElementById('addressModal');
    if (modal && event.target == modal) {
        modal.classList.remove('show');
        document.body.style.overflow = 'auto';
    }
}

// 地址表单提交处理
var addressForm = document.getElementById('addressForm');
if(addressForm) {
    addressForm.onsubmit = function(e) {
        // 同步最新的支付方式选择
        var selectedPayment = document.querySelector('input[name="payment_method"]:checked');
        if (selectedPayment) {
            document.getElementById('addressFormPaymentMethod').value = selectedPayment.value;
        }
        
        // 验证表单数据
        var consignee = document.getElementById('consignee').value;
        var phone = document.getElementById('phone').value;
        var province = document.getElementById('province').value;
        var city = document.getElementById('city').value;
        var district = document.getElementById('district').value;
        var address = document.getElementById('address').value;
        
        if(!consignee || !phone || !province || !city || !district || !address) {
            alert('请填写完整的收货信息');
            return false;
        }
        
        // 显示加载提示
        var submitBtn = this.querySelector('button[type="submit"]');
        submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 保存中...';
        submitBtn.disabled = true;
        
        // 提交表单
        return true;
    };
}

</script>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
