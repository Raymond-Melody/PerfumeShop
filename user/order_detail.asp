<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' ============================================
' 通用成分分割函数 - 支持所有类型的分隔符
' ============================================
Function SplitIngredients(rawStr)
    Dim result, arr, item, idx
    Set result = CreateObject("Scripting.Dictionary")
    If rawStr = "" Then
        Set SplitIngredients = result
        Exit Function
    End If
    rawStr = Replace(rawStr, "，", ",")
    rawStr = Replace(rawStr, vbCrLf, ",")
    rawStr = Replace(rawStr, vbLf, ",")
    rawStr = Replace(rawStr, vbCr, ",")
    rawStr = Replace(rawStr, Chr(160), ",")
    rawStr = Replace(rawStr, " ", ",")
    Do While InStr(rawStr, ",,") > 0
        rawStr = Replace(rawStr, ",,", ",")
    Loop
    arr = Split(rawStr, ",")
    For idx = 0 To UBound(arr)
        item = Trim(arr(idx))
        If item <> "" And Not result.Exists(item) Then
            result.Add item, True
        End If
    Next
    Set SplitIngredients = result
End Function
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/payment_config.asp"-->
<%
Call OpenConnection()

' 检查用户是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("URL"))
    Response.End
End If

Dim userId, orderId
userId = Session("UserID")
orderId = Request.QueryString("order_id")

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Redirect "/user/orders.asp"
    Response.End
End If

' ========== 评价提交处理 ==========
Dim reviewSubmitMessage, reviewSubmitSuccess
reviewSubmitMessage = ""
reviewSubmitSuccess = False

If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("action") = "submit_review" Then
    ' 验证CSRF Token
    If Not ValidateCSRFToken() Then
        reviewSubmitMessage = T("user_order_detail_review_csrf", Empty)
    Else
        ' 验证订单归属
        Dim checkOrderRs
        Set checkOrderRs = ExecuteQuery("SELECT Status FROM Orders WHERE OrderID = " & orderId & " AND UserID = " & userId)
        
        If checkOrderRs Is Nothing Or checkOrderRs.EOF Then
            reviewSubmitMessage = T("user_order_detail_review_no_order", Empty)
        Else
            Dim orderStatusForReview
            orderStatusForReview = checkOrderRs("Status")
            checkOrderRs.Close
            Set checkOrderRs = Nothing
            
            ' 检查订单状态是否允许评价（Paid, Delivered, Completed）
            If orderStatusForReview <> "Paid" And orderStatusForReview <> "Completed" Then
                reviewSubmitMessage = T("user_order_detail_cannot_review", Empty)
            Else
                ' 检查是否已评价
                Dim checkReviewRs
                Set checkReviewRs = ExecuteQuery("SELECT COUNT(*) as ReviewCount FROM ProductReviews WHERE OrderID = " & orderId & " AND UserID = " & userId)
                
                Dim hasReviewed
                hasReviewed = False
                If Not checkReviewRs Is Nothing And Not checkReviewRs.EOF Then
                    If CLng(checkReviewRs("ReviewCount")) > 0 Then
                        hasReviewed = True
                    End If
                    checkReviewRs.Close
                End If
                Set checkReviewRs = Nothing
                
                If hasReviewed Then
                    reviewSubmitMessage = T("user_order_detail_review_already", Empty)
                Else
                    ' 获取表单数据
                    Dim rating, comment
                    rating = Request.Form("rating")
                    comment = Trim(Request.Form("comment"))
                    
                    ' 验证评分
                    If rating = "" Or Not IsNumeric(rating) Then
                        reviewSubmitMessage = T("user_order_detail_review_no_rating", Empty)
                    ElseIf CInt(rating) < 1 Or CInt(rating) > 5 Then
                        reviewSubmitMessage = T("user_order_detail_review_invalid_rating", Empty)
                    Else
                        ' 插入评价
                        Dim reviewSql
                        reviewSql = "INSERT INTO ProductReviews (OrderID, UserID, Rating, Comment, [Status], CreatedAt) VALUES (" & _
                                    orderId & ", " & userId & ", " & CLng(rating) & ", '" & SafeSQL(comment) & "', 'Pending', GETDATE())"
                        
                        If ExecuteNonQuery(reviewSql) Then
                            reviewSubmitSuccess = True
                            reviewSubmitMessage = T("user_order_detail_review_submitted", Empty)
                        Else
                            reviewSubmitMessage = T("user_order_detail_review_failed", Empty)
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

' 获取订单信息
Dim rsOrder, orderInfo
Set rsOrder = ExecuteQuery("SELECT o.*, u.Username FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID WHERE o.OrderID = " & orderId & " AND o.UserID = " & userId)

If rsOrder Is Nothing Or rsOrder.EOF Then
    Response.Redirect "/user/orders.asp"
    Response.End
End If

orderInfo = rsOrder

' 复制订单信息到变量，因为RecordSet即将关闭
Dim orderNo, totalAmount, paymentMethod, orderStatus, createdAt, updatedAt, notes, shippingName, shippingPhone, shippingAddress
orderNo = orderInfo("OrderNo")
totalAmount = orderInfo("TotalAmount")
paymentMethod = orderInfo("PaymentMethod")
orderStatus = orderInfo("Status")
createdAt = orderInfo("CreatedAt")
updatedAt = orderInfo("UpdatedAt")
notes = orderInfo("Notes")
shippingName = orderInfo("ShippingName")
shippingPhone = orderInfo("ShippingPhone")
shippingAddress = orderInfo("ShippingAddress")

rsOrder.Close
Set rsOrder = Nothing

' 查询该订单的评价信息
Dim rsReview, hasReview, reviewRating, reviewComment, reviewStatus, reviewCreatedAt
hasReview = False
reviewRating = 0
reviewComment = ""
reviewStatus = ""
reviewCreatedAt = ""

Set rsReview = ExecuteQuery("SELECT Rating, Comment, [Status], CreatedAt FROM ProductReviews WHERE OrderID = " & orderId & " AND UserID = " & userId)

If Not rsReview Is Nothing Then
    If Not rsReview.EOF Then
        hasReview = True
        reviewRating = rsReview("Rating")
        reviewComment = rsReview("Comment")
        reviewStatus = rsReview("Status")
        reviewCreatedAt = rsReview("CreatedAt")
    End If
    rsReview.Close
End If
Set rsReview = Nothing

' 判断是否可以评价（订单状态为Paid或Completed且未评价）
Dim canReview
canReview = False
If Not hasReview And (orderStatus = "Paid" Or orderStatus = "Completed") Then
    canReview = True
End If
%>
<!--#include file="../includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <a href="/user/index.asp"><% If FEATURE_I18N Then %><%= T("user_nav_center", Empty) %><% Else %>个人中心<% End If %></a>
        <span class="separator">/</span>
        <a href="/user/orders.asp"><% If FEATURE_I18N Then %><%= T("user_orders_title", Empty) %><% Else %>我的订单<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then %><%= T("user_order_detail_title", Empty) %><% Else %>订单详情<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!--#include file="nav.asp"-->
        
        <!-- 主内容 -->
        <div class="user-main">
            <div class="user-card">
                <h2 class="card-title"><i class="fas fa-file-invoice"></i> <% If FEATURE_I18N Then %><%= T("user_order_detail_title", Empty) %><% Else %>订单详情<% End If %></h2>
                
                <div class="order-detail">
                    <!-- 订单基本信息 -->
                    <div class="order-basic-info">
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_order_no", Empty) %><% Else %>订单号<% End If %>:</span>
                            <span class="value"><%= orderNo %></span>
                        </div>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_amount", Empty) %><% Else %>订单金额<% End If %>:</span>
                            <span class="value amount"><%= FormatMoney(totalAmount) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_payment", Empty) %><% Else %>支付方式<% End If %>:</span>
                            <span class="value">
                                <% Dim pmVal
                                If IsNumeric(paymentMethod) Then pmVal = CInt(paymentMethod) Else pmVal = -1
                                Select Case pmVal
                                    Case PAYMENT_METHOD_WECHAT
                                        Response.Write T("user_orders_pm_wechat", Empty)
                                    Case PAYMENT_METHOD_ALIPAY
                                        Response.Write T("user_orders_pm_alipay", Empty)
                                    Case PAYMENT_METHOD_PAYPAL
                                        Response.Write T("user_orders_pm_paypal", Empty)
                                    Case PAYMENT_METHOD_COD
                                        Response.Write T("user_orders_pm_cod", Empty)
                                    Case Else
                                        Response.Write T("user_orders_pm_unknown", Empty)
                                End Select %>
                            </span>
                        </div>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_status", Empty) %><% Else %>订单状态<% End If %>:</span>
                            <span class="value status-<%= orderStatus %>">
                                <% Select Case orderStatus
                                    Case "Pending"
                                        Response.Write T("user_orders_status_pending", Empty)
                                    Case "Paid"
                                        Response.Write T("user_orders_status_paid", Empty)
                                    Case "Failed"
                                        Response.Write T("user_orders_status_failed", Empty)
                                    Case "Refunded"
                                        Response.Write T("user_orders_status_refunded", Empty)
                                    Case Else
                                        Response.Write T("user_orders_status_unknown", Empty)
                                End Select %>
                            </span>
                        </div>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_order_time", Empty) %><% Else %>下单时间<% End If %>:</span>
                            <span class="value"><%= createdAt %></span>
                        </div>
                        <% If Not IsNull(updatedAt) And updatedAt <> "" Then %>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_pay_time", Empty) %><% Else %>支付时间<% End If %>:</span>
                            <span class="value"><%= updatedAt %></span>
                        </div>
                        <% End If %>
                        <% If Not IsNull(notes) And notes <> "" Then %>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_transaction", Empty) %><% Else %>交易信息<% End If %>:</span>
                            <span class="value"><%= HTMLEncode(notes) %></span>
                        </div>
                        <% End If %>
                        <!-- 收货信息 -->
                        <!-- 使用订单表中保存的收货人信息，而不是从用户表获取实时信息 -->
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_consignee", Empty) %><% Else %>收货人<% End If %>:</span>
                            <span class="value"><%= HTMLEncode(shippingName) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_phone", Empty) %><% Else %>联系电话<% End If %>:</span>
                            <span class="value"><%= HTMLEncode(shippingPhone) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label"><% If FEATURE_I18N Then %><%= T("user_order_detail_address", Empty) %><% Else %>收货地址<% End If %>:</span>
                            <span class="value"><%= HTMLEncode(shippingAddress) %></span>
                        </div>
                    </div>
                    
                    <!-- 订单商品列表 -->
                    <div class="order-products">
                        <h3><% If FEATURE_I18N Then %><%= T("user_order_detail_products", Empty) %><% Else %>订单商品<% End If %></h3>
                        <div class="products-list">
                            <%
                            ' 从OrderDetails和OrderDetailNoteSelections查询订单商品详情
                            Dim rsOrderDetails, odDetailId, odProductId, odProductName, odQuantity, odUnitPrice, odSubtotal
                            Dim odVolumeName, odVolumeML, odBottleName, odCustomLabel
                            
                            Set rsOrderDetails = ExecuteQuery("SELECT od.*, p.ProductType FROM OrderDetails od LEFT JOIN Products p ON od.ProductID = p.ProductID WHERE od.OrderID = " & orderId)
                            
                            If Not rsOrderDetails Is Nothing Then
                                Do While Not rsOrderDetails.EOF
                                    odDetailId = rsOrderDetails("DetailID")
                                    odProductId = rsOrderDetails("ProductID")
                                    odProductName = rsOrderDetails("ProductName") & ""
                                    odQuantity = rsOrderDetails("Quantity")
                                    odUnitPrice = rsOrderDetails("UnitPrice")
                                    odSubtotal = rsOrderDetails("Subtotal")
                                    odVolumeName = rsOrderDetails("VolumeName") & ""
                                    odVolumeML = rsOrderDetails("VolumeML") & ""
                                    odBottleName = rsOrderDetails("BottleName") & ""
                                    odCustomLabel = rsOrderDetails("CustomLabel") & ""
                                    Dim odProductType, odProductTypeLC
                                    odProductType = rsOrderDetails("ProductType") & ""
                                    odProductTypeLC = LCase(odProductType)
                            %>
                            <div class="order-item">
                                <div class="item-info">
                                    <div class="base-info"><strong><%= HTMLEncode(odProductName) %></strong> ×<%= odQuantity %> | <% If FEATURE_I18N Then %><%= T("user_order_detail_unit_price", Empty) %><% Else %>单价<% End If %>: <%= FormatMoney(odUnitPrice) %> | <% If FEATURE_I18N Then %><%= T("subtotal", Empty) %><% Else %>小计<% End If %>: <%= FormatMoney(odSubtotal) %></div>
                                    <%
                                    ' KOL推荐产品与品牌定香产品不显示香调配比信息
                                    Dim rsOrderNotes, odTopList, odMidList, odBaseList, odNoteType, odNoteName, odPercent
                                    odTopList = "": odMidList = "": odBaseList = ""
                                    If odProductTypeLC = "custom" Then
                                        Set rsOrderNotes = ExecuteQuery("SELECT s.*, n.NoteName FROM OrderDetailNoteSelections s LEFT JOIN FragranceNotes n ON s.NoteID = n.NoteID WHERE s.DetailID = " & odDetailId & " ORDER BY s.NoteType")
                                        If Not rsOrderNotes Is Nothing Then
                                            Do While Not rsOrderNotes.EOF
                                                odNoteType = Trim(rsOrderNotes("NoteType") & "")
                                                odNoteName = HTMLEncode(rsOrderNotes("NoteName") & "")
                                                odPercent = rsOrderNotes("Percentage")
                                                If odNoteType = "前调" Then
                                                    If odTopList <> "" Then odTopList = odTopList & ", "
                                                    odTopList = odTopList & odNoteName & " (" & odPercent & "%)"
                                                ElseIf odNoteType = "中调" Then
                                                    If odMidList <> "" Then odMidList = odMidList & ", "
                                                    odMidList = odMidList & odNoteName & " (" & odPercent & "%)"
                                                ElseIf odNoteType = "后调" Then
                                                    If odBaseList <> "" Then odBaseList = odBaseList & ", "
                                                    odBaseList = odBaseList & odNoteName & " (" & odPercent & "%)"
                                                End If
                                                rsOrderNotes.MoveNext
                                            Loop
                                            rsOrderNotes.Close
                                            Set rsOrderNotes = Nothing
                                        End If
                                    End If
                                    
                                    If odProductTypeLC = "custom" And odTopList <> "" Then
                                        Response.Write "<div class='custom-info'>" & T("user_order_detail_note_top", Empty) & ": " & odTopList & "</div>"
                                    End If
                                    If odProductTypeLC = "custom" And odMidList <> "" Then
                                        Response.Write "<div class='custom-info'>" & T("user_order_detail_note_mid", Empty) & ": " & odMidList & "</div>"
                                    End If
                                    If odProductTypeLC = "custom" And odBaseList <> "" Then
                                        Response.Write "<div class='custom-info'>" & T("user_order_detail_note_base", Empty) & ": " & odBaseList & "</div>"
                                    End If
                                    
                                    If odVolumeName <> "" Then
                                        Response.Write "<div class='custom-info'>" & T("user_order_detail_volume", Empty) & ": " & odVolumeML & "ml (" & HTMLEncode(odVolumeName) & ")</div>"
                                    End If
                                    If odBottleName <> "" Then
                                        Response.Write "<div class='custom-info'>" & T("user_order_detail_bottle", Empty) & ": " & HTMLEncode(odBottleName) & "</div>"
                                    End If
                                    If odCustomLabel <> "" Then
                                        Response.Write "<div class='custom-info'>" & T("user_order_detail_engraving", Empty) & ": " & HTMLEncode(odCustomLabel) & "</div>"
                                    End If
                                    
                                    ' 显示成分/过敏原信息（仅对定制和KOL产品可见，品牌定香产品随包装附成分说明书）
                                    If odProductTypeLC = "custom" Or odProductTypeLC = "kol" Then
                                    ' 使用Dictionary去重 + SplitIngredients分割复合成分
                                    Dim rsOrderIngr, odUniqueIngr, odRawIngr, odSplitResult, odSplitKey
                                    Set odUniqueIngr = CreateObject("Scripting.Dictionary")
                                    Set rsOrderIngr = ExecuteQuery("SELECT IngredientName FROM OrderIngredients WHERE DetailID = " & odDetailId & " ORDER BY IngredientID")
                                    If Not rsOrderIngr Is Nothing Then
                                        Do While Not rsOrderIngr.EOF
                                            odRawIngr = rsOrderIngr("IngredientName") & ""
                                            Set odSplitResult = SplitIngredients(odRawIngr)
                                            For Each odSplitKey In odSplitResult.Keys
                                                If Not odUniqueIngr.Exists(odSplitKey) Then
                                                    odUniqueIngr.Add odSplitKey, True
                                                End If
                                            Next
                                            Set odSplitResult = Nothing
                                            rsOrderIngr.MoveNext
                                        Loop
                                        rsOrderIngr.Close
                                        Set rsOrderIngr = Nothing
                                    End If
                                    
                                    ' 排序后显示
                                    If odUniqueIngr.Count > 0 Then
                                        Dim odIngrArr(), odI, odJ, odTempKey, odIngrList
                                        ReDim odIngrArr(odUniqueIngr.Count - 1)
                                        odI = 0
                                        For Each odSplitKey In odUniqueIngr.Keys
                                            odIngrArr(odI) = odSplitKey
                                            odI = odI + 1
                                        Next
                                        For odI = 0 To UBound(odIngrArr) - 1
                                            For odJ = odI + 1 To UBound(odIngrArr)
                                                If odIngrArr(odI) > odIngrArr(odJ) Then
                                                    odTempKey = odIngrArr(odI)
                                                    odIngrArr(odI) = odIngrArr(odJ)
                                                    odIngrArr(odJ) = odTempKey
                                                End If
                                            Next
                                        Next
                                        odIngrList = ""
                                        For odI = 0 To UBound(odIngrArr)
                                            If odIngrList <> "" Then odIngrList = odIngrList & ", "
                                            odIngrList = odIngrList & HTMLEncode(odIngrArr(odI))
                                        Next
                                        Response.Write "<div class='custom-info' style='color:#888;font-size:12px;margin-top:4px;'><i class='fas fa-flask'></i> " & T("user_order_detail_ingredients_label", Empty) & ": " & odIngrList & "</div>"
                                    End If
                                    End If
                                    Set odUniqueIngr = Nothing
                                    %>
                                </div>
                            </div>
                            <%
                                    rsOrderDetails.MoveNext
                                Loop
                                rsOrderDetails.Close
                                Set rsOrderDetails = Nothing
                            Else
                            %>
                            <div class="order-item">
                                <div class="item-info">
                                    <p><% If FEATURE_I18N Then %><%= T("user_order_detail_no_products", Empty) %><% Else %>暂无订单商品详情<% End If %></p>
                                </div>
                            </div>
                            <% End If %>
                        </div>
                    </div>
                    
                    <!-- 操作按钮 -->
                    <!-- 操作按钮 -->
                    <div class="order-actions">
                        <% If orderStatus = "Pending" Then %>
                        <a href="/checkout.asp?order_id=<%= orderId %>" class="btn btn-primary"><% If FEATURE_I18N Then %><%= T("user_order_detail_pay_now", Empty) %><% Else %>立即支付<% End If %></a>
                        <% End If %>
                        
                        <%
                        ' 检查订单中是否有需要显示成分的产品（定制香水或 KOL 推荐）
                        Dim rsCheckProducts, hasCustomOrKOL
                        hasCustomOrKOL = False
                        Set rsCheckProducts = ExecuteQuery("SELECT COUNT(*) as cnt FROM OrderDetails od LEFT JOIN Products p ON od.ProductID = p.ProductID WHERE od.OrderID = " & orderId & " AND (p.ProductType = 'Custom' OR p.ProductType = 'KOL')")
                        If Not rsCheckProducts Is Nothing Then
                            If Not rsCheckProducts.EOF Then
                                If rsCheckProducts("cnt") > 0 Then
                                    hasCustomOrKOL = True
                                End If
                            End If
                            rsCheckProducts.Close
                            Set rsCheckProducts = Nothing
                        End If
                        
                        ' 只有当订单包含定制香水或 KOL 推荐产品时才显示查看成分按钮
                        If hasCustomOrKOL Then
                        %>
                        <a href="/order_ingredients.asp?order_id=<%= orderId %>" class="btn btn-success" target="_blank">
                            <i class="fas fa-list-ul"></i> <% If FEATURE_I18N Then %><%= T("user_order_detail_ingredients", Empty) %><% Else %>查看所有产品成分<% End If %>
                        </a>
                        <%
                        End If
                        %>
                        
                        <a href="/user/orders.asp" class="btn btn-outline"><% If FEATURE_I18N Then %><%= T("user_order_detail_back_list", Empty) %><% Else %>返回订单列表<% End If %></a>
                    </div>
                    
                    <!-- 评价区域 -->
                    <div id="review-section" class="review-section">
                        <h3><i class="fas fa-star"></i> <% If FEATURE_I18N Then %><%= T("user_order_detail_review_title", Empty) %><% Else %>订单评价<% End If %></h3>
                        
                        <% If reviewSubmitMessage <> "" Then %>
                        <div class="review-message <%= IIF(reviewSubmitSuccess, "success", "error") %>">
                            <%= reviewSubmitMessage %>
                        </div>
                        <% End If %>
                        
                        <% If hasReview Then %>
                        <!-- 已评价显示 -->
                        <div class="review-display">
                            <div class="review-rating">
                                <span class="rating-label"><% If FEATURE_I18N Then %><%= T("user_order_detail_rating_label", Empty) %><% Else %>评分：<% End If %></span>
                                <span class="rating-stars">
                                    <% 
                                    Dim starIdx
                                    For starIdx = 1 To 5
                                        If starIdx <= reviewRating Then
                                            Response.Write "<i class='fas fa-star star-filled'></i>"
                                        Else
                                            Response.Write "<i class='far fa-star'></i>"
                                        End If
                                    Next
                                    %>
                                </span>
                                <span class="rating-text"><%= reviewRating %> <% If FEATURE_I18N Then %><%= T("user_order_detail_star_hint", Empty) %><% Else %>星<% End If %></span>
                            </div>
                            <% If reviewComment <> "" Then %>
                            <div class="review-comment">
                                <span class="comment-label"><% If FEATURE_I18N Then %><%= T("user_order_detail_comment_label", Empty) %><% Else %>评价内容：<% End If %></span>
                                <p class="comment-text"><%= HTMLEncode(reviewComment) %></p>
                            </div>
                            <% End If %>
                            <div class="review-meta">
                                <span class="review-status status-<%= reviewStatus %>">
                                    <% 
                                    Select Case reviewStatus
                                        Case "Pending": Response.Write T("user_order_detail_review_pending", Empty)
                                        Case "Approved": Response.Write T("user_order_detail_review_approved", Empty)
                                        Case "Rejected": Response.Write T("user_order_detail_review_rejected", Empty)
                                        Case Else: Response.Write reviewStatus
                                    End Select
                                    %>
                                </span>
                                <span class="review-date"><% If FEATURE_I18N Then %><%= T("user_order_detail_submitted_at", Empty) %><% Else %>提交时间：<% End If %><%= reviewCreatedAt %></span>
                            </div>
                        </div>
                        <% ElseIf canReview Then %>
                        <!-- 评价表单 -->
                        <div class="review-form-container">
                            <p class="review-hint"><% If FEATURE_I18N Then %><%= T("user_order_detail_review_hint", Empty) %><% Else %>请对本次购物体验进行评价：<% End If %></p>
                            <form method="post" action="" class="review-form">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="action" value="submit_review">
                                
                                <div class="rating-input">
                                    <span class="rating-label"><% If FEATURE_I18N Then %><%= T("user_order_detail_rating_label", Empty) %><% Else %>评分：<% End If %></span>
                                    <div class="star-rating">
                                        <% For starIdx = 5 To 1 Step -1 %>
                                        <input type="radio" id="star<%= starIdx %>" name="rating" value="<%= starIdx %>" <%= IIF(starIdx = 5, "checked", "") %>>
                                        <label for="star<%= starIdx %>" title="<%= starIdx %> <% If FEATURE_I18N Then %><%= T("user_order_detail_star_hint", Empty) %><% Else %>星<% End If %>"><i class="fas fa-star"></i></label>
                                        <% Next %>
                                    </div>
                                </div>
                                
                                <div class="comment-input">
                                    <label for="comment"><% If FEATURE_I18N Then %><%= T("user_order_detail_comment_label", Empty) %><% Else %>评价内容（选填）：<% End If %></label>
                                    <textarea id="comment" name="comment" rows="4" maxlength="500" placeholder="<% If FEATURE_I18N Then %><%= T("user_order_detail_comment_placeholder", Empty) %><% Else %>请输入您的评价内容，最多500字...<% End If %>"></textarea>
                                    <span class="char-count">0/500</span>
                                </div>
                                
                                <div class="form-actions">
                                    <button type="submit" class="btn btn-primary"><% If FEATURE_I18N Then %><%= T("user_order_detail_submit_review", Empty) %><% Else %>提交评价<% End If %></button>
                                </div>
                            </form>
                        </div>
                        <% ElseIf orderStatus = "Pending" Then %>
                        <div class="review-notice">
                            <p><i class="fas fa-info-circle"></i> <% If FEATURE_I18N Then %><%= T("user_order_detail_wait_payment", Empty) %><% Else %>订单支付完成后可进行评价<% End If %></p>
                        </div>
                        <% Else %>
                        <div class="review-notice">
                            <p><i class="fas fa-info-circle"></i> <% If FEATURE_I18N Then %><%= T("user_order_detail_cannot_review", Empty) %><% Else %>当前订单状态不支持评价<% End If %></p>
                        </div>
                        <% End If %>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<style>
/* 评价区域样式 */
.review-section {
    margin-top: 30px;
    padding: 25px;
    background-color: #f8f9fa;
    border-radius: 8px;
    border: 1px solid #e9ecef;
}

.review-section h3 {
    margin-bottom: 20px;
    color: #333;
    font-size: 18px;
    border-bottom: 2px solid #dee2e6;
    padding-bottom: 10px;
}

.review-message {
    padding: 12px 15px;
    border-radius: 5px;
    margin-bottom: 20px;
}

.review-message.success {
    background-color: #d4edda;
    color: #155724;
    border: 1px solid #c3e6cb;
}

.review-message.error {
    background-color: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
}

/* 已评价显示 */
.review-display {
    background-color: #fff;
    padding: 20px;
    border-radius: 6px;
    border: 1px solid #dee2e6;
}

.review-rating {
    margin-bottom: 15px;
}

.rating-label {
    font-weight: bold;
    color: #555;
}

.rating-stars {
    margin: 0 10px;
}

.rating-stars .star-filled {
    color: #ffc107;
}

.rating-stars .fa-star {
    color: #ddd;
    font-size: 18px;
}

.rating-stars .star-filled {
    color: #ffc107;
}

.rating-text {
    color: #666;
    font-size: 14px;
}

.review-comment {
    margin-bottom: 15px;
}

.comment-label {
    font-weight: bold;
    color: #555;
    display: block;
    margin-bottom: 8px;
}

.comment-text {
    color: #333;
    line-height: 1.6;
    background-color: #f8f9fa;
    padding: 12px;
    border-radius: 4px;
    margin: 0;
}

.review-meta {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding-top: 15px;
    border-top: 1px solid #eee;
}

.review-status {
    padding: 4px 12px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: bold;
}

.review-status.status-Pending {
    background-color: #fff3cd;
    color: #856404;
}

.review-status.status-Approved {
    background-color: #d4edda;
    color: #155724;
}

.review-status.status-Rejected {
    background-color: #f8d7da;
    color: #721c24;
}

.review-date {
    color: #888;
    font-size: 13px;
}

/* 评价表单 */
.review-form-container {
    background-color: #fff;
    padding: 20px;
    border-radius: 6px;
    border: 1px solid #dee2e6;
}

.review-hint {
    color: #666;
    margin-bottom: 15px;
}

.rating-input {
    margin-bottom: 20px;
}

/* 星级评分交互 */
.star-rating {
    display: inline-flex;
    flex-direction: row-reverse;
    gap: 5px;
}

.star-rating input {
    display: none;
}

.star-rating label {
    cursor: pointer;
    font-size: 28px;
    color: #ddd;
    transition: color 0.2s;
}

.star-rating label:hover,
.star-rating label:hover ~ label,
.star-rating input:checked ~ label {
    color: #ffc107;
}

.star-rating label:hover i,
.star-rating label:hover ~ label i,
.star-rating input:checked ~ label i {
    color: #ffc107;
}

.comment-input {
    margin-bottom: 20px;
}

.comment-input label {
    display: block;
    margin-bottom: 8px;
    font-weight: bold;
    color: #555;
}

.comment-input textarea {
    width: 100%;
    padding: 12px;
    border: 1px solid #ddd;
    border-radius: 4px;
    resize: vertical;
    font-family: inherit;
    font-size: 14px;
}

.comment-input textarea:focus {
    outline: none;
    border-color: #80bdff;
    box-shadow: 0 0 0 3px rgba(0,123,255,.1);
}

.char-count {
    display: block;
    text-align: right;
    color: #888;
    font-size: 12px;
    margin-top: 5px;
}

.form-actions {
    text-align: center;
}

/* 评价提示 */
.review-notice {
    background-color: #e9ecef;
    padding: 20px;
    border-radius: 6px;
    text-align: center;
    color: #666;
}

.review-notice i {
    margin-right: 8px;
    color: #17a2b8;
}
</style>

<script>
// 评价表单字符计数
document.addEventListener('DOMContentLoaded', function() {
    var commentTextarea = document.getElementById('comment');
    var charCount = document.querySelector('.char-count');
    
    if (commentTextarea && charCount) {
        commentTextarea.addEventListener('input', function() {
            var currentLength = this.value.length;
            charCount.textContent = currentLength + '/500';
            
            if (currentLength >= 450) {
                charCount.style.color = '#dc3545';
            } else {
                charCount.style.color = '#888';
            }
        });
    }
});
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>
