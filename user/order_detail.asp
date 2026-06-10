<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
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
        reviewSubmitMessage = "安全验证失败，请刷新页面后重试"
    Else
        ' 验证订单归属
        Dim checkOrderRs
        Set checkOrderRs = ExecuteQuery("SELECT Status FROM Orders WHERE OrderID = " & orderId & " AND UserID = " & userId)
        
        If checkOrderRs Is Nothing Or checkOrderRs.EOF Then
            reviewSubmitMessage = "订单不存在或无权限评价"
        Else
            Dim orderStatusForReview
            orderStatusForReview = checkOrderRs("Status")
            checkOrderRs.Close
            Set checkOrderRs = Nothing
            
            ' 检查订单状态是否允许评价（Paid, Delivered, Completed）
            If orderStatusForReview <> "Paid" And orderStatusForReview <> "Completed" Then
                reviewSubmitMessage = "当前订单状态不允许评价"
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
                    reviewSubmitMessage = "您已评价过此订单"
                Else
                    ' 获取表单数据
                    Dim rating, comment
                    rating = Request.Form("rating")
                    comment = Trim(Request.Form("comment"))
                    
                    ' 验证评分
                    If rating = "" Or Not IsNumeric(rating) Then
                        reviewSubmitMessage = "请选择评分星级"
                    ElseIf CInt(rating) < 1 Or CInt(rating) > 5 Then
                        reviewSubmitMessage = "评分必须在1-5星之间"
                    Else
                        ' 插入评价
                        Dim reviewSql
                        reviewSql = "INSERT INTO ProductReviews (OrderID, UserID, Rating, Comment, [Status], CreatedAt) VALUES (" & _
                                    orderId & ", " & userId & ", " & CLng(rating) & ", '" & SafeSQL(comment) & "', 'Pending', GETDATE())"
                        
                        If ExecuteNonQuery(reviewSql) Then
                            reviewSubmitSuccess = True
                            reviewSubmitMessage = "评价已提交，等待审核"
                        Else
                            reviewSubmitMessage = "评价提交失败，请稍后重试"
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
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/user/index.asp">个人中心</a>
        <span class="separator">/</span>
        <a href="/user/orders.asp">我的订单</a>
        <span class="separator">/</span>
        <span>订单详情</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!-- 侧边栏 -->
        <aside class="user-sidebar">
            <div class="user-profile">
                <h3><%= HTMLEncode(Session("Username")) %></h3>
                <p><%= HTMLEncode(Session("Email")) %></p>
            </div>
            
            <nav class="user-nav">
                <a href="/user/index.asp"><i class="fas fa-home"></i> 个人中心</a>
                <a href="/user/orders.asp"><i class="fas fa-list"></i> 我的订单</a>
                <a href="/user/settings.asp"><i class="fas fa-user-edit"></i> 账户设置</a>
                <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> 收货地址</a>
                <a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a>
                <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
            </nav>
        </aside>
        
        <!-- 主内容 -->
        <div class="user-main">
            <div class="user-card">
                <h2 class="card-title"><i class="fas fa-file-invoice"></i> 订单详情</h2>
                
                <div class="order-detail">
                    <!-- 订单基本信息 -->
                    <div class="order-basic-info">
                        <div class="info-row">
                            <span class="label">订单号:</span>
                            <span class="value"><%= orderNo %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">订单金额:</span>
                            <span class="value amount"><%= FormatMoney(totalAmount) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">支付方式:</span>
                            <span class="value">
                                <% Dim pmVal
                                If IsNumeric(paymentMethod) Then pmVal = CInt(paymentMethod) Else pmVal = -1
                                Select Case pmVal
                                    Case PAYMENT_METHOD_WECHAT
                                        Response.Write "微信支付"
                                    Case PAYMENT_METHOD_ALIPAY
                                        Response.Write "支付宝"
                                    Case PAYMENT_METHOD_PAYPAL
                                        Response.Write "PayPal"
                                    Case PAYMENT_METHOD_COD
                                        Response.Write "货到付款"
                                    Case Else
                                        Response.Write "未知"
                                End Select %>
                            </span>
                        </div>
                        <div class="info-row">
                            <span class="label">订单状态:</span>
                            <span class="value status-<%= orderStatus %>">
                                <% Select Case orderStatus
                                    Case "Pending"
                                        Response.Write "待支付"
                                    Case "Paid"
                                        Response.Write "已支付"
                                    Case "Failed"
                                        Response.Write "支付失败"
                                    Case "Refunded"
                                        Response.Write "已退款"
                                    Case Else
                                        Response.Write "未知状态"
                                End Select %>
                            </span>
                        </div>
                        <div class="info-row">
                            <span class="label">下单时间:</span>
                            <span class="value"><%= createdAt %></span>
                        </div>
                        <% If Not IsNull(updatedAt) And updatedAt <> "" Then %>
                        <div class="info-row">
                            <span class="label">支付时间:</span>
                            <span class="value"><%= updatedAt %></span>
                        </div>
                        <% End If %>
                        <% If Not IsNull(notes) And notes <> "" Then %>
                        <div class="info-row">
                            <span class="label">交易信息:</span>
                            <span class="value"><%= HTMLEncode(notes) %></span>
                        </div>
                        <% End If %>
                        <!-- 收货信息 -->
                        <!-- 使用订单表中保存的收货人信息，而不是从用户表获取实时信息 -->
                        <div class="info-row">
                            <span class="label">收货人:</span>
                            <span class="value"><%= HTMLEncode(shippingName) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">联系电话:</span>
                            <span class="value"><%= HTMLEncode(shippingPhone) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">收货地址:</span>
                            <span class="value"><%= HTMLEncode(shippingAddress) %></span>
                        </div>
                    </div>
                    
                    <!-- 订单商品列表 -->
                    <div class="order-products">
                        <h3>订单商品</h3>
                        <div class="products-list">
                            <!-- 
                            注意：这里需要从订单关联的购物车记录中获取商品信息
                            由于订单表中没有直接存储商品信息，这里仅作为示例
                            实际应用中需要在订单创建时同时保存商品快照信息
                            -->
                            <div class="order-item">
                                <div class="item-info">
                                    <p>订单创建时的商品信息已保存</p>
                                    <p>如有疑问请联系客服</p>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- 订单商品列表 -->
                    <div class="order-products">
                        <h3>订单商品</h3>
                        <div class="products-list">
                            <%
                            ' 解析订单Notes中的商品信息
                            Dim productInfo, productItems, item, i
                            productInfo = notes
                            
                            ' 如果Notes包含支付流水号信息，先提取商品详情部分
                            If InStr(productInfo, " | 支付流水号:") > 0 Then
                                ' 分离商品信息和支付信息
                                Dim notesParts
                                notesParts = Split(productInfo, " | 支付流水号:")
                                productInfo = Trim(notesParts(0))  ' 只取商品信息部分
                            End If
                            
                            If InStr(productInfo, "商品详情: ") > 0 Then
                                ' 提取商品详情部分
                                productInfo = Mid(productInfo, InStr(productInfo, "商品详情: ") + 5)
                                
                                ' 按分隔符拆分商品
                                productItems = Split(productInfo, "|")
                                
                                For i = 0 To UBound(productItems)
                                    item = Trim(productItems(i))
                                    If item <> "" Then
                            %>
                            <div class="order-item">
                                <div class="item-info">
                                    <%
                                    ' 解析并格式化商品信息
                                    Dim formattedItem, customInfoStart
                                    formattedItem = item
                                    
                                    ' 查找定制信息部分（在方括号中）
                                    customInfoStart = InStr(formattedItem, " [")
                                    If customInfoStart > 0 Then
                                        ' 分离基本商品信息和定制信息
                                        Dim baseInfo, customInfo
                                        baseInfo = Left(formattedItem, customInfoStart - 1)
                                        customInfo = Mid(formattedItem, customInfoStart + 2)  ' 跳过" ["
                                        customInfo = Left(customInfo, Len(customInfo) - 1)   ' 去掉最后的"]"
                                        
                                        ' 显示基本商品信息
                                        Response.Write "<div class=""base-info""><strong>" & HTMLEncode(baseInfo) & "</strong></div>"
                                        
                                        ' 显示定制信息
                                        If customInfo <> "" Then
                                            Response.Write "<div class=""custom-info"">定制信息: " & HTMLEncode(customInfo) & "</div>"
                                        End If
                                    Else
                                        ' 没有定制信息，直接显示
                                        Response.Write HTMLEncode(formattedItem)
                                    End If
                                    %>
                                </div>
                            </div>
                            <%
                                    End If
                                Next
                            Else
                                ' 如果没有解析到商品详情，显示原始Notes内容
                                If Len(productInfo) > 0 Then
                            %>
                            <div class="order-item">
                                <div class="item-info">
                                    <span class="product-name"><%= HTMLEncode(productInfo) %></span>
                                </div>
                            </div>
                            <%
                                End If
                            End If
                            %>
                        </div>
                    </div>
                    
                    <!-- 操作按钮 -->
                    <!-- 操作按钮 -->
                    <div class="order-actions">
                        <% If orderStatus = "Pending" Then %>
                        <a href="/checkout.asp?order_id=<%= orderId %>" class="btn btn-primary">立即支付</a>
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
                            <i class="fas fa-list-ul"></i> 查看所有产品成分
                        </a>
                        <%
                        End If
                        %>
                        
                        <a href="/user/orders.asp" class="btn btn-outline">返回订单列表</a>
                    </div>
                    
                    <!-- 评价区域 -->
                    <div id="review-section" class="review-section">
                        <h3><i class="fas fa-star"></i> 订单评价</h3>
                        
                        <% If reviewSubmitMessage <> "" Then %>
                        <div class="review-message <%= IIF(reviewSubmitSuccess, "success", "error") %>">
                            <%= reviewSubmitMessage %>
                        </div>
                        <% End If %>
                        
                        <% If hasReview Then %>
                        <!-- 已评价显示 -->
                        <div class="review-display">
                            <div class="review-rating">
                                <span class="rating-label">评分：</span>
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
                                <span class="rating-text"><%= reviewRating %> 星</span>
                            </div>
                            <% If reviewComment <> "" Then %>
                            <div class="review-comment">
                                <span class="comment-label">评价内容：</span>
                                <p class="comment-text"><%= HTMLEncode(reviewComment) %></p>
                            </div>
                            <% End If %>
                            <div class="review-meta">
                                <span class="review-status status-<%= reviewStatus %>">
                                    <% 
                                    Select Case reviewStatus
                                        Case "Pending": Response.Write "待审核"
                                        Case "Approved": Response.Write "已通过"
                                        Case "Rejected": Response.Write "未通过"
                                        Case Else: Response.Write reviewStatus
                                    End Select
                                    %>
                                </span>
                                <span class="review-date">提交时间：<%= reviewCreatedAt %></span>
                            </div>
                        </div>
                        <% ElseIf canReview Then %>
                        <!-- 评价表单 -->
                        <div class="review-form-container">
                            <p class="review-hint">请对本次购物体验进行评价：</p>
                            <form method="post" action="" class="review-form">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="action" value="submit_review">
                                
                                <div class="rating-input">
                                    <span class="rating-label">评分：</span>
                                    <div class="star-rating">
                                        <% For starIdx = 5 To 1 Step -1 %>
                                        <input type="radio" id="star<%= starIdx %>" name="rating" value="<%= starIdx %>" <%= IIF(starIdx = 5, "checked", "") %>>
                                        <label for="star<%= starIdx %>" title="<%= starIdx %> 星"><i class="fas fa-star"></i></label>
                                        <% Next %>
                                    </div>
                                </div>
                                
                                <div class="comment-input">
                                    <label for="comment">评价内容（选填）：</label>
                                    <textarea id="comment" name="comment" rows="4" maxlength="500" placeholder="请输入您的评价内容，最多500字..."></textarea>
                                    <span class="char-count">0/500</span>
                                </div>
                                
                                <div class="form-actions">
                                    <button type="submit" class="btn btn-primary">提交评价</button>
                                </div>
                            </form>
                        </div>
                        <% ElseIf orderStatus = "Pending" Then %>
                        <div class="review-notice">
                            <p><i class="fas fa-info-circle"></i> 订单支付完成后可进行评价</p>
                        </div>
                        <% Else %>
                        <div class="review-notice">
                            <p><i class="fas fa-info-circle"></i> 当前订单状态不支持评价</p>
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
