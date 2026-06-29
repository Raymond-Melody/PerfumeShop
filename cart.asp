<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 自定义IIf函数
Function IIf(condition, trueVal, falseVal)
    If condition Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/dal_cart.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<%
Call OpenConnection()

' V14: 会员登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If

' 获取SessionID用于匿名购物车
Dim sessionId, userId
sessionId = Session.SessionID
userId = Session("UserID")
%>
<!--#include file="includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then %><%= T("cart_breadcrumb", Empty) %><% Else %>购物车<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="cart-page">
        <h1 class="page-title"><i class="fas fa-shopping-cart"></i> <% If FEATURE_I18N Then %><%= T("cart_title", Empty) %><% Else %>我的购物车<% End If %></h1>
        
        <%
        Dim rsCart, cartCount, cartTotal, totalEngravingFee
        Dim whereClause, subtotal, rsNoteSel, topList, midList, baseList, currentNoteType, currentNoteName, currentPercent
        Dim itemEngravingPrice
        cartTotal = 0
        cartCount = 0
        totalEngravingFee = 0
        
        ' V17: 使用参数化DAL查询
        If userId <> "" Then
            Set rsCart = DAL_Cart_GetByUser(userId)
        Else
            Set rsCart = DAL_Cart_GetBySession(sessionId)
        End If
        
        Response.Write "<!-- DEBUG: SQL executed successfully -->" & vbCrLf
        
        Response.Write "<!-- DEBUG: whereClause=" & whereClause & ", userId=" & userId & ", sessionId=" & sessionId & " -->" & vbCrLf
        
        Dim hasItems
        hasItems = False
        If Not rsCart Is Nothing Then
            Response.Write "<!-- DEBUG: rsCart is not Nothing -->" & vbCrLf
            If Not rsCart.EOF Then
                Response.Write "<!-- DEBUG: rsCart.EOF = False, has records -->" & vbCrLf
                hasItems = True
            Else
                Response.Write "<!-- DEBUG: rsCart.EOF = True, no records -->" & vbCrLf
            End If
        Else
            Response.Write "<!-- DEBUG: rsCart is Nothing -->" & vbCrLf
        End If
        
        If hasItems Then
        %>
        <div class="cart-content">
            <div class="cart-items">
                <div class="cart-header-row">
                    <div class="col-checkbox">
                        <input type="checkbox" id="selectAll" checked onchange="toggleAllSelection()"> <% If FEATURE_I18N Then %><%= T("cart_select_all", Empty) %><% Else %>全选<% End If %>
                    </div>
                    <span class="col-product"><% If FEATURE_I18N Then %><%= T("cart_col_product", Empty) %><% Else %>商品信息<% End If %></span>
                    <span class="col-price"><% If FEATURE_I18N Then %><%= T("cart_col_price", Empty) %><% Else %>单价<% End If %></span>
                    <span class="col-quantity"><% If FEATURE_I18N Then %><%= T("cart_col_quantity", Empty) %><% Else %>数量<% End If %></span>
                    <span class="col-subtotal"><% If FEATURE_I18N Then %><%= T("cart_col_subtotal", Empty) %><% Else %>小计<% End If %></span>
                    <span class="col-action"><% If FEATURE_I18N Then %><%= T("cart_col_action", Empty) %><% Else %>操作<% End If %></span>
                </div>
                
                <%
                Do While Not rsCart.EOF
                    cartCount = cartCount + 1
                    subtotal = CDbl(rsCart("UnitPrice")) * CDbl(rsCart("Quantity"))
                    cartTotal = cartTotal + subtotal
                    
                    ' 计算刻字费用
                    itemEngravingPrice = 0
                    On Error Resume Next
                    itemEngravingPrice = CDbl(rsCart("EngravingPrice"))
                    If Err.Number <> 0 Then itemEngravingPrice = 0
                    On Error GoTo 0
                    If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" And itemEngravingPrice > 0 Then
                        totalEngravingFee = totalEngravingFee + (itemEngravingPrice * rsCart("Quantity"))
                    End If
                %>
                <%
                ' 计算当前商品的刻字费用（用于前端JS动态计算）
                Dim itemEngFee
                itemEngFee = 0
                If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" And itemEngravingPrice > 0 Then
                    itemEngFee = itemEngravingPrice * CDbl(rsCart("Quantity"))
                End If
                %>
                <div class="cart-item" data-id="<%= rsCart("CartID") %>" data-price="<%= rsCart("UnitPrice") %>" data-quantity="<%= rsCart("Quantity") %>" data-subtotal="<%= subtotal %>" data-engraving-fee="<%= itemEngFee %>">
                    <div class="col-checkbox">
                        <input type="checkbox" class="item-checkbox" checked onchange="updateSelectedCount()">
                    </div>
                    <div class="col-product">
                        <div class="product-image">
                            <img src="<%= rsCart("ImageURL") %>" alt="<%= HTMLEncode(rsCart("ProductName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                        </div>
                        <div class="product-details">
                            <h3><%= HTMLEncode(rsCart("ProductName")) %></h3>
                            <div class="customize-info">
                                <%
                                ' 获取详细香调配比
                                topList = "": midList = "": baseList = ""
                                Dim cartProductType, cartProductTypeLC
                                cartProductType = rsCart("ProductType") & ""
                                cartProductTypeLC = LCase(cartProductType)
                                ' KOL推荐产品与品牌定香产品不显示香调配比信息
                                If cartProductTypeLC = "custom" Then
                                    ' V17: 使用参数化DAL查询
                                    Set rsNoteSel = DAL_Cart_GetNoteSelections(rsCart("CartID"))
                                    If Not rsNoteSel Is Nothing Then
                                        Do While Not rsNoteSel.EOF
                                            currentNoteType = Trim(rsNoteSel("NoteType") & "")
                                            currentNoteName = HTMLEncode(rsNoteSel("NoteName") & "")
                                            currentPercent = rsNoteSel("Percentage")
                                            
                                            If currentNoteType = "前调" Then
                                                If topList <> "" Then topList = topList & ", "
                                                topList = topList & currentNoteName & " (" & currentPercent & "%)"
                                            ElseIf currentNoteType = "中调" Then
                                                If midList <> "" Then midList = midList & ", "
                                                midList = midList & currentNoteName & " (" & currentPercent & "%)"
                                            ElseIf currentNoteType = "后调" Then
                                                If baseList <> "" Then baseList = baseList & ", "
                                                baseList = baseList & currentNoteName & " (" & currentPercent & "%)"
                                            End If
                                            rsNoteSel.MoveNext
                                        Loop
                                        rsNoteSel.Close
                                        Set rsNoteSel = Nothing
                                    End If
                                End If
                                %>
                                <% If cartProductTypeLC = "custom" And topList <> "" Then %>
                                <span><i class="fas fa-wind"></i> <% If FEATURE_I18N Then %><%= T("cart_note_top", Empty) %><% Else %>前调<% End If %>: <%= topList %></span>
                                <% End If %>
                                <% If cartProductTypeLC = "custom" And midList <> "" Then %>
                                <span><i class="fas fa-heart"></i> <% If FEATURE_I18N Then %><%= T("cart_note_mid", Empty) %><% Else %>中调<% End If %>: <%= midList %></span>
                                <% End If %>
                                <% If cartProductTypeLC = "custom" And baseList <> "" Then %>
                                <span><i class="fas fa-moon"></i> <% If FEATURE_I18N Then %><%= T("cart_note_base", Empty) %><% Else %>后调<% End If %>: <%= baseList %></span>
                                <% End If %>
                                <% If Not IsNull(rsCart("VolumeName")) Then %>
                                <span><i class="fas fa-tint"></i> <% If FEATURE_I18N Then %><%= T("cart_volume", Empty) %><% Else %>容量<% End If %>: <%= rsCart("VolumeML") %>ml (<%= HTMLEncode(rsCart("VolumeName")) %>)</span>
                                <% End If %>
                                <% If Not IsNull(rsCart("BottleName")) Then %>
                                <span><i class="fas fa-wine-bottle"></i> <% If FEATURE_I18N Then %><%= T("cart_bottle", Empty) %><% Else %>瓶身<% End If %>: <%= HTMLEncode(rsCart("BottleName")) %></span>
                                <% End If %>
                                <% If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" Then %>
                                <span><i class="fas fa-pen-fancy"></i> <% If FEATURE_I18N Then %><%= T("cart_label", Empty) %><% Else %>刻字<% End If %>: <%= HTMLEncode(rsCart("CustomLabel")) %></span>
                                <% End If %>
                                <% 
                                ' 显示刻字费用（如果有刻字且产品有刻字费用）
                                ' 注意：itemEngravingPrice 已在第43行声明，此处仅赋值
                                itemEngravingPrice = 0
                                On Error Resume Next
                                itemEngravingPrice = CDbl(rsCart("EngravingPrice"))
                                If Err.Number <> 0 Then itemEngravingPrice = 0
                                On Error GoTo 0
                                If Not IsNull(rsCart("CustomLabel")) And rsCart("CustomLabel") <> "" And itemEngravingPrice > 0 Then 
                                %>
                                <span style="color:#e91e63;"><i class="fas fa-tag"></i> <% If FEATURE_I18N Then %><%= T("cart_label_fee", Empty) %><% Else %>刻字费用<% End If %>: <%= FormatMoney(itemEngravingPrice) %></span>
                                <% End If %>
                            </div>
                        </div>
                    </div>
                    <div class="col-price">
                        <span class="price"><%= FormatMoney(rsCart("UnitPrice")) %></span>
                    </div>
                    <div class="col-quantity">
                        <div class="quantity-input">
                            <button type="button" class="qty-btn minus" onclick="updateQuantity(<%= rsCart("CartID") %>, -1)">-</button>
                            <input type="text" value="<%= rsCart("Quantity") %>" readonly class="qty-value">
                            <button type="button" class="qty-btn plus" onclick="updateQuantity(<%= rsCart("CartID") %>, 1)">+</button>
                        </div>
                    </div>
                    <div class="col-subtotal">
                        <span class="subtotal"><%= FormatMoney(subtotal) %></span>
                    </div>
                    <div class="col-action">
                        <button type="button" class="btn-remove" onclick="removeItem(<%= rsCart("CartID") %>)" title="删除">
                            <i class="fas fa-trash-alt"></i>
                        </button>
                    </div>
                </div>
                <%
                    rsCart.MoveNext
                Loop
                rsCart.Close
                Set rsCart = Nothing
                
                ' 计算应付总额（商品金额 + 刻字费用）
                Dim grandTotal
                grandTotal = cartTotal + totalEngravingFee
                %>
            </div>

            <!-- 购物车摘要 -->
            <div class="cart-summary">
                <div class="summary-card">
                    <h3><% If FEATURE_I18N Then %><%= T("cart_summary_title", Empty) %><% Else %>订单摘要<% End If %></h3>
                    <div class="summary-row">
                        <span><% If FEATURE_I18N Then %><%= T("cart_summary_count", Empty) %><% Else %>商品数量<% End If %>:</span>
                        <span id="selectedItemCount"><%= cartCount %> <% If FEATURE_I18N Then %><%= T("cart_footer_items", Empty) %><% Else %>件<% End If %></span>
                    </div>
                    <div class="summary-row">
                        <span><% If FEATURE_I18N Then %><%= T("cart_summary_amount", Empty) %><% Else %>商品金额<% End If %>:</span>
                        <span id="cartSubtotal"><%= FormatMoney(cartTotal) %></span>
                    </div>
                    <% If totalEngravingFee > 0 Then %>
                    <div class="summary-row">
                        <span><% If FEATURE_I18N Then %><%= T("cart_summary_engraving", Empty) %><% Else %>刻字费用<% End If %>:</span>
                        <span id="engravingFee"><%= FormatMoney(totalEngravingFee) %></span>
                    </div>
                    <% End If %>
                    <div class="summary-row">
                        <span><% If FEATURE_I18N Then %><%= T("cart_summary_shipping", Empty) %><% Else %>运费<% End If %>:</span>
                        <span id="shippingFee"><%
                        Dim displayShippingFee
                        If cartTotal <= 0 Then
                            Response.Write "<span>" & T("cart_summary_no_items", Empty) & "</span>"
                        ElseIf grandTotal >= FREE_SHIPPING_AMOUNT Then
                            Response.Write "<span>" & T("cart_summary_free_shipping", Empty) & "</span>"
                        Else
                            Response.Write FormatMoney(SHIPPING_FEE)
                        End If
                        %></span>
                    </div>
                    <div class="shipping-tip">
                        <i class="fas fa-info-circle"></i>
                        <span id="shippingTip"><% If FEATURE_I18N Then %><%= T("cart_summary_shipping_tip", Array(FREE_SHIPPING_AMOUNT)) %><% Else %>满<%= FREE_SHIPPING_AMOUNT %>元免运费<% End If %></span>
                    </div>
                    <div class="summary-total">
                        <span><% If FEATURE_I18N Then %><%= T("cart_summary_total", Empty) %><% Else %>应付总额<% End If %>:</span>
                        <span class="total-amount" id="cartTotal"><%= FormatMoney(grandTotal) %></span>
                    </div>
                    <a href="javascript:void(0)" onclick="goCheckout()" class="btn btn-primary btn-lg btn-block">
                        <i class="fas fa-credit-card"></i> <% If FEATURE_I18N Then %><%= T("cart_btn_checkout", Empty) %><% Else %>去结算<% End If %>
                    </a>
                    <a href="/products.asp" class="btn btn-outline btn-block">
                        <i class="fas fa-arrow-left"></i> <% If FEATURE_I18N Then %><%= T("cart_btn_continue", Empty) %><% Else %>继续购物<% End If %>
                    </a>
                </div>
                
                <div class="promo-code">
                    <h4><% If FEATURE_I18N Then %><%= T("cart_promo_title", Empty) %><% Else %>优惠码<% End If %></h4>
                    <div class="promo-input">
                        <input type="text" id="promoCode" placeholder="<% If FEATURE_I18N Then %><%= T("cart_promo_placeholder", Empty) %><% Else %>输入优惠码<% End If %>">
                        <button type="button" class="btn btn-sm" onclick="applyPromo()"><% If FEATURE_I18N Then %><%= T("cart_promo_btn", Empty) %><% Else %>应用<% End If %></button>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- 底部操作栏 -->
        <div class="cart-footer">
            <label class="select-all">
                <input type="checkbox" id="selectAllBottom" checked onchange="toggleAllSelection()"> <% If FEATURE_I18N Then %><%= T("cart_select_all", Empty) %><% Else %>全选<% End If %>
            </label>
            <button type="button" class="btn btn-text" onclick="clearCart()">
                <i class="fas fa-trash"></i> <% If FEATURE_I18N Then %><%= T("cart_clear", Empty) %><% Else %>清空购物车<% End If %>
            </button>
            <div class="footer-summary">
                <span><% If FEATURE_I18N Then %><%= T("cart_footer_selected", Empty) %><% Else %>已选<% End If %> <strong id="selectedCount"><%= cartCount %></strong> <% If FEATURE_I18N Then %><%= T("cart_footer_items", Empty) %><% Else %>件<% End If %></span>
                <span><% If FEATURE_I18N Then %><%= T("cart_footer_total", Empty) %><% Else %>合计<% End If %>: <strong class="total" id="footerTotal"><%= FormatMoney(IIf(grandTotal >= FREE_SHIPPING_AMOUNT Or grandTotal <= 0, grandTotal, grandTotal + SHIPPING_FEE)) %></strong></span>
<a href="javascript:void(0)" onclick="goCheckout()" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then %><%= T("cart_btn_checkout", Empty) %><% Else %>去结算<% End If %></a>
            </div>
        </div>
        
        <%
        Else
        %>
        <div class="cart-empty">
            <i class="fas fa-shopping-cart"></i>
            <h2><% If FEATURE_I18N Then %><%= T("cart_empty_title", Empty) %><% Else %>购物车是空的<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("cart_empty_desc", Empty) %><% Else %>快去挑选您喜欢的香水吧！<% End If %></p>
            <a href="/products.asp" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then %><%= T("cart_empty_btn", Empty) %><% Else %>开始选购<% End If %></a>
        </div>
        <%
        End If
        %>
    </div>
</div>

<script>
// 将ASP常量值传给JavaScript
var FREE_SHIPPING_AMOUNT = <%=FREE_SHIPPING_AMOUNT%>;
var SHIPPING_FEE = <%=SHIPPING_FEE%>;
var TOTAL_ENGRAVING_FEE = <%=totalEngravingFee%>;
// i18n JS strings
var i18nConfirmRemove = '<% If FEATURE_I18N Then %><%= T("cart_confirm_remove", Empty) %><% Else %>确定要删除这件商品吗？<% End If %>';
var i18nConfirmClear = '<% If FEATURE_I18N Then %><%= T("cart_confirm_clear", Empty) %><% Else %>确定要清空购物车吗？<% End If %>';
var i18nPromoSoon = '<% If FEATURE_I18N Then %><%= T("cart_promo_soon", Empty) %><% Else %>优惠码功能即将上线！<% End If %>';
var i18nAlertSelect = '<% If FEATURE_I18N Then %><%= T("cart_alert_select", Empty) %><% Else %>请至少选择一件商品进行结算<% End If %>';
var i18nUpdateFailed = '<% If FEATURE_I18N Then %><%= T("cart_alert_update_failed", Empty) %><% Else %>更新失败<% End If %>';
var i18nRemoveFailed = '<% If FEATURE_I18N Then %><%= T("cart_alert_remove_failed", Empty) %><% Else %>删除失败<% End If %>';
var i18nClearFailed = '<% If FEATURE_I18N Then %><%= T("cart_alert_clear_failed", Empty) %><% Else %>操作失败<% End If %>';
var i18nNetworkError = '<% If FEATURE_I18N Then %><%= T("cart_alert_network", Empty) %><% Else %>请求失败，请刷新页面重试<% End If %>';
var i18nNoItems = '<% If FEATURE_I18N Then %><%= T("cart_summary_no_items", Empty) %><% Else %>暂无商品<% End If %>';
var i18nFreeShipping = '<% If FEATURE_I18N Then %><%= T("cart_summary_free_shipping", Empty) %><% Else %>免运费<% End If %>';

function updateQuantity(cartId, delta) {
    $.ajax({
        url: '/api/cart_update.asp',
        type: 'POST',
        data: {
            cartId: cartId,
            delta: delta,
            csrf_token: csrfToken
        },
        dataType: 'json',
        success: function(response) {
            if (response.success) {
                location.reload();
            } else {
                alert(response.message || i18nUpdateFailed);
            }
        },
        error: function() {
            alert(i18nNetworkError);
        }
    });
}

function removeItem(cartId) {
    if (confirm(i18nConfirmRemove)) {
        $.ajax({
            url: '/api/cart_remove.asp',
            type: 'POST',
            data: {
                cartId: cartId,
                csrf_token: csrfToken
            },
            dataType: 'json',
            success: function(response) {
                if (response.success) {
                    location.reload();
                } else {
                    alert(response.message || i18nRemoveFailed);
                }
            },
            error: function() {
                alert(i18nNetworkError);
            }
        });
    }
}

function clearCart() {
    if (confirm(i18nConfirmClear)) {
        $.ajax({
            url: '/api/cart_clear.asp',
            type: 'POST',
            data: {
                csrf_token: csrfToken
            },
            dataType: 'json',
            success: function(response) {
                if (response.success) {
                    location.reload();
                } else {
                    alert(response.message || i18nClearFailed);
                }
            },
            error: function() {
                alert(i18nNetworkError);
            }
        });
    }
}

// 全选/取消全选功能
function toggleAllSelection() {
    // 获取当前点击的复选框的状态
    var currentTarget = $(event.target);
    var isChecked = currentTarget.prop('checked');
    
    // 设置所有商品复选框的状态
    $('.item-checkbox').prop('checked', isChecked);
    
    // 同步两个全选框的状态
    $('#selectAll').prop('checked', isChecked);
    $('#selectAllBottom').prop('checked', isChecked);
    
    updateSelectedCount();
}

// 更新选中商品数量和总计
function updateSelectedCount() {
    var selectedCount = $('.item-checkbox:checked').length;
    $('#selectedCount').text(selectedCount);
    $('#selectedItemCount').text(selectedCount);
    
    // 更新底部合计金额
    var total = 0;
    $('.item-checkbox:checked').each(function() {
        var itemRow = $(this).closest('.cart-item');
        var subtotal = parseFloat(itemRow.data('subtotal')) || 0;
        total += subtotal;
    });
    
    $('#cartSubtotal').text('¥' + total.toFixed(2));
    
    // 计算刻字费用（仅统计选中商品的刻字费用）
    var engravingTotal = 0;
    $('.item-checkbox:checked').each(function() {
        var itemRow = $(this).closest('.cart-item');
        var itemEngFee = parseFloat(itemRow.data('engraving-fee')) || 0;
        engravingTotal += itemEngFee;
    });
    
    // 更新刻字费用显示
    if (engravingTotal > 0) {
        $('#engravingFee').text('¥' + engravingTotal.toFixed(2));
        $('#engravingFee').closest('.summary-row').show();
    } else {
        $('#engravingFee').closest('.summary-row').hide();
    }
    
    // 计算总金额（含运费和刻字费用）
    var shippingFee = 0;
    if (total > 0 && total < FREE_SHIPPING_AMOUNT) {
        shippingFee = SHIPPING_FEE;
    }
    // 应付总额 = 商品金额 + 刻字费用 + 运费
    var grandTotal = total + engravingTotal + shippingFee;
    $('#cartTotal').text('¥' + grandTotal.toFixed(2));
    $('#footerTotal').text('¥' + grandTotal.toFixed(2));  // 更新底部合计金额（含运费）
    
    // 更新运费显示
    if (total <= 0) {
        $('#shippingFee').text(i18nNoItems);
    } else if (shippingFee > 0) {
        $('#shippingFee').text('¥' + shippingFee.toFixed(2));
    } else {
        $('#shippingFee').text(i18nFreeShipping);
    }
    
    // 同步两个全选框的状态
    var allCheckboxes = $('.item-checkbox');
    var checkedBoxes = $('.item-checkbox:checked');
    
    if (allCheckboxes.length === checkedBoxes.length && allCheckboxes.length > 0) {
        $('#selectAll, #selectAllBottom').prop('checked', true);
    } else {
        $('#selectAll, #selectAllBottom').prop('checked', false);
    }
}

// 初始化时计算选中状态
$(document).ready(function() {
    updateSelectedCount();
});

// 监听单个商品复选框变化
$(document).on('change', '.item-checkbox', function() {
    updateSelectedCount();
});

function applyPromo() {
    var code = $('#promoCode').val();
    if (code) {
        alert(i18nPromoSoon);
    }
}

// 去结算 - 只结算选中的商品
function goCheckout() {
    var selectedIds = [];
    $('.item-checkbox:checked').each(function() {
        var cartItem = $(this).closest('.cart-item');
        var cartId = cartItem.data('id');
        if (cartId) {
            selectedIds.push(cartId);
        }
    });
    
    if (selectedIds.length === 0) {
        alert(i18nAlertSelect);
        return;
    }
    
    // 将选中的商品ID传递到结算页面
    window.location.href = '/checkout.asp?cart_ids=' + selectedIds.join(',');
}
</script>

<!-- V18 AI 搭配推荐 -->
<%
If FEATURE_AI_RECOMMENDATIONS Then
    Dim cartUserId, cartRsRecommend
    cartUserId = Session("UserID")
    If cartUserId <> "" And Not IsNull(cartUserId) And IsNumeric(cartUserId) Then
        Set cartRsRecommend = RE_GetPersonalizedProducts(CLng(cartUserId), 6)
    Else
        Set cartRsRecommend = RE_GetTrendingNow(6)
    End If
%>
<div class="recommendation-section">
    <div class="container">
        <h2 class="section-title"><i class="fas fa-gem"></i> <% If FEATURE_I18N Then %><%= T("cart_recommend_title", Empty) %><% Else %>搭配推荐<% End If %></h2>
        <p class="section-desc"><% If FEATURE_I18N Then %><%= T("cart_recommend_desc", Empty) %><% Else %>这些香氛可能与您的购物车搭配得宜<% End If %></p>
        <% Call RE_RenderRecommendationsSafe(cartRsRecommend, "cart-rec-grid", True, "添加商品后为您推荐搭配~") %>
        <%
        If Not cartRsRecommend Is Nothing Then
            cartRsRecommend.Close
            Set cartRsRecommend = Nothing
        End If
        %>
    </div>
</div>
<% End If %>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>