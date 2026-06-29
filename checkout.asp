<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V16.0 结算页 - 主调度器
' 原文件 99KB/1827行 → 调度器 + 8个子模块
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/dal_admin.asp"-->
<!--#include file="includes/dal_users.asp"-->
<!--#include file="includes/dal_checkout.asp"-->
<!--#include file="includes/payment_handler.asp"-->
<!--#include file="includes/cost_engine.asp"-->
<!--#include file="includes/member_utils.asp"-->
<!--#include file="includes/points_engine.asp"-->
<%
Call OpenConnection()

' V14: 会员登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If
%>
<!--#include file="checkout_utils.asp"-->
<!--#include file="checkout_cart_loader.asp"-->
<!--#include file="checkout_address_handler.asp"-->
<!--#include file="checkout_order_creator.asp"-->
<!--#include file="checkout_payment_processor.asp"-->
<!--#include file="includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <a href="/cart.asp"><% If FEATURE_I18N Then %><%= T("cart_breadcrumb", Empty) %><% Else %>购物车<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then %><%= T("checkout_breadcrumb", Empty) %><% Else %>结算<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="checkout-page">
        <h1 class="page-title"><i class="fas fa-credit-card"></i> <% If FEATURE_I18N Then %><%= T("checkout_title", Empty) %><% Else %>订单结算<% End If %></h1>
        
        <% If Session("ErrorMessage") <> "" Then %>
        <div class="alert alert-error">
            <%= Session("ErrorMessage") %>
            <% Session("ErrorMessage") = "" %>
        </div>
        <% End If %>
        
        <!--#include file="checkout_items_display.asp"-->
        <!--#include file="checkout_summary.asp"-->
        <!--#include file="checkout_address_modal.asp"-->
            </div><!-- .checkout-summary -->
        </div><!-- .checkout-content -->
    </div><!-- .checkout-page -->
</div><!-- .container -->

<script src="/js/area_data.js"></script>
<script src="/js/checkout.js?v=16.0"></script>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
