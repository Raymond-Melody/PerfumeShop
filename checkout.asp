<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V14.6 结算页 - 主调度器
' 原文件 99KB/1827行 → 调度器 + 8个子模块
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/payment_handler.asp"-->
<!--#include file="includes/cost_engine.asp"-->
<!--#include file="includes/member_utils.asp"-->
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
        
        <!--#include file="checkout_items_display.asp"-->
        <!--#include file="checkout_summary.asp"-->
        <!--#include file="checkout_address_modal.asp"-->
            </div><!-- .checkout-summary -->
        </div><!-- .checkout-content -->
    </div><!-- .checkout-page -->
</div><!-- .container -->

<script src="/js/area_data.js"></script>
<script src="/js/checkout.js?v=14.6"></script>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
