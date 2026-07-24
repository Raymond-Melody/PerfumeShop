<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V14.6 产品设置 - 主调度器
' 原文件 169KB/3422行 → 调度器 + 7个子模块
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/common_utils.asp"-->
<!--#include file="../../includes/product_type_utils.asp"-->
<!--#include file="../../includes/cost_engine.asp"-->
<%
Call OpenConnection()

' ========== 预加载产品类型数据 ==========
Dim allProductTypes
allProductTypes = GetAllProductTypes()

' ========== 获取当前Tab ==========
Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "products"

' ========== 处理POST请求 ==========
' 注意：SSI #include 在 ASP 执行前由 IIS 预处理，
' 因此不能放在条件语句中。POST检查在子模块内部完成。
%>
<!--#include file="product_settings_post_handler.asp"-->
<!--#include file="product_settings_data_loader.asp"-->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>产品设置 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/admin/css/admin-theme.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2 class="page-title"><i class="fas fa-box-open"></i> 产品设置</h2>
                <div class="breadcrumb">
                    <a href="index.asp">技术中心</a> / <span>产品设置</span>
                </div>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success">
            <i class="fas fa-check-circle"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <% If Not isManager Then %>
        <div class="readonly-notice">
            <i class="fas fa-info-circle"></i> 您当前为技术人员，部分管理功能（删除、创建类型）需要技术经理权限
        </div>
        <% End If %>
        
        <!-- Tab导航 -->
        <div class="tab-nav">
            <a href="?tab=products" class="tab-link <%= IIf(currentTab = "products", "active", "") %>">
                <i class="fas fa-box"></i> 产品管理
            </a>
            <a href="?tab=types" class="tab-link <%= IIf(currentTab = "types", "active", "") %>">
                <i class="fas fa-tags"></i> 类型配置
            </a>
            <a href="?tab=ratio" class="tab-link <%= IIf(currentTab = "ratio", "active", "") %>">
                <i class="fas fa-percentage"></i> 香调配比参数
            </a>
        </div>
        
        <!-- 标签页内容（子模块） -->
        <!--#include file="product_settings_products_tab.asp"-->
        <!--#include file="product_settings_types_tab.asp"-->
        <!--#include file="product_settings_ratio_tab.asp"-->
        <% End If ' 关闭 currentTab 条件链 %>
    </div>
    
    <!-- 模态框（子模块） -->
    <!--#include file="product_settings_modals.asp"-->

    <!-- 服务端注入变量（必须在外部JS之前） -->
    <script nonce="<%= Session("csp_nonce") %>">
        var minTopPercent = <%=minTopPercent%>;
        var minMiddlePercent = <%=minMiddlePercent%>;
        var minBasePercent = <%=minBasePercent%>;
        <%=formulaDataJson%>
        <%=recipeDataJson%>
    </script>
    <!-- 公共脚本 -->
    <script src="/admin/js/admin-common.js"></script>
    <!-- 页面脚本 -->
    <script src="/admin/js/product-settings.js"></script>
</body>
</html>
