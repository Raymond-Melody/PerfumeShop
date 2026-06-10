<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/product_type_utils.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection

' 一次性数据修复：修复ProductType字段包含重复值的问题
Dim rsFixProductType, fixProductId, fixProductTypeValue, fixFirstValue, fixUpdateSql
Set rsFixProductType = ExecuteQuery("SELECT ProductID, ProductType FROM Products WHERE ProductType LIKE '%,%'")
If Not rsFixProductType Is Nothing Then
    Do While Not rsFixProductType.EOF
        fixProductId = rsFixProductType("ProductID")
        fixProductTypeValue = Trim(rsFixProductType("ProductType") & "")
        If InStr(fixProductTypeValue, ",") > 0 Then
            fixFirstValue = Trim(Split(fixProductTypeValue, ",")(0))
            fixUpdateSql = "UPDATE Products SET ProductType = '" & SafeSQL(fixFirstValue) & "' WHERE ProductID = " & fixProductId
            Call ExecuteNonQuery(fixUpdateSql)
        End If
        rsFixProductType.MoveNext
    Loop
    rsFixProductType.Close
    Set rsFixProductType = Nothing
End If

' 预加载商品类型数据（必须在任何Recordset打开之前获取，避免MARS限制）
Dim allProductTypes
allProductTypes = GetAllProductTypes()

' 构建类型属性映射（用于JS判断）
Dim typeRequiresReview, typeRequiresRatio, typeIsActive, ptIdx
Dim ptCode, ptName, ptIsActive
typeRequiresReview = ""
typeRequiresRatio = ""
typeIsActive = ""  ' 存储类型激活状态映射，用于JS判断
If IsArray(allProductTypes) Then
    For ptIdx = 0 To UBound(allProductTypes, 1)
        ' RequiresReview (index 5)
        If allProductTypes(ptIdx, 5) Then
            If typeRequiresReview <> "" Then typeRequiresReview = typeRequiresReview & ","
            typeRequiresReview = typeRequiresReview & "'" & allProductTypes(ptIdx, 0) & "'"
        End If
        ' RequiresRatio (index 6)
        If allProductTypes(ptIdx, 6) Then
            If typeRequiresRatio <> "" Then typeRequiresRatio = typeRequiresRatio & ","
            typeRequiresRatio = typeRequiresRatio & "'" & allProductTypes(ptIdx, 0) & "'"
        End If
        ' IsActive (index 8) - 用于前台状态判断
        If allProductTypes(ptIdx, 8) Then
            If typeIsActive <> "" Then typeIsActive = typeIsActive & ","
            typeIsActive = typeIsActive & "'" & allProductTypes(ptIdx, 0) & "'"
        End If
    Next
End If

' 处理表单提交
Dim action, productId
Dim pTypeDisplay, baseIng, reviewClass

' 处理表单提交 - 运营后台仅保留上下架切换功能
action = Request.Form("action")

If action = "toggleActive" Then
    ' 上下架切换
    productId = Request.Form("productId")
    Dim currentActive, newActive, toggleSql
    
    If IsNumeric(productId) Then
        ' 获取当前状态
        Dim rsCurrent
        Set rsCurrent = ExecuteQuery("SELECT IsActive FROM Products WHERE ProductID = " & CInt(productId))
        If Not rsCurrent Is Nothing Then
            If Not rsCurrent.EOF Then
                currentActive = (rsCurrent("IsActive") <> 0)
                newActive = Not currentActive
                rsCurrent.Close
                Set rsCurrent = Nothing
                
                toggleSql = "UPDATE Products SET IsActive = " & IIf(newActive, 1, 0) & " WHERE ProductID = " & CInt(productId)
                If ExecuteNonQuery(toggleSql) Then
                    Response.Redirect "products.asp?msg=" & IIf(newActive, "商品已上架", "商品已下架")
                Else
                    Response.Write "<script>alert('操作失败：" & Replace(Session("LastDBError"), "'", "\'") & "');</script>"
                End If
            End If
        End If
    End If
End If

' 处理前台状态筛选参数
Dim frontStatusFilter, disabledTypeCodes, disabledTypeListStr
frontStatusFilter = Request.QueryString("frontStatus")
disabledTypeCodes = ""  ' 存储被禁用的类型代码，用于筛选
disabledTypeListStr = ""  ' 用于SQL IN条件

' 构建被禁用类型的列表
If IsArray(allProductTypes) Then
    For ptIdx = 0 To UBound(allProductTypes, 1)
        If Not allProductTypes(ptIdx, 8) Then  ' IsActive = False (被禁用)
            If disabledTypeCodes <> "" Then
                disabledTypeCodes = disabledTypeCodes & ","
                disabledTypeListStr = disabledTypeListStr & ","
            End If
            disabledTypeCodes = disabledTypeCodes & allProductTypes(ptIdx, 0)
            disabledTypeListStr = disabledTypeListStr & "'" & allProductTypes(ptIdx, 0) & "'"
        End If
    Next
End If

' 获取商品列表（带筛选）
Dim rsProducts, sql, whereClause
whereClause = ""

Select Case frontStatusFilter
    Case "showing"  ' 展示中
        ' IsActive<>0 AND (ProductType的类型IsActive<>0) AND (不需审核 OR ReviewStatus='Approved')
        whereClause = "WHERE IsActive <> 0"
        If disabledTypeListStr <> "" Then
            whereClause = whereClause & " AND ProductType NOT IN (" & disabledTypeListStr & ")"
        End If
        ' 需要审核的类型必须Approved
        If typeRequiresReview <> "" Then
            whereClause = whereClause & " AND (ProductType NOT IN (" & typeRequiresReview & ") OR ReviewStatus = 'Approved')"
        End If
    Case "inactive"  ' 已下架
        whereClause = "WHERE IsActive = 0"
    Case "typeDisabled"  ' 类型已关闭
        If disabledTypeListStr <> "" Then
            whereClause = "WHERE ProductType IN (" & disabledTypeListStr & ")"
        Else
            ' 没有被禁用的类型，返回空结果
            whereClause = "WHERE 1=0"
        End If
    Case "pending"  ' 待审核
        whereClause = "WHERE ReviewStatus = 'Pending'"
    Case Else
        ' 全部，不加筛选条件
        whereClause = ""
End Select

sql = "SELECT * FROM Products " & whereClause & " ORDER BY ProductID DESC"
Set rsProducts = ExecuteQuery(sql)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>商品上下架管理 - 营运管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 深色主题 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        .main-content {
            color: #e0e0e0;
        }
        
        /* 页面标题区 */
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .page-title {
            font-size: 24px;
            color: #fff;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .page-title i { color: #00bcd4; }
        .breadcrumb { font-size: 13px; color: #888; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .breadcrumb a:hover { text-decoration: underline; }
        
        /* 提示横幅 */
        .info-banner {
            background: rgba(0,188,212,0.1);
            border-left: 4px solid #00bcd4;
            color: #80deea;
            padding: 12px 18px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 14px;
        }
        .info-banner a { color: #00bcd4; text-decoration: underline; }
        .info-banner a:hover { color: #26c6da; }
        
        /* 成功消息 */
        .alert-success {
            background: rgba(76,175,80,0.15);
            border-left: 4px solid #4caf50;
            color: #81c784;
            padding: 12px 18px;
            border-radius: 6px;
            margin-bottom: 20px;
        }
        
        /* 筛选栏 */
        .filter-bar {
            display: flex;
            align-items: center;
            gap: 15px;
            margin-bottom: 25px;
            flex-wrap: wrap;
            padding: 15px 20px;
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        .filter-label {
            color: #888;
            font-size: 13px;
            font-weight: 600;
            white-space: nowrap;
        }
        .filter-select {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 6px;
            color: #e0e0e0;
            padding: 8px 12px;
            font-size: 13px;
            cursor: pointer;
            min-width: 140px;
        }
        .filter-select:focus { border-color: #00bcd4; outline: none; }
        .filter-select option { background: #2d2d44; color: #e0e0e0; }
        
        /* 产品卡片网格 */
        .product-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        }
        .product-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            display: flex;
            flex-direction: column;
        }
        .product-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        .product-card-image {
            position: relative;
            width: 100%;
            padding-top: 100%;
            border-radius: 8px;
            overflow: hidden;
            margin-bottom: 15px;
            background: rgba(0,0,0,0.3);
        }
        .product-card-image img {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            object-fit: cover;
            transition: transform 0.3s ease;
        }
        .product-card-image .no-image {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            color: #666;
            font-size: 36px;
        }
        .product-card:hover .product-card-image img {
            transform: scale(1.05);
        }
        .product-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 10px;
        }
        .product-title {
            font-size: 15px;
            font-weight: 600;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 8px;
            flex: 1;
            min-width: 0;
        }
        .product-title span {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .product-title i { color: #00bcd4; flex-shrink: 0; }
        .product-id {
            font-size: 11px;
            color: #999;
            background: rgba(0,0,0,0.3);
            padding: 2px 8px;
            border-radius: 4px;
            flex-shrink: 0;
        }
        .product-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            margin-bottom: 10px;
            align-items: center;
        }
        .product-price {
            font-size: 18px;
            font-weight: 700;
            color: #00bcd4;
            margin-bottom: 10px;
        }
        .product-desc {
            font-size: 12px;
            color: #888;
            line-height: 1.5;
            margin-bottom: 10px;
            flex: 1;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }
        .product-type-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 10px;
            font-weight: 600;
            text-transform: uppercase;
        }
        .product-type-badge.fixed { background: rgba(33,150,243,0.2); color: #2196f3; }
        .product-type-badge.custom { background: rgba(76,175,80,0.2); color: #4caf50; }
        .product-type-badge.kol { background: rgba(156,39,176,0.2); color: #9c27b0; }
        
        /* 状态标签 */
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 500;
        }
        .status-active { background: rgba(76,175,80,0.2); color: #4caf50; }
        .status-inactive { background: rgba(244,67,54,0.2); color: #ef5350; }
        .status-pending { background: rgba(255,152,0,0.2); color: #ff9800; }
        .status-approved { background: rgba(76,175,80,0.2); color: #4caf50; }
        .status-disabled { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        
        .product-footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 12px;
            border-top: 1px solid rgba(255,255,255,0.05);
            margin-top: auto;
            gap: 8px;
        }
        .product-footer .action-btns {
            display: flex;
            gap: 8px;
            align-items: center;
        }
        
        /* 空状态 */
        .empty-state {
            grid-column: 1 / -1;
            text-align: center;
            padding: 60px 20px;
            color: #666;
        }
        .empty-state i {
            font-size: 48px;
            margin-bottom: 15px;
            color: #555;
        }
        .empty-state h3 { color: #888; margin-bottom: 10px; }
        .empty-state p { color: #666; font-size: 14px; }
        
        /* 模态框深色 */
        .admin-modal-content {
            background: #1e1e32 !important;
            color: #e0e0e0 !important;
        }
        .admin-modal-content .admin-modal-header,
        .admin-modal-content .admin-modal-footer {
            background: rgba(0,0,0,0.2) !important;
            border-color: rgba(255,255,255,0.05) !important;
        }
        .admin-modal-content .admin-modal-title { color: #fff !important; }
        .admin-modal-content .admin-modal-close { color: #888 !important; }
        .admin-modal-content .admin-modal-close:hover { color: #fff !important; }
        .admin-modal-content .admin-form-label { color: #aaa !important; }
        .admin-modal-content p {
            background: rgba(255,255,255,0.05) !important;
            color: #e0e0e0 !important;
            border: 1px solid rgba(255,255,255,0.08);
        }
        .admin-modal-content .alert-info {
            background: rgba(255,193,7,0.1) !important;
            border-left-color: #ffc107 !important;
            color: #ffc107 !important;
        }
        .admin-modal-content .alert-info .alert-link {
            color: #ffc107 !important;
            text-decoration: underline;
            font-weight: 600;
        }
        .admin-modal-content .alert-info .alert-link:hover {
            color: #ffd54f !important;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .product-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .product-grid { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .filter-select { width: 100%; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2 class="page-title"><i class="fas fa-box"></i> 商品上下架管理</h2>
                <div class="breadcrumb">
                    <a href="index.asp">运营中心</a> / <span>商品上下架</span>
                </div>
            </div>
            <a href="../techcenter/product_settings.asp" class="btn btn-outline">
                <i class="fas fa-external-link-alt"></i> 产品技术管理中心
            </a>
        </div>
        
        <div class="info-banner">
            <i class="fas fa-info-circle"></i>
            <strong>提示：</strong>运营后台仅管理产品上下架状态，产品创建和修改请前往<a href="../techcenter/product_settings.asp">产品技术管理中心</a>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert-success">
            <i class="fas fa-check-circle"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <span class="filter-label"><i class="fas fa-filter"></i> 前台状态筛选：</span>
            <form method="get" style="display: flex; gap: 10px; align-items: center; flex-wrap: wrap;">
                <select name="frontStatus" class="filter-select" onchange="this.form.submit()">
                    <option value="" <%= IIf(frontStatusFilter = "", "selected", "") %>>全部商品</option>
                    <option value="showing" <%= IIf(frontStatusFilter = "showing", "selected", "") %>>展示中</option>
                    <option value="inactive" <%= IIf(frontStatusFilter = "inactive", "selected", "") %>>已下架</option>
                    <option value="typeDisabled" <%= IIf(frontStatusFilter = "typeDisabled", "selected", "") %>>类型已关闭</option>
                    <option value="pending" <%= IIf(frontStatusFilter = "pending", "selected", "") %>>待审核</option>
                </select>
                <% If frontStatusFilter <> "" Then %>
                <a href="products.asp" class="btn btn-outline btn-sm">
                    <i class="fas fa-times"></i> 清除筛选
                </a>
                <% End If %>
            </form>
        </div>
        
        <!-- 产品卡片网格 -->
        <div class="product-grid">
            <%
            Dim pTypeValue, ptIdxDisplay, needsReviewDisplay, ptIdxReview
            Dim frontStatusText, frontStatusClass, frontStatusTitle
            Dim typeIsActiveFlag, typeRequiresReviewFlag, ptIdxFront
            Dim hasProducts
            hasProducts = False
            If Not rsProducts Is Nothing Then
                If Not rsProducts.EOF Then hasProducts = True
            End If
            
            If hasProducts Then
                Do While Not rsProducts.EOF
                    pTypeDisplay = ""
                    reviewStatus = ""
                    reviewClass = ""
                    
                    pTypeValue = Trim(rsProducts("ProductType") & "")
                    
                    ' 从预加载的类型数组中查找显示名称
                    pTypeDisplay = ""
                    If IsArray(allProductTypes) Then
                        For ptIdxDisplay = 0 To UBound(allProductTypes, 1)
                            If allProductTypes(ptIdxDisplay, 0) = pTypeValue Then
                                pTypeDisplay = allProductTypes(ptIdxDisplay, 1)  ' DisplayName
                                Exit For
                            End If
                        Next
                    End If
                    If pTypeDisplay = "" Then pTypeDisplay = "未知(" & pTypeValue & ")"
                    
                    baseIng = ""
                    On Error Resume Next
                    baseIng = Trim(rsProducts("BaseIngredients") & "")
                    On Error GoTo 0
                    
                    ' 设置类型badge样式
                    Dim pBadgeClass
                    Select Case pTypeValue
                        Case "Fixed": pBadgeClass = "fixed"
                        Case "Custom": pBadgeClass = "custom"
                        Case "KOL": pBadgeClass = "kol"
                        Case Else: pBadgeClass = ""
                    End Select
                    
                    ' 审核状态
                    reviewStatus = ""
                    On Error Resume Next
                    reviewStatus = rsProducts("ReviewStatus") & ""
                    On Error GoTo 0
                    Select Case reviewStatus
                        Case "Pending": reviewClass = "status-pending"
                        Case "Approved": reviewClass = "status-approved"
                        Case "Rejected": reviewClass = "status-inactive"
                        Case Else: reviewClass = ""
                    End Select
                    needsReviewDisplay = False
                    If IsArray(allProductTypes) Then
                        For ptIdxReview = 0 To UBound(allProductTypes, 1)
                            If allProductTypes(ptIdxReview, 0) = pTypeValue Then
                                needsReviewDisplay = allProductTypes(ptIdxReview, 5)
                                Exit For
                            End If
                        Next
                    End If
                    
                    ' 前台状态判断
                    Dim productIsActive, prodReviewStatus
                    typeIsActiveFlag = False
                    typeRequiresReviewFlag = False
                    productIsActive = False
                    prodReviewStatus = ""
                    On Error Resume Next
                    productIsActive = (rsProducts("IsActive") <> 0)
                    prodReviewStatus = rsProducts("ReviewStatus") & ""
                    On Error GoTo 0
                    If IsArray(allProductTypes) Then
                        For ptIdxFront = 0 To UBound(allProductTypes, 1)
                            If allProductTypes(ptIdxFront, 0) = pTypeValue Then
                                typeIsActiveFlag = allProductTypes(ptIdxFront, 8)
                                typeRequiresReviewFlag = allProductTypes(ptIdxFront, 5)
                                Exit For
                            End If
                        Next
                    End If
                    
                    If Not typeIsActiveFlag Then
                        frontStatusText = "类型已关闭"
                        frontStatusClass = "status-disabled"
                        frontStatusTitle = "该商品所属类型已被禁用，商品将不在前台显示"
                    ElseIf Not productIsActive Then
                        frontStatusText = "已下架"
                        frontStatusClass = "status-inactive"
                        frontStatusTitle = "商品已下架"
                    ElseIf typeRequiresReviewFlag Then
                        Select Case prodReviewStatus
                            Case "Pending": frontStatusText = "待审核": frontStatusClass = "status-pending": frontStatusTitle = "商品待审核，审核通过后才在前台展示"
                            Case "Rejected": frontStatusText = "已驳回": frontStatusClass = "status-inactive": frontStatusTitle = "商品审核被驳回，请修改后重新提交"
                            Case "Approved": frontStatusText = "展示中": frontStatusClass = "status-active": frontStatusTitle = "商品正在前台正常展示"
                            Case Else: frontStatusText = "待审核": frontStatusClass = "status-pending": frontStatusTitle = "商品待审核，审核通过后才在前台展示"
                        End Select
                    Else
                        frontStatusText = "展示中"
                        frontStatusClass = "status-active"
                        frontStatusTitle = "商品正在前台正常展示"
                    End If
                    
                    ' 图片URL
                    Dim imgUrl
                    imgUrl = ""
                    On Error Resume Next
                    imgUrl = Trim(rsProducts("ImageURL") & "")
                    On Error GoTo 0
                    If imgUrl = "" Then imgUrl = "/images/default-product.svg"
            %>
            <div class="product-card">
                <div class="product-card-image">
                    <img src="<%= HTMLEncode(imgUrl) %>" alt="<%= HTMLEncode(rsProducts("ProductName")) %>" onerror="this.style.display='none';this.parentElement.querySelector('.no-image').style.display='block';">
                    <i class="fas fa-image no-image" style="display:none;"></i>
                </div>
                
                <div class="product-header">
                    <div class="product-title">
                        <i class="fas fa-box-open"></i>
                        <span title="<%= HTMLEncode(rsProducts("ProductName")) %>"><%= HTMLEncode(rsProducts("ProductName")) %></span>
                    </div>
                    <span class="product-id">#<%= rsProducts("ProductID") %></span>
                </div>
                
                <div class="product-meta">
                    <span class="product-type-badge <%= pBadgeClass %>">
                        <% Select Case pTypeValue
                            Case "Fixed" %><i class="fas fa-box"></i> 品牌定香
                        <%  Case "Custom" %><i class="fas fa-paint-brush"></i> 用户定制
                        <%  Case "KOL" %><i class="fas fa-star"></i> KOL推荐
                        <%  Case Else %><%= HTMLEncode(pTypeDisplay) %>
                        <% End Select %>
                    </span>
                    
                    <% If needsReviewDisplay Then
                        Select Case reviewStatus
                            Case "Pending": Response.Write "<span class='status-badge status-pending'><i class='fas fa-clock'></i> 待审核</span>"
                            Case "Approved": Response.Write "<span class='status-badge status-approved'><i class='fas fa-check'></i> 已通过</span>"
                            Case "Rejected": Response.Write "<span class='status-badge status-inactive'><i class='fas fa-times'></i> 已驳回</span>"
                            Case Else: Response.Write "<span class='status-badge status-pending'><i class='fas fa-clock'></i> 待审核</span>"
                        End Select
                    End If %>
                    
                    <span class="status-badge <%= IIf(rsProducts("IsActive") <> 0, "status-active", "status-inactive") %>">
                        <i class="fas <%= IIf(rsProducts("IsActive") <> 0, "fa-check-circle", "fa-ban") %>"></i>
                        <%= IIf(rsProducts("IsActive") <> 0, "上架", "下架") %>
                    </span>
                    
                    <span class="status-badge <%= frontStatusClass %>" title="<%= frontStatusTitle %>">
                        <i class="fas <%= IIf(frontStatusClass = "status-active", "fa-eye", IIf(frontStatusClass = "status-pending", "fa-clock", IIf(frontStatusClass = "status-inactive", "fa-ban", "fa-minus-circle"))) %>"></i>
                        <%= frontStatusText %>
                    </span>
                </div>
                
                <div class="product-price"><%= FormatMoney(rsProducts("BasePrice")) %></div>
                
                <% If baseIng <> "" Then %>
                <div style="font-size:12px; color:#888; margin-bottom:8px;">
                    <i class="fas fa-leaf" style="color:#00bcd4;"></i> 基香成分已配置
                </div>
                <% End If %>
                
                <div class="product-desc">
                    <%= HTMLEncode(Left(rsProducts("Description") & "", 80)) %>
                </div>
                
                <div class="product-footer">
                    <div class="action-btns">
                        <button class="btn btn-outline btn-sm" onclick="showViewModal(this)" 
                            data-id="<%= rsProducts("ProductID") %>" 
                            data-name="<%= HTMLEncode(rsProducts("ProductName") & "") %>" 
                            data-desc="<%= HTMLEncode(rsProducts("Description") & "") %>" 
                            data-price="<%= rsProducts("BasePrice") %>" 
                            data-type="<%= rsProducts("ProductType") %>"
                            data-active="<%= rsProducts("IsActive") %>" 
                            data-image="<%= HTMLEncode(imgUrl) %>">
                            <i class="fas fa-eye"></i> 详情
                        </button>
                        <% If pTypeValue = "Fixed" Then %>
                        <a href="../purchase/fixed_brand/product_management.asp" class="btn btn-outline btn-sm" title="品牌定香产品管理">
                            <i class="fas fa-truck"></i> 采购管理
                        </a>
                        <% End If %>
                    </div>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="toggleActive">
                        <input type="hidden" name="productId" value="<%= rsProducts("ProductID") %>">
                        <button type="submit" class="btn <%= IIf(rsProducts("IsActive") <> 0, "btn-danger", "btn-primary") %> btn-sm">
                            <i class="fas <%= IIf(rsProducts("IsActive") <> 0, "fa-arrow-down", "fa-arrow-up") %>"></i>
                            <%= IIf(rsProducts("IsActive") <> 0, "下架", "上架") %>
                        </button>
                    </form>
                </div>
            </div>
            <% rsProducts.MoveNext %>
            <% Loop %>
            <% Else %>
            <div class="empty-state">
                <i class="fas fa-box"></i>
                <h3>暂无商品</h3>
                <p>没有符合条件的商品数据</p>
            </div>
            <% End If %>
        </div>
    </div>
    
        <!-- 查看商品详情模态框（只读） -->
    <div id="productModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 id="modalTitle" class="admin-modal-title">商品详情</h3>
                <button class="admin-modal-close" onclick="closeModal()">&times;</button>
            </div>
            <div class="admin-modal-body">
                <div class="admin-form-group">
                    <label class="admin-form-label">商品名称</label>
                    <p id="viewProductName" style="padding: 10px; border-radius: 4px; margin: 0;"></p>
                </div>
                
                <div class="admin-form-group">
                    <label class="admin-form-label">描述</label>
                    <p id="viewDescription" style="padding: 10px; border-radius: 4px; margin: 0; min-height: 60px;"></p>
                </div>
                
                <div class="admin-form-row">
                    <div class="admin-form-col">
                        <div class="admin-form-group">
                            <label class="admin-form-label">商品类型</label>
                            <p id="viewProductType" style="padding: 10px; border-radius: 4px; margin: 0;"></p>
                        </div>
                    </div>
                    <div class="admin-form-col">
                        <div class="admin-form-group">
                            <label class="admin-form-label">基础价格</label>
                            <p id="viewBasePrice" style="padding: 10px; border-radius: 4px; margin: 0;"></p>
                        </div>
                    </div>
                </div>
                
                <div class="admin-form-row">
                    <div class="admin-form-col">
                        <div class="admin-form-group">
                            <label class="admin-form-label">上架状态</label>
                            <p id="viewIsActive" style="padding: 10px; border-radius: 4px; margin: 0;"></p>
                        </div>
                    </div>
                    <div class="admin-form-col">
                        <div class="admin-form-group">
                            <label class="admin-form-label">商品ID</label>
                            <p id="viewProductId" style="padding: 10px; border-radius: 4px; margin: 0;"></p>
                        </div>
                    </div>
                </div>
                
                <div class="admin-form-group">
                    <label class="admin-form-label">图片</label>
                    <div style="padding: 10px; border-radius: 4px;">
                        <img id="viewImage" src="" alt="商品图片" style="max-width: 200px; max-height: 200px; border-radius: 4px;">
                    </div>
                </div>
                
                <div class="alert alert-info" style="margin-top:15px;">
                    <i class="fas fa-info-circle"></i>
                    如需修改商品信息，请前往<a href="../techcenter/product_settings.asp" class="alert-link">产品技术管理中心</a>
                </div>
            </div>
            <div class="admin-modal-footer">
                <button type="button" class="admin-btn admin-btn-outline" onclick="closeModal()">关闭</button>
            </div>
        </div>
    </div>
    
    <script>
        // 商品类型名称映射
        var productTypeNames = {};
        <% 
        If IsArray(allProductTypes) Then
            For ptIdx = 0 To UBound(allProductTypes, 1)
        %>
        productTypeNames['<%= allProductTypes(ptIdx, 0) %>'] = '<%= Replace(Server.HTMLEncode(allProductTypes(ptIdx, 1)), "'", "\'") %>';
        <%
            Next
        End If
        %>
        
        function showViewModal(button) {
            var id = button.getAttribute('data-id');
            var name = button.getAttribute('data-name');
            var desc = button.getAttribute('data-desc');
            var price = button.getAttribute('data-price');
            var type = button.getAttribute('data-type') || 'Custom';
            var active = button.getAttribute('data-active');
            var image = button.getAttribute('data-image');
            
            document.getElementById('modalTitle').textContent = '商品详情 #' + id;
            document.getElementById('viewProductId').textContent = id;
            document.getElementById('viewProductName').textContent = name;
            document.getElementById('viewDescription').textContent = desc || '（无描述）';
            document.getElementById('viewBasePrice').textContent = '¥' + parseFloat(price).toFixed(2);
            document.getElementById('viewProductType').textContent = productTypeNames[type] || type;
            document.getElementById('viewIsActive').innerHTML = active != '0' ? 
                '<span style="color: #27ae60;"><i class="fas fa-check-circle"></i> 上架中</span>' : 
                '<span style="color: #e74c3c;"><i class="fas fa-times-circle"></i> 已下架</span>';
            document.getElementById('viewImage').src = image || '/images/default-product.svg';
            
            document.getElementById('productModal').style.display = 'block';
        }
        
        function closeModal() {
            document.getElementById('productModal').style.display = 'none';
        }
        
        // 点击模态框外部关闭
        window.onclick = function(event) {
            var modal = document.getElementById('productModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }
    </script>
</body>
</html>
<%
If Not rsProducts Is Nothing Then
    rsProducts.Close
    Set rsProducts = Nothing
End If
Call CloseConnection
%>
