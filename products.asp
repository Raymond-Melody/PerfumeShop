<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/dal_products.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<%
Call OpenConnection()

' V14: 会员登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If

' 获取所有已启用的商品类型
Dim activeTypes, activeTypeCodes
activeTypes = GetActiveProductTypes()
activeTypeCodes = GetActiveTypeCodesForSQL()

' 获取参数
Dim keyword, sortBy, page, totalCount, totalPages, productType
keyword = Request.QueryString("keyword")
If keyword <> "" Then
    keyword = Trim(keyword)
    ' V17: 记录搜索历史
    Dim userIdForHistory
    userIdForHistory = Session("UserID")
    If userIdForHistory <> "" Then
        Call DAL_Products_RecordSearch(userIdForHistory, keyword)
    End If
End If
sortBy = Request.QueryString("sort")
page = Request.QueryString("page")
productType = Request.QueryString("type")

If page = "" Or Not IsNumeric(page) Then page = 1
page = CInt(page)
If page < 1 Then page = 1

' 构建查询SQL
Dim whereClause, orderClause
whereClause = " WHERE IsActive <> 0"

' 产品类型筛选
Dim typeFilter
If productType <> "" Then
    typeFilter = BuildProductTypeFilter(productType)
    ' BuildProductTypeFilter 会返回 " AND ProductType='xxx'" 且自动添加审核过滤
    If typeFilter <> "" Then
        whereClause = whereClause & typeFilter
    Else
        ' 类型不存在或未启用，重定向到全部产品
        Response.Redirect "/products.asp"
    End If
Else
    ' 无类型筛选时，只显示已启用类型的商品
    If activeTypeCodes <> "" Then
        whereClause = whereClause & " AND ProductType IN (" & activeTypeCodes & ")"
    End If
End If

If keyword <> "" Then
    ' V18: AI增强搜索 - 相关性评分 + 同义词扩展
    Dim safeKw, expandedWhere
    safeKw = SafeLike(keyword)
    
    ' 基础匹配（名称+描述）
    expandedWhere = " AND (p.ProductName LIKE '%" & safeKw & "%' OR p.Description LIKE '%" & safeKw & "%' OR p.Category LIKE '%" & safeKw & "%')"
    
    ' V18: 同义词扩展搜索
    If FEATURE_AI_SEARCH Then
        Dim synonymKw
        synonymKw = SEARCH_ExpandSearchKeywords(keyword)
        If synonymKw <> "" Then
            expandedWhere = expandedWhere & " OR p.ProductName LIKE '%" & synonymKw & "%' OR p.Description LIKE '%" & synonymKw & "%'"
        End If
        ' 拼音扩展
        Dim pinyinKw
        pinyinKw = SEARCH_PinyinExpand(keyword)
        If pinyinKw <> "" Then
            expandedWhere = expandedWhere & " OR p.ProductName LIKE '%" & pinyinKw & "%'"
        End If
    End If
    
    whereClause = whereClause & expandedWhere
    
    ' V18: 加权相关性排序
    ' 分数: 精确名称匹配=100, 名称开头=80, 名称包含=50, 描述包含=25, 分类包含=15
    orderClause = " ORDER BY (" & _
        "CASE WHEN p.ProductName = '" & safeKw & "' THEN 100 ELSE 0 END + " & _
        "CASE WHEN p.ProductName LIKE '" & safeKw & "%' THEN 80 ELSE 0 END + " & _
        "CASE WHEN p.ProductName LIKE '%" & safeKw & "%' THEN 50 ELSE 0 END + " & _
        "CASE WHEN p.Description LIKE '%" & safeKw & "%' THEN 25 ELSE 0 END + " & _
        "CASE WHEN p.Category LIKE '%" & safeKw & "%' THEN 15 ELSE 0 END" & _
        ") DESC, ISNULL(p.CreatedAt, '2099-12-31') DESC, p.ProductID DESC"
End If

Select Case sortBy
    Case "price_asc"
        orderClause = " ORDER BY BasePrice ASC"
    Case "price_desc"
        orderClause = " ORDER BY BasePrice DESC"
    Case "name"
        orderClause = " ORDER BY ProductName ASC"
    Case Else
        orderClause = " ORDER BY ISNULL(CreatedAt, '2099-12-31') DESC, ProductID DESC"
End Select

' 获取总数
totalCount = GetScalar("SELECT COUNT(*) FROM Products" & whereClause)
If IsNull(totalCount) Or totalCount = "" Then totalCount = 0
totalPages = Int((totalCount + PAGE_SIZE - 1) / PAGE_SIZE)
If totalPages < 1 Then totalPages = 1
If page > totalPages Then page = totalPages
%>
<!--#include file="includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <%
        If keyword <> "" Then
        %>
        <span><% If FEATURE_I18N Then %><%= T("products_breadcrumb_search", Array(HTMLEncode(keyword))) %><% Else %>搜索结果: <%= HTMLEncode(keyword) %><% End If %></span>
        <%
        Else
        %>
        <span><% If FEATURE_I18N Then %><%= T("products_breadcrumb_all", Empty) %><% Else %>全部香水<% End If %></span>
        <%
        End If
        %>
    </div>
</div>

<div class="container">
    <div class="products-page">
        <!-- V11 移动端侧边栏切换 -->
        <!--#include file="includes/mobile_filter.asp"-->
        
        <!-- 侧边栏筛选 -->
        <aside class="sidebar">
            <div class="filter-section">
                <h4><% If FEATURE_I18N Then %><%= T("products_filter_type", Empty) %><% Else %>产品类型<% End If %></h4>
                <ul class="filter-list">
                    <li><a href="/products.asp" class="<%= IIf(productType = "", "active", "") %>"><% If FEATURE_I18N Then %><%= T("products_filter_all", Empty) %><% Else %>全部<% End If %></a></li>
                    <% 
                    Dim fIdx, fCode, fName
                    If IsArray(activeTypes) Then
                        For fIdx = 0 To UBound(activeTypes, 1)
                            fCode = activeTypes(fIdx, 0)
                            fName = activeTypes(fIdx, 1)
                    %>
                    <li><a href="/products.asp?type=<%= Server.URLEncode(fCode) %>" class="<%= IIf(productType = fCode, "active", "") %>"><%= Server.HTMLEncode(GetProductTypeI18nName(fCode, fName, "display")) %></a></li>
                    <%
                        Next
                    End If
                    %>
                </ul>
            </div>
            
            <div class="filter-section">
                <h3><% If FEATURE_I18N Then %><%= T("products_filter_price", Empty) %><% Else %>价格区间<% End If %></h3>
                <ul class="filter-list">
                    <li><a href="/products.asp?price=0-300"><% If FEATURE_I18N Then %><%= T("products_filter_price_1", Empty) %><% Else %>¥300以下<% End If %></a></li>
                    <li><a href="/products.asp?price=300-500"><% If FEATURE_I18N Then %><%= T("products_filter_price_2", Empty) %><% Else %>¥300-500<% End If %></a></li>
                    <li><a href="/products.asp?price=500-800"><% If FEATURE_I18N Then %><%= T("products_filter_price_3", Empty) %><% Else %>¥500-800<% End If %></a></li>
                    <li><a href="/products.asp?price=800"><% If FEATURE_I18N Then %><%= T("products_filter_price_4", Empty) %><% Else %>¥800以上<% End If %></a></li>
                </ul>
            </div>
            
            <div class="filter-section cta-box">
                <h3><% If FEATURE_I18N Then %><%= T("products_filter_cta_title", Empty) %><% Else %>想要专属定制？<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("products_filter_cta_desc", Empty) %><% Else %>打造独一无二的个性香水<% End If %></p>
                <a href="/customize.asp" class="btn btn-primary btn-block"><% If FEATURE_I18N Then %><%= T("products_filter_cta_btn", Empty) %><% Else %>开始定制<% End If %></a>
            </div>
        </aside>

        <!-- 产品列表 -->
        <div class="products-main">
            <div class="products-header">
                <div class="results-info">
                    <%
                    If keyword <> "" Then
                    %>
                    <span><% If FEATURE_I18N Then %><%= T("products_search_for", Array(HTMLEncode(keyword))) %><% Else %>搜索"<%= HTMLEncode(keyword) %>"<% End If %></span>
                    <%
                    End If
                    %>
                    <span><% If FEATURE_I18N Then %><%= T("products_count", Array(totalCount)) %><% Else %>共 <%= totalCount %> 件商品<% End If %></span>
                </div>
                <div class="sort-options">
                    <label><% If FEATURE_I18N Then %><%= T("products_sort_label", Empty) %><% Else %>排序<% End If %>:</label>
                    <select id="sortSelect" onchange="changeSort(this.value)">
                        <option value="newest" <% If sortBy = "" Or sortBy = "newest" Then Response.Write "selected" End If %>><% If FEATURE_I18N Then %><%= T("products_sort_newest", Empty) %><% Else %>最新上架<% End If %></option>
                        <option value="price_asc" <% If sortBy = "price_asc" Then Response.Write "selected" End If %>><% If FEATURE_I18N Then %><%= T("products_sort_price_asc", Empty) %><% Else %>价格从低到高<% End If %></option>
                        <option value="price_desc" <% If sortBy = "price_desc" Then Response.Write "selected" End If %>><% If FEATURE_I18N Then %><%= T("products_sort_price_desc", Empty) %><% Else %>价格从高到低<% End If %></option>
                        <option value="name" <% If sortBy = "name" Then Response.Write "selected" End If %>><% If FEATURE_I18N Then %><%= T("products_sort_name", Empty) %><% Else %>名称排序<% End If %></option>
                    </select>
                </div>
            </div>

            <div class="products-grid">
                <%
                Dim offset, rsProducts, productCount
                offset = (page - 1) * PAGE_SIZE
                productCount = 0
                
                Dim sql, pId, pType, pPrice, pCategory, minP
                ' Access分页查询 - 使用TOP和子查询
                If offset = 0 Then
                    sql = "SELECT TOP " & PAGE_SIZE & " * FROM Products" & whereClause & orderClause
                Else
                    ' Access分页：获取所有记录，程序中跳过
                    sql = "SELECT * FROM Products" & whereClause & orderClause
                End If
                
                Set rsProducts = ExecuteQuery(sql)
                If Not rsProducts Is Nothing Then
                    ' 跳过前offset条记录
                    Dim skipCount
                    skipCount = 0
                    Do While Not rsProducts.EOF And skipCount < offset
                        rsProducts.MoveNext
                        skipCount = skipCount + 1
                    Loop
                    
                    Do While Not rsProducts.EOF And productCount < PAGE_SIZE
                        productCount = productCount + 1
                        pId = rsProducts("ProductID")
                        pType = rsProducts("ProductType")
                        If IsNull(pType) Then pType = "Custom"
                        pCategory = rsProducts("Category")
                        If IsNull(rsProducts("BasePrice")) Then
                            pPrice = 0
                        Else
                            pPrice = CDbl(rsProducts("BasePrice"))
                        End If
                        
                        If pType = "standard" Then
                            ' standard类型产品：从ProductVolumePrices获取最低价格，如果没有则使用BasePrice
                            Dim rsFixedPrice
                            Set rsFixedPrice = ExecuteQuery("SELECT MIN(Price) AS MinPrice FROM ProductVolumePrices WHERE ProductID = " & pId)
                            If Not rsFixedPrice Is Nothing Then
                                If Not rsFixedPrice.EOF Then
                                    If Not IsNull(rsFixedPrice("MinPrice")) Then
                                        pPrice = CDbl(rsFixedPrice("MinPrice"))
                                    End If
                                End If
                                rsFixedPrice.Close
                            End If
                            Set rsFixedPrice = Nothing
                        End If
                %>
                <div class="product-card">
                    <div class="product-image">
                        <img data-src="<%= rsProducts("ImageURL") %>" src="<%= DEFAULT_PRODUCT_IMAGE %>" alt="<%= HTMLEncode(rsProducts("ProductName")) %>" class="lazy-placeholder" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                        <div class="product-badges">
                            <% If pType = "standard" Then %>
                            <span class="badge badge-fixed"><% If FEATURE_I18N Then %><%= T("products_badge_fixed", Empty) %><% Else %>品牌定香<% End If %></span>
                            <% ElseIf pType = "KOL" Then %>
                            <span class="badge badge-kol"><% If FEATURE_I18N Then %><%= T("products_badge_kol", Empty) %><% Else %>KOL推荐<% End If %></span>
                            <% End If %>
                            <% If pCategory = "奢华系列" Or pCategory = "Luxury" Then %>
                            <span class="badge badge-premium"><% If FEATURE_I18N Then %><%= T("products_badge_premium", Empty) %><% Else %>奢华<% End If %></span>
                            <% End If %>
                        </div>
                        <div class="product-overlay">
                            <a href="/product.asp?id=<%= pId %>" class="btn btn-white"><% If FEATURE_I18N Then %><%= T("view_detail", Empty) %><% Else %>查看详情<% End If %></a>
                            <button class="btn btn-icon" onclick="quickView(<%= pId %>)" title="<% If FEATURE_I18N Then %><%= T("products_quick_view", Empty) %><% Else %>快速预览<% End If %>">
                                <i class="fas fa-eye"></i>
                            </button>
                        </div>
                    </div>
                    <div class="product-info">
                        <span class="product-category"><% If FEATURE_I18N And pCategory = "奢华系列" Then %><%= T("products_badge_premium", Empty) %><% Else %><%= HTMLEncode(pCategory) %><% End If %></span>
                        <h3><a href="/product.asp?id=<%= pId %>"><%= HTMLEncode(rsProducts("ProductName")) %></a></h3>
                        <p class="product-desc"><%= HTMLEncode(Left(rsProducts("Description") & "", 60)) %>...</p>
                        <div class="product-footer">
                            <div class="product-price">
                                <span class="price"><%= FormatMoney(pPrice) %></span>
                                <span class="price-label"><% If FEATURE_I18N Then %><%= T("price_from", Empty) %><% Else %>起<% End If %></span>
                            </div>
                            <a href="/product.asp?id=<%= pId %>" class="btn btn-sm btn-outline">
                                <% If pType = "standard" Then %><% If FEATURE_I18N Then %><%= T("products_btn_purchase", Empty) %><% Else %>选购<% End If %><% Else %><% If FEATURE_I18N Then %><%= T("products_btn_customize", Empty) %><% Else %>定制<% End If %><% End If %>
                            </a>
                        </div>
                    </div>
                </div>
                <%
                        rsProducts.MoveNext
                    Loop
                    rsProducts.Close
                    Set rsProducts = Nothing
                End If
                
                If productCount = 0 Then
                %>
                <div class="no-results">
                    <i class="fas fa-search"></i>
                    <h3><% If FEATURE_I18N Then %><%= T("products_no_results", Empty) %><% Else %>未找到相关产品<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("products_no_results_hint", Empty) %><% Else %>换个关键词试试，或者浏览我们的全部产品<% End If %></p>
                    <a href="/products.asp" class="btn btn-primary"><% If FEATURE_I18N Then %><%= T("products_no_results_btn", Empty) %><% Else %>查看全部<% End If %></a>
                </div>
                <%
                End If
                %>
            </div>

            <!-- 分页 -->
            <%
            If totalPages > 1 Then
            %>
            <div class="pagination">
                <%
                Dim pageUrl, i
                pageUrl = "products.asp?"
                If keyword <> "" Then pageUrl = pageUrl & "keyword=" & Server.URLEncode(keyword) & "&"
                If sortBy <> "" Then pageUrl = pageUrl & "sort=" & sortBy & "&"
                If productType <> "" Then pageUrl = pageUrl & "type=" & Server.URLEncode(productType) & "&"
                
                If page > 1 Then
                %>
                <a href="<%= pageUrl %>page=<%= page - 1 %>" class="page-link"><i class="fas fa-chevron-left"></i></a>
                <%
                End If
                
                For i = 1 To totalPages
                    If i = page Then
                %>
                <span class="page-link active"><%= i %></span>
                <%
                    ElseIf Abs(i - page) <= 2 Or i = 1 Or i = totalPages Then
                %>
                <a href="<%= pageUrl %>page=<%= i %>" class="page-link"><%= i %></a>
                <%
                    ElseIf Abs(i - page) = 3 Then
                %>
                <span class="page-dots">...</span>
                <%
                    End If
                Next
                
                If page < totalPages Then
                %>
                <a href="<%= pageUrl %>page=<%= page + 1 %>" class="page-link"><i class="fas fa-chevron-right"></i></a>
                <%
                End If
                %>
            </div>
            <%
            End If
            %>
        </div>
    </div>
</div>

<script>
function changeSort(value) {
    var url = new URL(window.location.href);
    url.searchParams.set('sort', value);
    url.searchParams.delete('page');
    window.location.href = url.toString();
}

function quickView(productId) {
    // 快速预览功能
    window.location.href = '/product.asp?id=' + productId;
}
</script>

<!--#include file="includes/footer.asp"-->
<%
' ============================================
' V18 AI增强搜索辅助函数
' ============================================

' 同义词扩展：将搜索关键词映射为同义词用于SQL LIKE扩展
' 输入: "清新" → 输出: "海洋"
Function SEARCH_ExpandSearchKeywords(keyword)
    Dim kw
    kw = LCase(Trim(keyword))
    Select Case kw
        Case "清新", "清爽", "干净": SEARCH_ExpandSearchKeywords = SafeLike("海洋")
        Case "浓郁", "持久", "浓香": SEARCH_ExpandSearchKeywords = SafeLike("木质")
        Case "温柔", "柔和", "淡雅": SEARCH_ExpandSearchKeywords = SafeLike("花香")
        Case "阳光", "活力", "运动": SEARCH_ExpandSearchKeywords = SafeLike("柑橘")
        Case "性感", "神秘", "魅惑": SEARCH_ExpandSearchKeywords = SafeLike("麝香")
        Case "成熟", "稳重", "经典": SEARCH_ExpandSearchKeywords = SafeLike("木质")
        Case "甜美", "可爱", "少女": SEARCH_ExpandSearchKeywords = SafeLike("果香")
        Case "中性", "百搭", "日常": SEARCH_ExpandSearchKeywords = SafeLike("柑橘")
        Case Else: SEARCH_ExpandSearchKeywords = ""
    End Select
End Function

' 拼音扩展：将拼音输入映射为中文搜索词
' 输入: "hua" → 输出: "花香"
' 输入: "ganju" → 输出: "柑橘"
Function SEARCH_PinyinExpand(keyword)
    Dim kw
    kw = LCase(Trim(keyword))
    Select Case kw
        Case "hua", "huaxiang":           SEARCH_PinyinExpand = SafeLike("花香")
        Case "ganju", "juzi":             SEARCH_PinyinExpand = SafeLike("柑橘")
        Case "muzhi", "mu":               SEARCH_PinyinExpand = SafeLike("木质")
        Case "dongfang":                  SEARCH_PinyinExpand = SafeLike("东方")
        Case "qingxin", "qing":           SEARCH_PinyinExpand = SafeLike("清新")
        Case "haian", "haiyang", "hai":   SEARCH_PinyinExpand = SafeLike("海洋")
        Case "meigui", "gui":             SEARCH_PinyinExpand = SafeLike("玫瑰")
        Case "moli":                      SEARCH_PinyinExpand = SafeLike("茉莉")
        Case "tanxiang", "tan":           SEARCH_PinyinExpand = SafeLike("檀香")
        Case "xuesong", "song":           SEARCH_PinyinExpand = SafeLike("雪松")
        Case "hupo":                      SEARCH_PinyinExpand = SafeLike("琥珀")
        Case "shexiang", "she":           SEARCH_PinyinExpand = SafeLike("麝香")
        Case "xiangcao":                  SEARCH_PinyinExpand = SafeLike("香草")
        Case "guoxiang", "guo":           SEARCH_PinyinExpand = SafeLike("果香")
        Case "xiangshui":                 SEARCH_PinyinExpand = SafeLike("香水")
        Case "dingzhi":                   SEARCH_PinyinExpand = SafeLike("定制")
        Case "xiangfen":                  SEARCH_PinyinExpand = SafeLike("香氛")
        Case "nishi":                     SEARCH_PinyinExpand = SafeLike("女士")
        Case "nanshi":                    SEARCH_PinyinExpand = SafeLike("男士")
        Case "zhongxing", "zhong":        SEARCH_PinyinExpand = SafeLike("中性")
        Case "lvye", "lvcha", "lv":       SEARCH_PinyinExpand = SafeLike("绿茶")
        Case "xunyi", "xun":              SEARCH_PinyinExpand = SafeLike("薰衣草")
        Case Else:                        SEARCH_PinyinExpand = ""
    End Select
End Function

Call CloseConnection()
%>
