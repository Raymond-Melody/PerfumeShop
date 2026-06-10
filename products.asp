<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<%
Call OpenConnection()

' 获取所有已启用的商品类型
Dim activeTypes, activeTypeCodes
activeTypes = GetActiveProductTypes()
activeTypeCodes = GetActiveTypeCodesForSQL()

' 获取参数
Dim keyword, sortBy, page, totalCount, totalPages, productType
keyword = SafeSQL(Request.QueryString("keyword"))
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
    whereClause = whereClause & " AND (ProductName LIKE '%" & keyword & "%' OR Description LIKE '%" & keyword & "%')"
End If

Select Case sortBy
    Case "price_asc"
        orderClause = " ORDER BY BasePrice ASC"
    Case "price_desc"
        orderClause = " ORDER BY BasePrice DESC"
    Case "name"
        orderClause = " ORDER BY ProductName ASC"
    Case Else
        orderClause = " ORDER BY CreatedAt DESC"
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
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <%
        If keyword <> "" Then
        %>
        <span>搜索结果: <%= HTMLEncode(keyword) %></span>
        <%
        Else
        %>
        <span>全部香水</span>
        <%
        End If
        %>
    </div>
</div>

<div class="container">
    <div class="products-page">
        <!-- 侧边栏筛选 -->
        <aside class="sidebar">
            <div class="filter-section">
                <h4>产品类型</h4>
                <ul class="filter-list">
                    <li><a href="/products.asp" class="<%= IIf(productType = "", "active", "") %>">全部</a></li>
                    <% 
                    Dim fIdx, fCode, fName
                    If IsArray(activeTypes) Then
                        For fIdx = 0 To UBound(activeTypes, 1)
                            fCode = activeTypes(fIdx, 0)
                            fName = activeTypes(fIdx, 1)
                    %>
                    <li><a href="/products.asp?type=<%= Server.URLEncode(fCode) %>" class="<%= IIf(productType = fCode, "active", "") %>"><%= Server.HTMLEncode(fName) %></a></li>
                    <%
                        Next
                    End If
                    %>
                </ul>
            </div>
            
            <div class="filter-section">
                <h3>价格区间</h3>
                <ul class="filter-list">
                    <li><a href="/products.asp?price=0-300">¥300以下</a></li>
                    <li><a href="/products.asp?price=300-500">¥300-500</a></li>
                    <li><a href="/products.asp?price=500-800">¥500-800</a></li>
                    <li><a href="/products.asp?price=800">¥800以上</a></li>
                </ul>
            </div>
            
            <div class="filter-section cta-box">
                <h3>想要专属定制？</h3>
                <p>打造独一无二的个性香水</p>
                <a href="/customize.asp" class="btn btn-primary btn-block">开始定制</a>
            </div>
        </aside>

        <!-- 产品列表 -->
        <div class="products-main">
            <div class="products-header">
                <div class="results-info">
                    <%
                    If keyword <> "" Then
                    %>
                    <span>搜索"<%= HTMLEncode(keyword) %>"</span>
                    <%
                    End If
                    %>
                    <span>共 <%= totalCount %> 件商品</span>
                </div>
                <div class="sort-options">
                    <label>排序:</label>
                    <select id="sortSelect" onchange="changeSort(this.value)">
                        <option value="newest" <% If sortBy = "" Or sortBy = "newest" Then Response.Write "selected" End If %>>最新上架</option>
                        <option value="price_asc" <% If sortBy = "price_asc" Then Response.Write "selected" End If %>>价格从低到高</option>
                        <option value="price_desc" <% If sortBy = "price_desc" Then Response.Write "selected" End If %>>价格从高到低</option>
                        <option value="name" <% If sortBy = "name" Then Response.Write "selected" End If %>>名称排序</option>
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
                        
                        If pType = "Fixed" Then
                            ' Fixed类型产品：从ProductVolumePrices获取最低价格，如果没有则使用BasePrice
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
                        <img src="<%= rsProducts("ImageURL") %>" alt="<%= HTMLEncode(rsProducts("ProductName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                        <div class="product-badges">
                            <% If pType = "Fixed" Then %>
                            <span class="badge badge-fixed">品牌定香</span>
                            <% ElseIf pType = "KOL" Then %>
                            <span class="badge badge-kol">KOL推荐</span>
                            <% End If %>
                            <% If pCategory = "奢华系列" Then %>
                            <span class="badge badge-premium">奢华</span>
                            <% End If %>
                        </div>
                        <div class="product-overlay">
                            <a href="/product.asp?id=<%= pId %>" class="btn btn-white">查看详情</a>
                            <button class="btn btn-icon" onclick="quickView(<%= pId %>)" title="快速预览">
                                <i class="fas fa-eye"></i>
                            </button>
                        </div>
                    </div>
                    <div class="product-info">
                        <span class="product-category"><%= HTMLEncode(pCategory) %></span>
                        <h3><a href="/product.asp?id=<%= pId %>"><%= HTMLEncode(rsProducts("ProductName")) %></a></h3>
                        <p class="product-desc"><%= HTMLEncode(Left(rsProducts("Description") & "", 60)) %>...</p>
                        <div class="product-footer">
                            <div class="product-price">
                                <span class="price"><%= FormatMoney(pPrice) %></span>
                                <span class="price-label">起</span>
                            </div>
                            <a href="/product.asp?id=<%= pId %>" class="btn btn-sm btn-outline">
                                <% If pType = "Fixed" Then %>选购<% Else %>定制<% End If %>
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
                    <h3>未找到相关产品</h3>
                    <p>换个关键词试试，或者浏览我们的全部产品</p>
                    <a href="/products.asp" class="btn btn-primary">查看全部</a>
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
Call CloseConnection()
%>
