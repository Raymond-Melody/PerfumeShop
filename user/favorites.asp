<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

' 检查用户是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("URL"))
End If

Dim userId
userId = Session("UserID")

' 获取参数
Dim sortBy
sortBy = Request.QueryString("sort")

Select Case sortBy
    Case "price_asc"
        orderClause = " ORDER BY p.BasePrice ASC"
    Case "price_desc"
        orderClause = " ORDER BY p.BasePrice DESC"
    Case "name"
        orderClause = " ORDER BY p.ProductName ASC"
    Case Else
        orderClause = " ORDER BY f.CreatedTime DESC"
End Select
%>
<!--#include file="../includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/user/index.asp">个人中心</a>
        <span class="separator">/</span>
        <span>我的收藏</span>
    </div>
</div>

<div class="container">
    <div class="products-page">
        <!-- 侧边栏筛选 -->
        <aside class="sidebar">
            <div class="filter-section">
                <h3>收藏管理</h3>
                <ul class="filter-list">
                    <li><a href="/user/favorites.asp" class="active">我的收藏</a></li>
                </ul>
            </div>
            
            <div class="filter-section">
                <h3>香调分类</h3>
                <ul class="filter-list">
                    <li><a href="/user/favorites.asp?sort=<%= sortBy %>">全部</a></li>
                    <li><a href="/user/favorites.asp?category=花香调&sort=<%= sortBy %>">花香调</a></li>
                    <li><a href="/user/favorites.asp?category=东方调&sort=<%= sortBy %>">东方调</a></li>
                    <li><a href="/user/favorites.asp?category=木质调&sort=<%= sortBy %>">木质调</a></li>
                    <li><a href="/user/favorites.asp?category=海洋调&sort=<%= sortBy %>">海洋调</a></li>
                    <li><a href="/user/favorites.asp?category=美食调&sort=<%= sortBy %>">美食调</a></li>
                    <li><a href="/user/favorites.asp?category=奢华系列&sort=<%= sortBy %>">奢华系列</a></li>
                </ul>
            </div>
            
            <div class="filter-section cta-box">
                <h3>继续探索</h3>
                <p>发现更多心仪香水</p>
                <a href="/products.asp" class="btn btn-primary btn-block">浏览商品</a>
            </div>
        </aside>

        <!-- 产品列表 -->
        <div class="products-main">
            <div class="products-header">
                <div class="results-info">
                    <span>我的收藏</span>
                </div>
                <div class="sort-options">
                    <label>排序:</label>
                    <select id="sortSelect" onchange="changeSort(this.value)">
                        <option value="newest" <% If sortBy = "" Or sortBy = "newest" Then Response.Write "selected" End If %>>最新收藏</option>
                        <option value="price_asc" <% If sortBy = "price_asc" Then Response.Write "selected" End If %>>价格从低到高</option>
                        <option value="price_desc" <% If sortBy = "price_desc" Then Response.Write "selected" End If %>>价格从高到低</option>
                        <option value="name" <% If sortBy = "name" Then Response.Write "selected" End If %>>名称排序</option>
                    </select>
                </div>
            </div>
            
            <%
            ' 构建查询SQL
            Dim whereClause, categoryFilter
            categoryFilter = Request.QueryString("category")
            whereClause = " WHERE f.UserID = " & userId & " AND p.IsActive <> 0"
            If categoryFilter <> "" Then
                whereClause = whereClause & " AND p.Category = '" & SafeSQL(categoryFilter) & "'"
            End If
            
            ' 获取用户收藏的商品列表
            Dim rsFavorites
            Set rsFavorites = ExecuteQuery("SELECT f.*, p.ProductName, p.Description, p.BasePrice, p.ImageURL, p.Category FROM UserFavorites f INNER JOIN Products p ON f.ProductID = p.ProductID" & whereClause & orderClause)
            If Not rsFavorites Is Nothing Then
                If Not rsFavorites.EOF Then
            %>
            <div class="products-grid">
                <%
                Do While Not rsFavorites.EOF
                %>
                <div class="product-card">
                    <div class="product-image">
                        <img src="<%= rsFavorites("ImageURL") %>" alt="<%= HTMLEncode(rsFavorites("ProductName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                        <div class="product-badges">
                            <% If rsFavorites("Category") = "奢华系列" Then %>
                            <span class="badge badge-premium">奢华</span>
                            <% End If %>
                        </div>
                        <div class="product-overlay">
                            <a href="/product.asp?id=<%= rsFavorites("ProductID") %>" class="btn btn-white">查看详情</a>
                            <button class="btn btn-icon" title="取消收藏" data-product-id="<%= rsFavorites("ProductID") %>" onclick="removeFromFavorites(this.getAttribute('data-product-id'))">
                                <i class="fas fa-heart"></i>
                            </button>
                        </div>
                    </div>
                    <div class="product-info">
                        <span class="product-category"><%= HTMLEncode(rsFavorites("Category")) %></span>
                        <h3><a href="/product.asp?id=<%= rsFavorites("ProductID") %>"><%= HTMLEncode(rsFavorites("ProductName")) %></a></h3>
                        <p class="product-desc"><%= HTMLEncode(Left(rsFavorites("Description"), 60)) %>...</p>
                        <div class="product-footer">
                            <div class="product-price">
                                <span class="price"><%= FormatMoney(rsFavorites("BasePrice")) %></span>
                                <span class="price-label">起</span>
                            </div>
                            <a href="/product.asp?id=<%= rsFavorites("ProductID") %>" class="btn btn-sm btn-outline">定制</a>
                        </div>
                    </div>
                </div>
                <%
                rsFavorites.MoveNext
                Loop
                rsFavorites.Close
                Set rsFavorites = Nothing
                %>
            </div>
            <%
                Else
            %>
            <div class="no-results">
                <i class="fas fa-heart"></i>
                <h3>暂无收藏商品</h3>
                <p>收藏喜欢的商品，方便以后快速找到</p>
                <a href="/products.asp" class="btn btn-primary">去逛逛</a>
            </div>
            <%
                End If
            End If
            %>
        </div>
    </div>
</div>

<script>
function removeFromFavorites(productId) {
    // 确保productId是数字
    productId = parseInt(productId);
    if (isNaN(productId)) {
        alert('无效的商品ID');
        return;
    }
    
    if (confirm('确定要取消收藏这个商品吗？')) {
        // 发送AJAX请求取消收藏
        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/api/favorites.asp', true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var response = JSON.parse(xhr.responseText);
                        if (response.success) {
                            alert(response.message);
                            location.reload();
                        } else {
                            alert(response.message || '操作失败');
                        }
                    } catch (e) {
                        alert('操作失败');
                    }
                } else {
                    alert('网络错误');
                }
            }
        };
        
        xhr.send('action=remove&productId=' + encodeURIComponent(productId));
    }
}

function changeSort(value) {
    var url = new URL(window.location.href);
    url.searchParams.set('sort', value);
    url.searchParams.delete('page'); // 删除页码参数
    window.location.href = url.toString();
}

function quickView(productId) {
    // 快速预览功能
    window.location.href = '/product.asp?id=' + productId;
}
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>