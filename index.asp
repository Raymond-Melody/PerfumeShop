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

' 获取动态栏目配置
Dim activeTypes
activeTypes = GetActiveProductTypes()
%>
<!--#include file="includes/header.asp"-->

<!-- 轮播横幅 -->
<section class="hero-banner">
    <div class="banner-slider">
        <div class="slide active">
            <div class="slide-content">
                <h1>定制你的专属香氛</h1>
                <p>从前调到后调，每一个选择都是你的独特表达</p>
                <a href="/customize.asp" class="btn btn-primary btn-lg">开始定制</a>
            </div>
        </div>
        <div class="slide">
            <div class="slide-content">
                <h1>新品上市 - 皇室典藏</h1>
                <p>顶级原料，匠心之作，尊享奢华体验</p>
                <a href="/product.asp?id=6" class="btn btn-primary btn-lg">立即查看</a>
            </div>
        </div>
        <div class="slide">
            <div class="slide-content">
                <h1>满299免运费</h1>
                <p>首单立减50元，开启香氛之旅</p>
                <a href="/products.asp" class="btn btn-primary btn-lg">立即选购</a>
            </div>
        </div>
    </div>
    <div class="banner-dots">
        <span class="dot active" data-index="0"></span>
        <span class="dot" data-index="1"></span>
        <span class="dot" data-index="2"></span>
    </div>
</section>

<!-- 特色服务 -->
<section class="features-section">
    <div class="container">
        <div class="features-grid">
            <div class="feature-item">
                <i class="fas fa-magic"></i>
                <h3>个性定制</h3>
                <p>自由搭配香调，创造专属香水</p>
            </div>
            <div class="feature-item">
                <i class="fas fa-leaf"></i>
                <h3>天然原料</h3>
                <p>精选全球优质天然香料</p>
            </div>
            <div class="feature-item">
                <i class="fas fa-truck"></i>
                <h3>快速配送</h3>
                <p>下单后48小时内发货</p>
            </div>
            <div class="feature-item">
                <i class="fas fa-undo"></i>
                <h3>无忧退换</h3>
                <p>7天无理由退换货</p>
            </div>
        </div>
    </div>
</section>

<!-- 产品栏目展示 -->
<%
If IsArray(activeTypes) Then
    Dim secIdx, pId, pPrice, pCategory, minP
    Dim secTypeCode, secDisplayName, secNavName, secIcon, secRequiresReview
    Dim secSql, secWhere, rsSection
    For secIdx = 0 To UBound(activeTypes, 1)
        secTypeCode = activeTypes(secIdx, 0)
        secDisplayName = activeTypes(secIdx, 1)
        secNavName = activeTypes(secIdx, 2)
        secIcon = activeTypes(secIdx, 4)
        secRequiresReview = activeTypes(secIdx, 5)
        
        ' V12.0修复：处理Null值
        If IsNull(secIcon) Or secIcon = "" Then secIcon = "fas fa-box"
        If IsNull(secDisplayName) Or secDisplayName = "" Then secDisplayName = secNavName
        
        ' 构建查询SQL
        secWhere = "WHERE IsActive <> 0 AND ProductType='" & SafeSQL(secTypeCode) & "'"
        If secRequiresReview Then
            secWhere = secWhere & " AND ReviewStatus='Approved'"
        End If
        secSql = "SELECT TOP 8 * FROM Products " & secWhere & " ORDER BY CreatedAt DESC"
        
        Set rsSection = ExecuteQuery(secSql)
        
        ' 只在有商品时显示栏目
        If Not rsSection Is Nothing Then
            If Not rsSection.EOF Then
%>
<!-- <%= Server.HTMLEncode(secDisplayName) %>栏目 -->
<section class="products-section">
    <div class="container">
        <div class="section-header">
            <h2><i class="<%= Server.HTMLEncode(secIcon) %>"></i> <%= Server.HTMLEncode(secDisplayName) %></h2>
            <p><%= Server.HTMLEncode(secNavName) %></p>
        </div>
        <div class="products-grid">
            <%
                Do While Not rsSection.EOF
                    pId = rsSection("ProductID")
                    pCategory = rsSection("Category")
                    pPrice = rsSection("BasePrice")
                    
                    ' 品牌定香类型需要获取最低价格
                    If secTypeCode = "standard" Then
                        minP = GetScalar("SELECT MIN(Price) FROM ProductVolumePrices WHERE ProductID = " & pId)
                        If IsNumeric(minP) And minP <> "" Then pPrice = minP
                    End If
            %>
            <div class="product-card">
                <div class="product-image">
                    <img src="<%= rsSection("ImageURL") %>" alt="<%= HTMLEncode(rsSection("ProductName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                    <% If secTypeCode = "standard" Then %>
                    <div class="product-badges">
                        <span class="badge badge-fixed">品牌定香</span>
                    </div>
                    <% ElseIf secTypeCode = "KOL" Then %>
                    <div class="product-badges">
                        <span class="badge badge-kol">KOL推荐</span>
                    </div>
                    <% End If %>
                    <div class="product-overlay">
                        <a href="/product.asp?id=<%= pId %>" class="btn btn-white"><% If secTypeCode = "Custom" Then %>开始定制<% Else %>查看详情<% End If %></a>
                    </div>
                </div>
                <div class="product-info">
                    <span class="product-category"><%= HTMLEncode(pCategory) %></span>
                    <h3><a href="/product.asp?id=<%= pId %>"><%= HTMLEncode(rsSection("ProductName")) %></a></h3>
                    <p class="product-desc"><%= HTMLEncode(Left(rsSection("Description") & "", 50)) %>...</p>
                    <div class="product-price">
                        <span class="price"><%= FormatMoney(pPrice) %></span>
                        <span class="price-label">起</span>
                    </div>
                </div>
            </div>
            <%
                    rsSection.MoveNext
                Loop
            %>
        </div>
        <div class="section-footer">
            <a href="/products.asp?type=<%= Server.URLEncode(secTypeCode) %>" class="btn btn-outline">查看更多<%= Server.HTMLEncode(secDisplayName) %></a>
        </div>
    </div>
</section>
<%
            End If
            rsSection.Close
            Set rsSection = Nothing
        End If
    Next
End If
%>

<!-- 定制流程 -->
<section class="process-section">
    <div class="container">
        <div class="section-header">
            <h2>定制流程</h2>
            <p>简单四步，打造专属于你的香水</p>
        </div>
        <div class="process-steps">
            <div class="step">
                <div class="step-number">1</div>
                <div class="step-icon"><i class="fas fa-flask"></i></div>
                <h3>选择基调</h3>
                <p>从花香、东方、木质等风格中选择</p>
            </div>
            <div class="step-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="step">
                <div class="step-number">2</div>
                <div class="step-icon"><i class="fas fa-palette"></i></div>
                <h3>搭配香调</h3>
                <p>选择前调、中调、后调的香料组合</p>
            </div>
            <div class="step-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="step">
                <div class="step-number">3</div>
                <div class="step-icon"><i class="fas fa-wine-bottle"></i></div>
                <h3>选择瓶型</h3>
                <p>挑选心仪的瓶身款式和容量</p>
            </div>
            <div class="step-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="step">
                <div class="step-number">4</div>
                <div class="step-icon"><i class="fas fa-pen-fancy"></i></div>
                <h3>专属标签</h3>
                <p>添加个性化文字，完成专属定制</p>
            </div>
        </div>
        <div class="section-footer">
            <a href="/customize.asp" class="btn btn-primary btn-lg">立即开始定制</a>
        </div>
    </div>
</section>

<!-- 香调分类 -->
<section class="categories-section">
    <div class="container">
        <div class="section-header">
            <h2>探索香调</h2>
            <p>找到属于你的香氛风格</p>
        </div>
        <div class="categories-grid">
            <a href="/products.asp?category=花香调" class="category-card">
                <div class="category-bg category-bg-floral"></div>
                <div class="category-content">
                    <h3>花香调</h3>
                    <p>浪漫优雅，女性魅力</p>
                </div>
            </a>
            <a href="/products.asp?category=东方调" class="category-card">
                <div class="category-bg category-bg-oriental"></div>
                <div class="category-content">
                    <h3>东方调</h3>
                    <p>神秘深邃，异域风情</p>
                </div>
            </a>
            <a href="/products.asp?category=木质调" class="category-card">
                <div class="category-bg category-bg-woody"></div>
                <div class="category-content">
                    <h3>木质调</h3>
                    <p>沉稳内敛，自然气息</p>
                </div>
            </a>
            <a href="/products.asp?category=海洋调" class="category-card">
                <div class="category-bg category-bg-oceanic"></div>
                <div class="category-content">
                    <h3>海洋调</h3>
                    <p>清新自由，活力四射</p>
                </div>
            </a>
        </div>
    </div>
</section>

<!-- 品牌故事 -->
<section class="story-section">
    <div class="container">
        <div class="story-content">
            <div class="story-text">
                <h2>我们的故事</h2>
                <p>香氛定制诞生于对香水艺术的热爱与追求。我们相信，每个人都值得拥有一款独一无二的香水，它不仅是一种气味，更是个性与情感的表达。</p>
                <p>我们的调香师团队拥有超过20年的行业经验，精选来自世界各地的优质天然香料，为您打造专属的香氛体验。</p>
                <a href="/about.asp" class="btn btn-outline">了解更多</a>
            </div>
            <div class="story-image">
                <div class="story-image-placeholder"></div>
            </div>
        </div>
    </div>
</section>

<!-- 客户评价 -->
<section class="reviews-section">
    <div class="container">
        <div class="section-header">
            <h2>客户评价</h2>
            <p>听听他们怎么说</p>
        </div>
        <div class="reviews-slider">
            <div class="review-card">
                <div class="review-stars">
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                </div>
                <p class="review-text">"定制的香水太惊艳了！调香师的建议非常专业，成品完全符合我想要的感觉，独一无二的气味让我收到很多赞美。"</p>
                <div class="review-author">
                    <span class="author-name">李女士</span>
                    <span class="author-info">购买了花漾年华定制款</span>
                </div>
            </div>
            <div class="review-card">
                <div class="review-stars">
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                </div>
                <p class="review-text">"第一次尝试定制香水，整个过程很有趣！选香调的时候就像在创作艺术品，最后的成品留香时间也很长。"</p>
                <div class="review-author">
                    <span class="author-name">王先生</span>
                    <span class="author-info">购买了东方秘境定制款</span>
                </div>
            </div>
            <div class="review-card">
                <div class="review-stars">
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star"></i>
                    <i class="fas fa-star-half-alt"></i>
                </div>
                <p class="review-text">"送给闺蜜的生日礼物，她超级喜欢！瓶身还可以刻字，非常有心意。下次还会再来定制！"</p>
                <div class="review-author">
                    <span class="author-name">张小姐</span>
                    <span class="author-info">购买了清新森林定制款</span>
                </div>
            </div>
        </div>
    </div>
</section>

<!-- 订阅 -->
<section class="subscribe-section">
    <div class="container">
        <div class="subscribe-content">
            <h2>订阅我们</h2>
            <p>获取最新产品资讯和独家优惠</p>
            <form class="subscribe-form" id="subscribeForm">
                <input type="email" name="email" placeholder="请输入您的邮箱地址" required>
                <button type="submit" class="btn btn-primary">订阅</button>
            </form>
        </div>
    </div>
</section>

<script>
// 轮播图功能
$(document).ready(function() {
    var currentSlide = 0;
    var totalSlides = $('.slide').length;
    
    function showSlide(index) {
        $('.slide').removeClass('active');
        $('.dot').removeClass('active');
        $('.slide').eq(index).addClass('active');
        $('.dot').eq(index).addClass('active');
    }
    
    function nextSlide() {
        currentSlide = (currentSlide + 1) % totalSlides;
        showSlide(currentSlide);
    }
    
    // 自动播放
    setInterval(nextSlide, 5000);
    
    // 点击切换
    $('.dot').click(function() {
        currentSlide = $(this).data('index');
        showSlide(currentSlide);
    });
    
    // 订阅表单
    $('#subscribeForm').submit(function(e) {
        e.preventDefault();
        alert('感谢您的订阅！');
        $(this).find('input[type="email"]').val('');
    });
});
</script>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
