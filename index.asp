<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
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
                <h1><% If FEATURE_I18N Then %><%= T("home_hero_title_1", Empty) %><% Else %>定制你的专属香氛<% End If %></h1>
                <p><% If FEATURE_I18N Then %><%= T("home_hero_desc_1", Empty) %><% Else %>从前调到后调，每一个选择都是你的独特表达<% End If %></p>
                <a href="/customize.asp" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then %><%= T("home_hero_btn_1", Empty) %><% Else %>开始定制<% End If %></a>
            </div>
        </div>
        <div class="slide">
            <div class="slide-content">
                <h1><% If FEATURE_I18N Then %><%= T("home_hero_title_2", Empty) %><% Else %>新品上市 - 皇室典藏<% End If %></h1>
                <p><% If FEATURE_I18N Then %><%= T("home_hero_desc_2", Empty) %><% Else %>顶级原料，匠心之作，尊享奢华体验<% End If %></p>
                <a href="/product.asp?id=6" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then %><%= T("home_hero_btn_2", Empty) %><% Else %>立即查看<% End If %></a>
            </div>
        </div>
        <div class="slide">
            <div class="slide-content">
                <h1><% If FEATURE_I18N Then %><%= T("home_hero_title_3", Empty) %><% Else %>满299免运费<% End If %></h1>
                <p><% If FEATURE_I18N Then %><%= T("home_hero_desc_3", Empty) %><% Else %>首单立减50元，开启香氛之旅<% End If %></p>
                <a href="/products.asp" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then %><%= T("home_hero_btn_3", Empty) %><% Else %>立即选购<% End If %></a>
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

<% ' V18: 限时活动推荐区块 %>
<% If FEATURE_FLASH_SALE Or FEATURE_GROUP_BUY Or FEATURE_SUBSCRIPTION Or FEATURE_COMMUNITY Then %>
<% Dim promoCount : promoCount = 0 %>
<section class="promo-section">
    <div class="container">
        <div class="promo-grid">
            <% If FEATURE_FLASH_SALE Then %>
            <a href="/flash_sale.asp" class="promo-card promo-flash">
                <div class="promo-icon"><i class="fas fa-bolt"></i></div>
                <div class="promo-content">
                    <h3><% If FEATURE_I18N Then %><%= T("home_promo_flash_title", Empty) %><% Else %>限时秒杀<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("home_promo_flash_desc", Empty) %><% Else %>超值好物限时抢购，最低至3折<% End If %></p>
                    <span class="promo-link"><% If FEATURE_I18N Then %><%= T("home_promo_go", Empty) %><% Else %>立即抢购<% End If %> <i class="fas fa-arrow-right"></i></span>
                </div>
            </a>
            <% End If %>
            <% If FEATURE_GROUP_BUY Then %>
            <a href="/group_buy.asp" class="promo-card promo-group">
                <div class="promo-icon"><i class="fas fa-users"></i></div>
                <div class="promo-content">
                    <h3><% If FEATURE_I18N Then %><%= T("home_promo_group_title", Empty) %><% Else %>拼团惠购<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("home_promo_group_desc", Empty) %><% Else %>邀请好友一起拼，享超低团购价<% End If %></p>
                    <span class="promo-link"><% If FEATURE_I18N Then %><%= T("home_promo_go", Empty) %><% Else %>去拼团<% End If %> <i class="fas fa-arrow-right"></i></span>
                </div>
            </a>
            <% End If %>
            <% If FEATURE_SUBSCRIPTION Then %>
            <a href="/subscribe.asp" class="promo-card promo-subscribe">
                <div class="promo-icon"><i class="fas fa-box-open"></i></div>
                <div class="promo-content">
                    <h3><% If FEATURE_I18N Then %><%= T("home_promo_sub_title", Empty) %><% Else %>香氛订阅盒<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("home_promo_sub_desc", Empty) %><% Else %>每月新香到家，AI个性化选品<% End If %></p>
                    <span class="promo-link"><% If FEATURE_I18N Then %><%= T("home_promo_go", Empty) %><% Else %>了解详情<% End If %> <i class="fas fa-arrow-right"></i></span>
                </div>
            </a>
            <% End If %>
            <% If FEATURE_COMMUNITY Then %>
            <a href="/community.asp" class="promo-card promo-community">
                <div class="promo-icon"><i class="fas fa-comments"></i></div>
                <div class="promo-content">
                    <h3><% If FEATURE_I18N Then %>香氛社区<% Else %>香氛社区<% End If %></h3>
                    <p><% If FEATURE_I18N Then %>分享配方、交流心得、发现同好<% Else %>分享配方、交流心得、发现同好<% End If %></p>
                    <span class="promo-link"><% If FEATURE_I18N Then %><%= T("home_promo_go", Empty) %><% Else %>加入社区<% End If %> <i class="fas fa-arrow-right"></i></span>
                </div>
            </a>
            <% End If %>
        </div>
    </div>
</section>
<% End If %>

<section class="features-section">
    <div class="container">
        <div class="features-grid">
            <div class="feature-item">
                <i class="fas fa-magic"></i>
                <h3><% If FEATURE_I18N Then %><%= T("home_feature_custom", Empty) %><% Else %>个性定制<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_feature_custom_desc", Empty) %><% Else %>自由搭配香调，创造专属香水<% End If %></p>
            </div>
            <div class="feature-item">
                <i class="fas fa-leaf"></i>
                <h3><% If FEATURE_I18N Then %><%= T("home_feature_natural", Empty) %><% Else %>天然原料<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_feature_natural_desc", Empty) %><% Else %>精选全球优质天然香料<% End If %></p>
            </div>
            <div class="feature-item">
                <i class="fas fa-truck"></i>
                <h3><% If FEATURE_I18N Then %><%= T("home_feature_shipping", Empty) %><% Else %>快速配送<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_feature_shipping_desc", Empty) %><% Else %>下单后48小时内发货<% End If %></p>
            </div>
            <div class="feature-item">
                <i class="fas fa-undo"></i>
                <h3><% If FEATURE_I18N Then %><%= T("home_feature_return", Empty) %><% Else %>无忧退换<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_feature_return_desc", Empty) %><% Else %>7天无理由退换货<% End If %></p>
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
            <h2><i class="<%= Server.HTMLEncode(secIcon) %>"></i> <%= Server.HTMLEncode(GetProductTypeI18nName(secTypeCode, secDisplayName, "display")) %></h2>
            <p><%= Server.HTMLEncode(GetProductTypeI18nName(secTypeCode, secNavName, "nav")) %></p>
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
                        <span class="badge badge-fixed"><% If FEATURE_I18N Then %><%= T("home_badge_fixed", Empty) %><% Else %>品牌定香<% End If %></span>
                    </div>
                    <% ElseIf secTypeCode = "KOL" Then %>
                    <div class="product-badges">
                        <span class="badge badge-kol"><% If FEATURE_I18N Then %><%= T("home_badge_kol", Empty) %><% Else %>KOL推荐<% End If %></span>
                    </div>
                    <% End If %>
                    <div class="product-overlay">
                        <a href="/product.asp?id=<%= pId %>" class="btn btn-white"><% If secTypeCode = "Custom" Then %><% If FEATURE_I18N Then %><%= T("start_customize", Empty) %><% Else %>开始定制<% End If %><% Else %><% If FEATURE_I18N Then %><%= T("view_detail", Empty) %><% Else %>查看详情<% End If %><% End If %></a>
                    </div>
                </div>
                <div class="product-info">
                    <span class="product-category"><%= HTMLEncode(pCategory) %></span>
                    <h3><a href="/product.asp?id=<%= pId %>"><%= HTMLEncode(rsSection("ProductName")) %></a></h3>
                    <p class="product-desc"><%= HTMLEncode(Left(rsSection("Description") & "", 50)) %>...</p>
                    <div class="product-price">
                        <span class="price"><%= FormatMoney(pPrice) %></span>
                        <span class="price-label"><% If FEATURE_I18N Then %><%= T("price_from", Empty) %><% Else %>起<% End If %></span>
                    </div>
                </div>
            </div>
            <%
                    rsSection.MoveNext
                Loop
            %>
        </div>
        <div class="section-footer">
            <a href="/products.asp?type=<%= Server.URLEncode(secTypeCode) %>" class="btn btn-outline"><% If FEATURE_I18N Then %><%= T("home_section_more", Array(GetProductTypeI18nName(secTypeCode, secDisplayName, "display"))) %><% Else %>查看更多<%= Server.HTMLEncode(secDisplayName) %><% End If %></a>
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
            <h2><% If FEATURE_I18N Then %><%= T("home_process_title", Empty) %><% Else %>定制流程<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("home_process_desc", Empty) %><% Else %>简单四步，打造专属于你的香水<% End If %></p>
        </div>
        <div class="process-steps">
            <div class="step">
                <div class="step-number">1</div>
                <div class="step-icon"><i class="fas fa-flask"></i></div>
                <h3><% If FEATURE_I18N Then %><%= T("home_process_step1_title", Empty) %><% Else %>选择基调<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_process_step1_desc", Empty) %><% Else %>从花香、东方、木质等风格中选择<% End If %></p>
            </div>
            <div class="step-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="step">
                <div class="step-number">2</div>
                <div class="step-icon"><i class="fas fa-palette"></i></div>
                <h3><% If FEATURE_I18N Then %><%= T("home_process_step2_title", Empty) %><% Else %>搭配香调<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_process_step2_desc", Empty) %><% Else %>选择前调、中调、后调的香料组合<% End If %></p>
            </div>
            <div class="step-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="step">
                <div class="step-number">3</div>
                <div class="step-icon"><i class="fas fa-wine-bottle"></i></div>
                <h3><% If FEATURE_I18N Then %><%= T("home_process_step3_title", Empty) %><% Else %>选择瓶型<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_process_step3_desc", Empty) %><% Else %>挑选心仪的瓶身款式和容量<% End If %></p>
            </div>
            <div class="step-arrow"><i class="fas fa-arrow-right"></i></div>
            <div class="step">
                <div class="step-number">4</div>
                <div class="step-icon"><i class="fas fa-pen-fancy"></i></div>
                <h3><% If FEATURE_I18N Then %><%= T("home_process_step4_title", Empty) %><% Else %>专属标签<% End If %></h3>
                <p><% If FEATURE_I18N Then %><%= T("home_process_step4_desc", Empty) %><% Else %>添加个性化文字，完成专属定制<% End If %></p>
            </div>
        </div>
        <div class="section-footer">
            <a href="/customize.asp" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then %><%= T("home_process_btn", Empty) %><% Else %>立即开始定制<% End If %></a>
        </div>
    </div>
</section>

<!-- 香调分类 -->
<section class="categories-section">
    <div class="container">
        <div class="section-header">
            <h2><% If FEATURE_I18N Then %><%= T("home_categories_title", Empty) %><% Else %>探索香调<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("home_categories_desc", Empty) %><% Else %>找到属于你的香氛风格<% End If %></p>
        </div>
        <div class="categories-grid">
            <a href="/products.asp?category=花香调" class="category-card">
                <div class="category-bg category-bg-floral"></div>
                <div class="category-content">
                    <h3><% If FEATURE_I18N Then %><%= T("home_cat_floral", Empty) %><% Else %>花香调<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("home_cat_floral_desc", Empty) %><% Else %>浪漫优雅，女性魅力<% End If %></p>
                </div>
            </a>
            <a href="/products.asp?category=东方调" class="category-card">
                <div class="category-bg category-bg-oriental"></div>
                <div class="category-content">
                    <h3><% If FEATURE_I18N Then %><%= T("home_cat_oriental", Empty) %><% Else %>东方调<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("home_cat_oriental_desc", Empty) %><% Else %>神秘深邃，异域风情<% End If %></p>
                </div>
            </a>
            <a href="/products.asp?category=木质调" class="category-card">
                <div class="category-bg category-bg-woody"></div>
                <div class="category-content">
                    <h3><% If FEATURE_I18N Then %><%= T("home_cat_woody", Empty) %><% Else %>木质调<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("home_cat_woody_desc", Empty) %><% Else %>沉稳内敛，自然气息<% End If %></p>
                </div>
            </a>
            <a href="/products.asp?category=海洋调" class="category-card">
                <div class="category-bg category-bg-oceanic"></div>
                <div class="category-content">
                    <h3><% If FEATURE_I18N Then %><%= T("home_cat_oceanic", Empty) %><% Else %>海洋调<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("home_cat_oceanic_desc", Empty) %><% Else %>清新自由，活力四射<% End If %></p>
                </div>
            </a>
        </div>
    </div>
</section>

<!-- V18 AI 为你推荐 -->
<%
Dim idxUserId, idxRsRecommend
idxUserId = Session("UserID")
If idxUserId <> "" And Not IsNull(idxUserId) And IsNumeric(idxUserId) Then
    Set idxRsRecommend = RE_GetPersonalizedProducts(CLng(idxUserId), 8)
Else
    Set idxRsRecommend = RE_GetTrendingNow(8)
End If
%>
<section class="recommend-section">
    <div class="container">
        <div class="section-header">
            <h2><i class="fas fa-magic"></i> <% If FEATURE_I18N Then %><%= T("home_recommend_title", Empty) %><% Else %>为你推荐<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("home_recommend_desc", Empty) %><% Else %>基于AI智能分析，为您精选好物<% End If %></p>
        </div>
        <% Call RE_RenderRecommendationsSafe(idxRsRecommend, "ai-recommend-grid", True, "正在为您准备个性化推荐...") %>
        <%
        If Not idxRsRecommend Is Nothing Then
            idxRsRecommend.Close
            Set idxRsRecommend = Nothing
        End If
        %>
    </div>
</section>

<!-- 品牌故事 -->
<section class="story-section">
    <div class="container">
        <div class="story-content">
            <div class="story-text">
                <h2><% If FEATURE_I18N Then %><%= T("home_story_title", Empty) %><% Else %>我们的故事<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("home_story_text_1", Empty) %><% Else %>香氛定制诞生于对香水艺术的热爱与追求。我们相信，每个人都值得拥有一款独一无二的香水，它不仅是一种气味，更是个性与情感的表达。<% End If %></p>
                <p><% If FEATURE_I18N Then %><%= T("home_story_text_2", Empty) %><% Else %>我们的调香师团队拥有超过20年的行业经验，精选来自世界各地的优质天然香料，为您打造专属的香氛体验。<% End If %></p>
                <a href="/about.asp" class="btn btn-outline"><% If FEATURE_I18N Then %><%= T("home_story_btn", Empty) %><% Else %>了解更多<% End If %></a>
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
            <h2><% If FEATURE_I18N Then %><%= T("home_reviews_title", Empty) %><% Else %>客户评价<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("home_reviews_desc", Empty) %><% Else %>听听他们怎么说<% End If %></p>
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
                <p class="review-text">"<% If FEATURE_I18N Then %><%= T("home_review_1_text", Empty) %><% Else %>定制的香水太惊艳了！调香师的建议非常专业，成品完全符合我想要的感觉，独一无二的气味让我收到很多赞美。<% End If %>"</p>
                <div class="review-author">
                    <span class="author-name"><% If FEATURE_I18N Then %><%= T("home_review_1_author", Empty) %><% Else %>李女士<% End If %></span>
                    <span class="author-info"><% If FEATURE_I18N Then %><%= T("home_review_1_info", Empty) %><% Else %>购买了花漾年华定制款<% End If %></span>
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
                <p class="review-text">"<% If FEATURE_I18N Then %><%= T("home_review_2_text", Empty) %><% Else %>第一次尝试定制香水，整个过程很有趣！选香调的时候就像在创作艺术品，最后的成品留香时间也很长。<% End If %>"</p>
                <div class="review-author">
                    <span class="author-name"><% If FEATURE_I18N Then %><%= T("home_review_2_author", Empty) %><% Else %>王先生<% End If %></span>
                    <span class="author-info"><% If FEATURE_I18N Then %><%= T("home_review_2_info", Empty) %><% Else %>购买了东方秘境定制款<% End If %></span>
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
                <p class="review-text">"<% If FEATURE_I18N Then %><%= T("home_review_3_text", Empty) %><% Else %>送给闺蜜的生日礼物，她超级喜欢！瓶身还可以刻字，非常有心意。下次还会再来定制！<% End If %>"</p>
                <div class="review-author">
                    <span class="author-name"><% If FEATURE_I18N Then %><%= T("home_review_3_author", Empty) %><% Else %>张小姐<% End If %></span>
                    <span class="author-info"><% If FEATURE_I18N Then %><%= T("home_review_3_info", Empty) %><% Else %>购买了清新森林定制款<% End If %></span>
                </div>
            </div>
        </div>
    </div>
</section>

<!-- 订阅 -->
<section class="subscribe-section">
    <div class="container">
        <div class="subscribe-content">
            <h2><% If FEATURE_I18N Then %><%= T("home_subscribe_title", Empty) %><% Else %>订阅我们<% End If %></h2>
            <p><% If FEATURE_I18N Then %><%= T("home_subscribe_desc", Empty) %><% Else %>获取最新产品资讯和独家优惠<% End If %></p>
            <form class="subscribe-form" id="subscribeForm">
                <input type="email" name="email" placeholder="<% If FEATURE_I18N Then %><%= T("home_subscribe_placeholder", Empty) %><% Else %>请输入您的邮箱地址<% End If %>" required>
                <button type="submit" class="btn btn-primary"><% If FEATURE_I18N Then %><%= T("home_subscribe_btn", Empty) %><% Else %>订阅<% End If %></button>
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
        alert('<% If FEATURE_I18N Then %><%= T("home_subscribe_thanks", Empty) %><% Else %>感谢您的订阅！<% End If %>');
        $(this).find('input[type="email"]').val('');
    });
});
</script>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
