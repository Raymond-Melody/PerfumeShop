<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()

' V14: 会员登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If
%>
<!--#include file="includes/header.asp"-->

<!-- 定制专区横幅 -->
<section class="customize-hero">
    <div class="container">
        <div class="hero-content">
            <h1><% If FEATURE_I18N Then Response.Write T("customize_hero_title", Empty) Else %>定制你的专属香水<% End If %></h1>
            <p><% If FEATURE_I18N Then Response.Write T("customize_hero_desc", Empty) Else %>从数十种香调中自由搭配，创造独一无二的个人香氛<% End If %></p>
        </div>
    </div>
</section>

<div class="container">
    <div class="customize-page">
        <!-- 定制流程说明 -->
        <section class="customize-intro">
            <div class="intro-grid">
                <div class="intro-item">
                    <div class="intro-icon"><i class="fas fa-1"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("customize_step1_title", Empty) Else %>选择基础香型<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("customize_step1_desc", Empty) Else %>从花香、东方、木质、海洋等经典香型中选择您喜欢的基调<% End If %></p>
                </div>
                <div class="intro-item">
                    <div class="intro-icon"><i class="fas fa-2"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("customize_step2_title", Empty) Else %>搭配香调层次<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("customize_step2_desc", Empty) Else %>选择前调、中调、后调，构建您独特的香氛金字塔<% End If %></p>
                </div>
                <div class="intro-item">
                    <div class="intro-icon"><i class="fas fa-3"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("customize_step3_title", Empty) Else %>选择容量瓶型<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("customize_step3_desc", Empty) Else %>从多种容量规格和精美瓶身中选择，完美呈现您的香水<% End If %></p>
                </div>
                <div class="intro-item">
                    <div class="intro-icon"><i class="fas fa-4"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("customize_step4_title", Empty) Else %>添加专属标签<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("customize_step4_desc", Empty) Else %>刻印专属文字，让这瓶香水成为真正的独家定制<% End If %></p>
                </div>
            </div>
        </section>

        <!-- 选择基础香型 -->
        <section class="customize-section">
            <div class="section-header">
                <h2><i class="fas fa-flask"></i> <% If FEATURE_I18N Then Response.Write T("customize_base_title", Empty) Else %>第一步：选择基础香型<% End If %></h2>
                <p><% If FEATURE_I18N Then Response.Write T("customize_base_desc", Empty) Else %>不同的基础香型带来不同的整体风格<% End If %></p>
            </div>
            <div class="base-products">
                <%
                Dim rsProducts
                Set rsProducts = ExecuteQuery("SELECT * FROM Products WHERE IsActive <> 0 AND ProductType IN ('Custom', 'KOL') ORDER BY BasePrice")
                If Not rsProducts Is Nothing Then
                    Do While Not rsProducts.EOF
                %>
                <div class="base-product-card">
                    <div class="card-image">
                        <img src="<%= rsProducts("ImageURL") %>" alt="<%= HTMLEncode(rsProducts("ProductName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                    </div>
                    <div class="card-content">
                        <span class="card-category"><%= HTMLEncode(rsProducts("Category")) %></span>
                        <h3><%= HTMLEncode(rsProducts("ProductName")) %></h3>
                        <p><%= HTMLEncode(rsProducts("Description")) %></p>
                        <div class="card-footer">
                            <span class="card-price"><%= FormatMoney(rsProducts("BasePrice")) %> <% If FEATURE_I18N Then Response.Write T("price_from", Empty) Else %>起<% End If %></span>
                            <a href="/product.asp?id=<%= rsProducts("ProductID") %>" class="btn btn-primary"><% If FEATURE_I18N Then Response.Write T("start_customize", Empty) Else %>开始定制<% End If %></a>
                        </div>
                    </div>
                </div>
                <%
                        rsProducts.MoveNext
                    Loop
                    rsProducts.Close
                    Set rsProducts = Nothing
                End If
                %>
            </div>
        </section>

        <!-- 香调说明 -->
        <section class="customize-section notes-section">
            <div class="section-header">
                <h2><i class="fas fa-layer-group"></i> <% If FEATURE_I18N Then Response.Write T("customize_notes_title", Empty) Else %>香调层次说明<% End If %></h2>
                <p><% If FEATURE_I18N Then Response.Write T("customize_notes_desc", Empty) Else %>了解香水的三层结构，帮助您做出更好的选择<% End If %></p>
            </div>
            <div class="notes-explanation">
                <div class="note-level top-note">
                    <div class="note-icon"><i class="fas fa-wind"></i></div>
                    <div class="note-content">
                        <h3><% If FEATURE_I18N Then Response.Write T("customize_notes_top_title", Empty) Else %>前调 (Top Notes)<% End If %></h3>
                        <p><% If FEATURE_I18N Then Response.Write T("customize_notes_top_full", Empty) Else %>香水的第一印象，喷洒后立即感受到的香气。通常由轻盈、清新的香调组成，如柑橘、薰衣草、薄荷等。持续时间约15-30分钟。<% End If %></p>
                        <div class="available-notes">
                            <%
                            Dim rsTopNotes
                            Set rsTopNotes = ExecuteQuery("SELECT NoteName FROM FragranceNotes WHERE NoteType = '前调' AND IsActive <> 0")
                            If Not rsTopNotes Is Nothing Then
                                Do While Not rsTopNotes.EOF
                            %>
                            <span class="note-tag"><%= HTMLEncode(rsTopNotes("NoteName")) %></span>
                            <%
                                    rsTopNotes.MoveNext
                                Loop
                                rsTopNotes.Close
                                Set rsTopNotes = Nothing
                            End If
                            %>
                        </div>
                    </div>
                </div>
                <div class="note-level middle-note">
                    <div class="note-icon"><i class="fas fa-heart"></i></div>
                    <div class="note-content">
                        <h3><% If FEATURE_I18N Then Response.Write T("customize_notes_mid_title", Empty) Else %>中调 (Heart Notes)<% End If %></h3>
                        <p><% If FEATURE_I18N Then Response.Write T("customize_notes_mid_full", Empty) Else %>香水的核心灵魂，在前调消散后显现。通常由花香、果香等构成，是整体香氛的主体。持续时间约2-4小时。<% End If %></p>
                        <div class="available-notes">
                            <%
                            Dim rsMiddleNotes
                            Set rsMiddleNotes = ExecuteQuery("SELECT NoteName FROM FragranceNotes WHERE NoteType = '中调' AND IsActive <> 0")
                            If Not rsMiddleNotes Is Nothing Then
                                Do While Not rsMiddleNotes.EOF
                            %>
                            <span class="note-tag"><%= HTMLEncode(rsMiddleNotes("NoteName")) %></span>
                            <%
                                    rsMiddleNotes.MoveNext
                                Loop
                                rsMiddleNotes.Close
                                Set rsMiddleNotes = Nothing
                            End If
                            %>
                        </div>
                    </div>
                </div>
                <div class="note-level base-note">
                    <div class="note-icon"><i class="fas fa-moon"></i></div>
                    <div class="note-content">
                        <h3><% If FEATURE_I18N Then Response.Write T("customize_notes_base_title", Empty) Else %>后调 (Base Notes)<% End If %></h3>
                        <p><% If FEATURE_I18N Then Response.Write T("customize_notes_base_full", Empty) Else %>香水最持久的部分，为整体香氛提供深度和持久力。通常由木质香、麝香、琥珀等组成。持续时间可达数小时甚至一整天。<% End If %></p>
                        <div class="available-notes">
                            <%
                            Dim rsBaseNotes
                            Set rsBaseNotes = ExecuteQuery("SELECT NoteName FROM FragranceNotes WHERE NoteType = '后调' AND IsActive <> 0")
                            If Not rsBaseNotes Is Nothing Then
                                Do While Not rsBaseNotes.EOF
                            %>
                            <span class="note-tag"><%= HTMLEncode(rsBaseNotes("NoteName")) %></span>
                            <%
                                    rsBaseNotes.MoveNext
                                Loop
                                rsBaseNotes.Close
                                Set rsBaseNotes = Nothing
                            End If
                            %>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- 瓶身展示 -->
        <section class="customize-section bottles-section">
            <div class="section-header">
                <h2><i class="fas fa-wine-bottle"></i> <% If FEATURE_I18N Then Response.Write T("customize_bottles_title", Empty) Else %>精美瓶身<% End If %></h2>
                <p><% If FEATURE_I18N Then Response.Write T("customize_bottles_desc", Empty) Else %>选择心仪的瓶身，完美呈现您的专属香水<% End If %></p>
            </div>
            <div class="bottles-grid">
                <%
                Dim rsBottles
                Set rsBottles = ExecuteQuery("SELECT * FROM BottleStyles WHERE IsActive <> 0")
                If Not rsBottles Is Nothing Then
                    Do While Not rsBottles.EOF
                %>
                <div class="bottle-card">
                    <div class="bottle-image">
                        <img src="<%= rsBottles("ImageURL") %>" alt="<%= HTMLEncode(rsBottles("BottleName")) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                    </div>
                    <h4><%= HTMLEncode(rsBottles("BottleName")) %></h4>
                    <p><%= HTMLEncode(rsBottles("Description")) %></p>
                    <%
                Dim bottlePriceAdd
                bottlePriceAdd = 0
                If Not IsNull(rsBottles("PriceAddition")) Then
                    ' DECIMAL类型不能用IsNumeric判断，直接CDbl转换
                    On Error Resume Next
                    bottlePriceAdd = CDbl(rsBottles("PriceAddition"))
                    If Err.Number <> 0 Then bottlePriceAdd = 0
                    On Error GoTo 0
                End If
                If bottlePriceAdd > 0 Then
                %>
                    <span class="bottle-price">+<%= FormatMoney(bottlePriceAdd) %></span>
                    <% Else %>
                    <span class="bottle-price"><% If FEATURE_I18N Then Response.Write T("customize_bottle_free", Empty) Else %>免费<% End If %></span>
                    <% End If %>
                </div>
                <%
                        rsBottles.MoveNext
                    Loop
                    rsBottles.Close
                    Set rsBottles = Nothing
                End If
                %>
            </div>
        </section>

        <!-- 定制FAQ -->
        <section class="customize-section faq-section">
            <div class="section-header">
                <h2><i class="fas fa-question-circle"></i> <% If FEATURE_I18N Then Response.Write T("customize_faq_title", Empty) Else %>常见问题<% End If %></h2>
            </div>
            <div class="faq-list">
                <div class="faq-item">
                    <div class="faq-question">
                        <span><% If FEATURE_I18N Then Response.Write T("customize_faq_q1", Empty) Else %>定制香水需要多长时间？<% End If %></span>
                        <i class="fas fa-chevron-down"></i>
                    </div>
                    <div class="faq-answer">
                        <p><% If FEATURE_I18N Then Response.Write T("customize_faq_a1", Empty) Else %>定制香水需要3-5个工作日进行调配和制作，之后通过顺丰快递发货，一般2-5天送达。<% End If %></p>
                    </div>
                </div>
                <div class="faq-item">
                    <div class="faq-question">
                        <span><% If FEATURE_I18N Then Response.Write T("customize_faq_q2", Empty) Else %>定制香水可以退换吗？<% End If %></span>
                        <i class="fas fa-chevron-down"></i>
                    </div>
                    <div class="faq-answer">
                        <p><% If FEATURE_I18N Then Response.Write T("customize_faq_a2", Empty) Else %>由于定制香水是根据您的个人选择专门调配的，非质量问题不支持退换。如果收到的产品存在质量问题，请在签收后7天内联系客服处理。<% End If %></p>
                    </div>
                </div>
                <div class="faq-item">
                    <div class="faq-question">
                        <span><% If FEATURE_I18N Then Response.Write T("customize_faq_q3", Empty) Else %>香水的留香时间有多长？<% End If %></span>
                        <i class="fas fa-chevron-down"></i>
                    </div>
                    <div class="faq-answer">
                        <p><% If FEATURE_I18N Then Response.Write T("customize_faq_a3", Empty) Else %>我们的香水均采用优质香精，一般留香时间为6-8小时，具体时长会因个人肤质、环境温度等因素有所不同。<% End If %></p>
                    </div>
                </div>
                <div class="faq-item">
                    <div class="faq-question">
                        <span><% If FEATURE_I18N Then Response.Write T("customize_faq_q4", Empty) Else %>如何选择适合自己的香调？<% End If %></span>
                        <i class="fas fa-chevron-down"></i>
                    </div>
                    <div class="faq-answer">
                        <p><% If FEATURE_I18N Then Response.Write T("customize_faq_a4", Empty) Else %>建议根据您的个人喜好和使用场合来选择：日常使用可选清新淡雅的花香调或海洋调；约会或晚宴可选浓郁神秘的东方调；职场可选沉稳内敛的木质调。<% End If %></p>
                    </div>
                </div>
            </div>
        </section>

        <!-- 开始定制CTA -->
        <section class="customize-cta">
            <div class="cta-content">
                <h2><% If FEATURE_I18N Then Response.Write T("customize_cta_title", Empty) Else %>准备好创造您的专属香水了吗？<% End If %></h2>
                <p><% If FEATURE_I18N Then Response.Write T("customize_cta_desc", Empty) Else %>选择一款基础香型，开始您的定制之旅<% End If %></p>
                <a href="#" onclick="scrollToProducts()" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then Response.Write T("customize_cta_btn", Empty) Else %>立即开始<% End If %></a>
            </div>
        </section>
    </div>
</div>

<style>
.customize-hero {
    background: linear-gradient(135deg, var(--bg-dark) 0%, var(--primary-color) 100%);
    color: #fff;
    padding: 80px 0;
    text-align: center;
}

.customize-hero h1 {
    font-size: 42px;
    margin-bottom: 15px;
}

.customize-hero p {
    font-size: 18px;
    opacity: 0.9;
}

.customize-page {
    padding: 60px 0;
}

.customize-intro {
    margin-bottom: 60px;
}

.intro-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 30px;
}

.intro-item {
    text-align: center;
    padding: 30px;
    background: #fff;
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow);
}

.intro-icon {
    width: 60px;
    height: 60px;
    background: var(--primary-color);
    color: #fff;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
    font-weight: bold;
    margin: 0 auto 20px;
}

.intro-item h3 {
    font-size: 18px;
    margin-bottom: 10px;
}

.intro-item p {
    color: var(--text-light);
    font-size: 14px;
}

.customize-section {
    margin-bottom: 60px;
}

.base-products {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 30px;
}

.base-product-card {
    background: #fff;
    border-radius: var(--radius-lg);
    overflow: hidden;
    box-shadow: var(--shadow);
    transition: var(--transition);
}

.base-product-card:hover {
    transform: translateY(-5px);
    box-shadow: var(--shadow-lg);
}

.card-image {
    height: 200px;
    overflow: hidden;
}

.card-image img {
    width: 100%;
    height: 100%;
    object-fit: cover;
}

.card-content {
    padding: 25px;
}

.card-category {
    display: inline-block;
    padding: 3px 10px;
    background: var(--accent-color);
    color: var(--primary-color);
    border-radius: 20px;
    font-size: 12px;
    margin-bottom: 10px;
}

.card-content h3 {
    font-size: 20px;
    margin-bottom: 10px;
}

.card-content p {
    color: var(--text-light);
    font-size: 14px;
    margin-bottom: 20px;
    line-height: 1.6;
}

.card-footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.card-price {
    font-size: 20px;
    font-weight: bold;
    color: var(--primary-color);
}

.notes-explanation {
    display: flex;
    flex-direction: column;
    gap: 20px;
}

.note-level {
    display: flex;
    gap: 25px;
    padding: 30px;
    border-radius: var(--radius-lg);
    background: #fff;
    box-shadow: var(--shadow);
}

.top-note {
    border-left: 4px solid #FFD700;
}

.middle-note {
    border-left: 4px solid #FF69B4;
}

.base-note {
    border-left: 4px solid #8B4513;
}

.note-icon {
    width: 60px;
    height: 60px;
    background: var(--bg-light);
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}

.note-icon i {
    font-size: 24px;
    color: var(--primary-color);
}

.note-content h3 {
    font-size: 18px;
    margin-bottom: 10px;
}

.note-content p {
    color: var(--text-light);
    margin-bottom: 15px;
    line-height: 1.6;
}

.available-notes {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
}

.note-tag {
    padding: 5px 15px;
    background: var(--bg-light);
    border-radius: 20px;
    font-size: 13px;
    color: var(--text-color);
}

.bottles-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 20px;
}

.bottle-card {
    background: #fff;
    padding: 25px;
    border-radius: var(--radius-lg);
    text-align: center;
    box-shadow: var(--shadow);
    transition: var(--transition);
}

.bottle-card:hover {
    transform: translateY(-5px);
    box-shadow: var(--shadow-lg);
}

.bottle-image {
    height: 120px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 15px;
}

.bottle-image img {
    max-height: 100%;
    max-width: 100%;
}

.bottle-card h4 {
    font-size: 16px;
    margin-bottom: 8px;
}

.bottle-card p {
    font-size: 12px;
    color: var(--text-light);
    margin-bottom: 10px;
}

.bottle-price {
    color: var(--primary-color);
    font-weight: 600;
}

.faq-list {
    max-width: 800px;
    margin: 0 auto;
}

.faq-item {
    background: #fff;
    margin-bottom: 15px;
    border-radius: var(--radius);
    overflow: hidden;
    box-shadow: var(--shadow);
}

.faq-question {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 20px 25px;
    cursor: pointer;
    font-weight: 600;
}

.faq-question i {
    color: var(--primary-color);
    transition: var(--transition);
}

.faq-item.active .faq-question i {
    transform: rotate(180deg);
}

.faq-answer {
    display: none;
    padding: 0 25px 20px;
}

.faq-answer p {
    color: var(--text-light);
    line-height: 1.8;
}

.faq-item.active .faq-answer {
    display: block;
}

.customize-cta {
    background: linear-gradient(135deg, var(--primary-color) 0%, var(--bg-dark) 100%);
    color: #fff;
    padding: 60px;
    border-radius: var(--radius-lg);
    text-align: center;
}

.customize-cta h2 {
    font-size: 28px;
    margin-bottom: 10px;
}

.customize-cta p {
    opacity: 0.9;
    margin-bottom: 25px;
}

@media (max-width: 992px) {
    .intro-grid {
        grid-template-columns: repeat(2, 1fr);
    }
    
    .base-products {
        grid-template-columns: repeat(2, 1fr);
    }
    
    .bottles-grid {
        grid-template-columns: repeat(3, 1fr);
    }
}

@media (max-width: 768px) {
    .intro-grid {
        grid-template-columns: 1fr;
    }
    
    .base-products {
        grid-template-columns: 1fr;
    }
    
    .bottles-grid {
        grid-template-columns: repeat(2, 1fr);
    }
    
    .note-level {
        flex-direction: column;
        text-align: center;
    }
    
    .note-icon {
        margin: 0 auto;
    }
}
</style>

<script>
$(document).ready(function() {
    // FAQ切换
    $('.faq-question').click(function() {
        var $item = $(this).parent();
        $item.toggleClass('active');
    });
});

function scrollToProducts() {
    $('html, body').animate({
        scrollTop: $('.base-products').offset().top - 100
    }, 500);
}
</script>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
