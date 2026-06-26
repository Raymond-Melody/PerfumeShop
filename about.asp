<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()
%>
<!--#include file="includes/header.asp"-->

<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then Response.Write T("breadcrumb_home", Empty) Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then Response.Write T("about_breadcrumb", Empty) Else %>品牌故事<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="about-page">
        <!-- 品牌介绍 -->
        <section class="about-hero">
            <div class="hero-content">
                <h1><% If FEATURE_I18N Then Response.Write T("about_hero_title", Empty) Else %>关于香氛定制<% End If %></h1>
                <p class="tagline"><% If FEATURE_I18N Then Response.Write T("about_hero_tagline", Empty) Else %>让每一瓶香水都是独一无二的存在<% End If %></p>
            </div>
        </section>

        <!-- 品牌故事 -->
        <section class="about-section">
            <div class="section-grid">
                <div class="section-image">
                    <div class="about-image-placeholder"></div>
                </div>
                <div class="section-content">
                    <h2><% If FEATURE_I18N Then Response.Write T("about_story_title", Empty) Else %>我们的故事<% End If %></h2>
                    <p><% If FEATURE_I18N Then Response.Write T("about_story_para1", Empty) Else %>香氛定制诞生于2018年，源自创始人对香水艺术的热爱与追求。<% End If %></p>
                    <p><% If FEATURE_I18N Then Response.Write T("about_story_para2", Empty) Else %>我们相信，香水不仅仅是一种气味，更是个性与情感的表达。每个人都值得拥有一款独一无二的香水，它能够完美诠释你的独特魅力。<% End If %></p>
                    <p><% If FEATURE_I18N Then Response.Write T("about_story_para3", Empty) Else %>从最初的小工作室到如今的专业定制品牌，我们始终坚持"以人为本，匠心定制"的理念，为每一位顾客打造专属的香氛体验。<% End If %></p>
                </div>
            </div>
        </section>

        <!-- 我们的理念 -->
        <section class="about-section values-section">
            <h2 class="section-title"><% If FEATURE_I18N Then Response.Write T("about_values_title", Empty) Else %>我们的理念<% End If %></h2>
            <div class="values-grid">
                <div class="value-item">
                    <div class="value-icon"><i class="fas fa-gem"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("about_value_quality_title", Empty) Else %>品质至上<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("about_value_quality_desc", Empty) Else %>精选全球优质天然香料，每一滴香精都经过严格筛选，确保最纯正的香气体验。<% End If %></p>
                </div>
                <div class="value-item">
                    <div class="value-icon"><i class="fas fa-fingerprint"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("about_value_custom_title", Empty) Else %>个性定制<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("about_value_custom_desc", Empty) Else %>每一瓶香水都根据您的喜好量身定制，让香氛成为您独特个性的延伸。<% End If %></p>
                </div>
                <div class="value-item">
                    <div class="value-icon"><i class="fas fa-leaf"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("about_value_eco_title", Empty) Else %>环保可持续<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("about_value_eco_desc", Empty) Else %>采用环保包装材料，支持可持续发展，让美好与责任并行。<% End If %></p>
                </div>
                <div class="value-item">
                    <div class="value-icon"><i class="fas fa-heart"></i></div>
                    <h3><% If FEATURE_I18N Then Response.Write T("about_value_service_title", Empty) Else %>用心服务<% End If %></h3>
                    <p><% If FEATURE_I18N Then Response.Write T("about_value_service_desc", Empty) Else %>专业的调香顾问团队，为您提供一对一的定制建议和贴心服务。<% End If %></p>
                </div>
            </div>
        </section>

        <!-- 专业团队 -->
        <section class="about-section team-section">
            <h2 class="section-title"><% If FEATURE_I18N Then Response.Write T("about_team_title", Empty) Else %>专业团队<% End If %></h2>
            <p class="section-desc"><% If FEATURE_I18N Then Response.Write T("about_team_desc", Empty) Else %>我们的调香师团队拥有超过20年的行业经验<% End If %></p>
            <div class="team-stats">
                <div class="stat">
                    <span class="stat-number">20+</span>
                    <span class="stat-label"><% If FEATURE_I18N Then Response.Write T("about_stat_exp", Empty) Else %>年行业经验<% End If %></span>
                </div>
                <div class="stat">
                    <span class="stat-number">50+</span>
                    <span class="stat-label"><% If FEATURE_I18N Then Response.Write T("about_stat_notes", Empty) Else %>种优质香料<% End If %></span>
                </div>
                <div class="stat">
                    <span class="stat-number">10000+</span>
                    <span class="stat-label"><% If FEATURE_I18N Then Response.Write T("about_stat_customers", Empty) Else %>满意客户<% End If %></span>
                </div>
                <div class="stat">
                    <span class="stat-number">99%</span>
                    <span class="stat-label"><% If FEATURE_I18N Then Response.Write T("about_stat_rate", Empty) Else %>好评率<% End If %></span>
                </div>
            </div>
        </section>

        <!-- CTA -->
        <section class="about-cta">
            <h2><% If FEATURE_I18N Then Response.Write T("about_cta_title", Empty) Else %>开始您的香氛定制之旅<% End If %></h2>
            <p><% If FEATURE_I18N Then Response.Write T("about_cta_desc", Empty) Else %>让我们为您创造独一无二的专属香水<% End If %></p>
            <a href="/customize.asp" class="btn btn-primary btn-lg"><% If FEATURE_I18N Then Response.Write T("about_cta_btn", Empty) Else %>立即定制<% End If %></a>
        </section>
    </div>
</div>

<style>
.about-page {
    padding: 40px 0;
}

.about-hero {
    text-align: center;
    padding: 60px 0;
    background: linear-gradient(135deg, var(--bg-light) 0%, #fff 100%);
    border-radius: var(--radius-lg);
    margin-bottom: 60px;
}

.about-hero h1 {
    font-size: 42px;
    color: var(--primary-color);
    margin-bottom: 15px;
}

.tagline {
    font-size: 20px;
    color: var(--text-light);
}

.about-section {
    margin-bottom: 60px;
}

.section-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 60px;
    align-items: center;
}

.section-image img {
    width: 100%;
    border-radius: var(--radius-lg);
}

.section-content h2 {
    font-size: 28px;
    color: var(--primary-color);
    margin-bottom: 20px;
}

.section-content p {
    color: var(--text-light);
    line-height: 1.8;
    margin-bottom: 15px;
}

.section-title {
    font-size: 28px;
    color: var(--primary-color);
    text-align: center;
    margin-bottom: 15px;
}

.section-desc {
    text-align: center;
    color: var(--text-light);
    margin-bottom: 40px;
}

.values-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 30px;
}

.value-item {
    text-align: center;
    padding: 30px;
    background: #fff;
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow);
}

.value-icon {
    width: 80px;
    height: 80px;
    background: var(--accent-color);
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    margin: 0 auto 20px;
}

.value-icon i {
    font-size: 32px;
    color: var(--primary-color);
}

.value-item h3 {
    font-size: 18px;
    margin-bottom: 10px;
}

.value-item p {
    color: var(--text-light);
    font-size: 14px;
    line-height: 1.6;
}

.team-stats {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 30px;
}

.stat {
    text-align: center;
    padding: 40px;
    background: #fff;
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow);
}

.stat-number {
    display: block;
    font-size: 48px;
    font-weight: bold;
    color: var(--primary-color);
    margin-bottom: 10px;
}

.stat-label {
    color: var(--text-light);
}

.about-cta {
    text-align: center;
    padding: 60px;
    background: linear-gradient(135deg, var(--primary-color), var(--bg-dark));
    color: #fff;
    border-radius: var(--radius-lg);
}

.about-cta h2 {
    font-size: 28px;
    margin-bottom: 10px;
}

.about-cta p {
    opacity: 0.9;
    margin-bottom: 25px;
}

@media (max-width: 992px) {
    .section-grid {
        grid-template-columns: 1fr;
    }
    
    .values-grid,
    .team-stats {
        grid-template-columns: repeat(2, 1fr);
    }
}

@media (max-width: 768px) {
    .values-grid,
    .team-stats {
        grid-template-columns: 1fr;
    }
}
</style>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
