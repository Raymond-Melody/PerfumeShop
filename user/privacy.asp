<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()
%>
<!--#include file="../includes/header.asp"-->

<div class="privacy-page">
    <div class="container">
        <div class="privacy-header">
            <h1><% If FEATURE_I18N Then %><%= T("privacy_title", Empty) %><% Else %>隐私政策<% End If %></h1>
            <p class="privacy-updated"><% If FEATURE_I18N Then %><%= T("privacy_last_updated", Empty) %><% Else %>最后更新：2026年1月<% End If %></p>
        </div>

        <div class="privacy-content">
            <section class="privacy-section">
                <h2><% If FEATURE_I18N Then %><%= T("privacy_s1_title", Empty) %><% Else %>1. 信息收集<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("privacy_s1_text", Empty) %><% Else %>我们收集以下类型的个人信息：<% End If %></p>
                <ul>
                    <li><strong><% If FEATURE_I18N Then %><%= T("privacy_s1_account", Empty) %><% Else %>账户信息<% End If %>：</strong><% If FEATURE_I18N Then %><%= T("privacy_s1_account_text", Empty) %><% Else %>用户名、邮箱、手机号、密码（加密存储）<% End If %></li>
                    <li><strong><% If FEATURE_I18N Then %><%= T("privacy_s1_profile", Empty) %><% Else %>个人资料<% End If %>：</strong><% If FEATURE_I18N Then %><%= T("privacy_s1_profile_text", Empty) %><% Else %>姓名、收货地址、香氛偏好<% End If %></li>
                    <li><strong><% If FEATURE_I18N Then %><%= T("privacy_s1_transaction", Empty) %><% Else %>交易信息<% End If %>：</strong><% If FEATURE_I18N Then %><%= T("privacy_s1_transaction_text", Empty) %><% Else %>订单记录、定制配方、支付记录（不含完整卡号）<% End If %></li>
                    <li><strong><% If FEATURE_I18N Then %><%= T("privacy_s1_technical", Empty) %><% Else %>技术信息<% End If %>：</strong><% If FEATURE_I18N Then %><%= T("privacy_s1_technical_text", Empty) %><% Else %>IP地址、浏览器类型、设备信息、Cookie<% End If %></li>
                </ul>
            </section>

            <section class="privacy-section">
                <h2><% If FEATURE_I18N Then %><%= T("privacy_s2_title", Empty) %><% Else %>2. 信息使用<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("privacy_s2_text", Empty) %><% Else %>我们使用收集的信息用于：<% End If %></p>
                <ul>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s2_order", Empty) %><% Else %>处理订单、配送和售后服务<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s2_personalize", Empty) %><% Else %>个性化香氛推荐与定制服务<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s2_improve", Empty) %><% Else %>改进产品和服务体验<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s2_legal", Empty) %><% Else %>遵守法律法规要求<% End If %></li>
                </ul>
            </section>

            <section class="privacy-section">
                <h2><% If FEATURE_I18N Then %><%= T("privacy_s3_title", Empty) %><% Else %>3. 信息保护<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("privacy_s3_text", Empty) %><% Else %>我们采用以下安全措施保护您的信息：<% End If %></p>
                <ul>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s3_encrypt", Empty) %><% Else %>数据传输使用 HTTPS/TLS 加密<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s3_password", Empty) %><% Else %>密码使用 SHA-512 加盐哈希存储<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s3_access", Empty) %><% Else %>严格的数据库访问控制和审计日志<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s3_csrf", Empty) %><% Else %>CSRF 防护和 API 速率限制<% End If %></li>
                </ul>
            </section>

            <section class="privacy-section">
                <h2><% If FEATURE_I18N Then %><%= T("privacy_s4_title", Empty) %><% Else %>4. Cookie 使用<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("privacy_s4_text", Empty) %><% Else %>我们使用以下类型的 Cookie：<% End If %></p>
                <ul>
                    <li><strong><% If FEATURE_I18N Then %><%= T("privacy_s4_essential", Empty) %><% Else %>必要 Cookie<% End If %>：</strong><% If FEATURE_I18N Then %><%= T("privacy_s4_essential_text", Empty) %><% Else %>维持网站基本功能，如登录会话、购物车<% End If %></li>
                    <li><strong><% If FEATURE_I18N Then %><%= T("privacy_s4_analytics", Empty) %><% Else %>分析 Cookie<% End If %>：</strong><% If FEATURE_I18N Then %><%= T("privacy_s4_analytics_text", Empty) %><% Else %>帮助我们了解网站使用情况，改进服务<% End If %></li>
                    <li><strong><% If FEATURE_I18N Then %><%= T("privacy_s4_preference", Empty) %><% Else %>偏好 Cookie<% End If %>：</strong><% If FEATURE_I18N Then %><%= T("privacy_s4_preference_text", Empty) %><% Else %>记住您的语言偏好和显示设置<% End If %></li>
                </ul>
            </section>

            <section class="privacy-section">
                <h2><% If FEATURE_I18N Then %><%= T("privacy_s5_title", Empty) %><% Else %>5. 您的权利<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("privacy_s5_text", Empty) %><% Else %>根据适用的数据保护法律，您拥有以下权利：<% End If %></p>
                <ul class="privacy-rights">
                    <li>
                        <i class="fas fa-download"></i>
                        <strong><% If FEATURE_I18N Then %><%= T("privacy_right_access", Empty) %><% Else %>数据访问权<% End If %></strong>
                        <span><% If FEATURE_I18N Then %><%= T("privacy_right_access_text", Empty) %><% Else %>您可以请求获取我们持有的您的个人数据副本<% End If %></span>
                        <a href="data_export.asp" class="btn btn-sm btn-outline"><% If FEATURE_I18N Then %><%= T("privacy_btn_export", Empty) %><% Else %>导出数据<% End If %></a>
                    </li>
                    <li>
                        <i class="fas fa-edit"></i>
                        <strong><% If FEATURE_I18N Then %><%= T("privacy_right_rectify", Empty) %><% Else %>更正权<% End If %></strong>
                        <span><% If FEATURE_I18N Then %><%= T("privacy_right_rectify_text", Empty) %><% Else %>您可以在账户设置中更新您的个人信息<% End If %></span>
                        <a href="settings.asp" class="btn btn-sm btn-outline"><% If FEATURE_I18N Then %><%= T("privacy_btn_settings", Empty) %><% Else %>账户设置<% End If %></a>
                    </li>
                    <li>
                        <i class="fas fa-trash-alt"></i>
                        <strong><% If FEATURE_I18N Then %><%= T("privacy_right_delete", Empty) %><% Else %>删除权（被遗忘权）<% End If %></strong>
                        <span><% If FEATURE_I18N Then %><%= T("privacy_right_delete_text", Empty) %><% Else %>您可以请求删除您的账户和个人数据（30天冷静期）<% End If %></span>
                        <a href="account_delete.asp" class="btn btn-sm btn-outline btn-danger"><% If FEATURE_I18N Then %><%= T("privacy_btn_delete", Empty) %><% Else %>注销账户<% End If %></a>
                    </li>
                    <li>
                        <i class="fas fa-ban"></i>
                        <strong><% If FEATURE_I18N Then %><%= T("privacy_right_restrict", Empty) %><% Else %>限制处理权<% End If %></strong>
                        <span><% If FEATURE_I18N Then %><%= T("privacy_right_restrict_text", Empty) %><% Else %>您可以限制我们对您数据的某些处理方式<% End If %></span>
                    </li>
                    <li>
                        <i class="fas fa-file-export"></i>
                        <strong><% If FEATURE_I18N Then %><%= T("privacy_right_portability", Empty) %><% Else %>数据可携带权<% End If %></strong>
                        <span><% If FEATURE_I18N Then %><%= T("privacy_right_portability_text", Empty) %><% Else %>您可以请求将数据以结构化格式导出并转移<% End If %></span>
                    </li>
                </ul>
            </section>

            <section class="privacy-section">
                <h2><% If FEATURE_I18N Then %><%= T("privacy_s6_title", Empty) %><% Else %>6. 数据保留<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("privacy_s6_text", Empty) %><% Else %>我们仅在实现收集目的所必需的期限内保留您的个人数据：<% End If %></p>
                <ul>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s6_account", Empty) %><% Else %>账户数据：账户存续期间及注销后30天<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s6_order", Empty) %><% Else %>订单数据：根据税务法规保留7年<% End If %></li>
                    <li><% If FEATURE_I18N Then %><%= T("privacy_s6_log", Empty) %><% Else %>访问日志：保留90天<% End If %></li>
                </ul>
            </section>

            <section class="privacy-section">
                <h2><% If FEATURE_I18N Then %><%= T("privacy_s7_title", Empty) %><% Else %>7. 联系我们<% End If %></h2>
                <p><% If FEATURE_I18N Then %><%= T("privacy_s7_text", Empty) %><% Else %>如对隐私政策有任何疑问，请联系：<% End If %></p>
                <ul class="privacy-contact">
                    <li><i class="fas fa-envelope"></i> <%= SITE_EMAIL %></li>
                    <li><i class="fas fa-phone"></i> <%= SITE_PHONE %></li>
                </ul>
            </section>
        </div>
    </div>
</div>

<style>
.privacy-page { padding: 40px 0; max-width: 900px; margin: 0 auto; }
.privacy-header { text-align: center; margin-bottom: 40px; }
.privacy-header h1 { font-size: 2rem; color: #2d3748; margin-bottom: 8px; }
.privacy-updated { color: #a0aec0; font-size: 0.9rem; }
.privacy-section { margin-bottom: 32px; padding-bottom: 24px; border-bottom: 1px solid #edf2f7; }
.privacy-section:last-child { border-bottom: none; }
.privacy-section h2 { font-size: 1.3rem; color: #4a5568; margin-bottom: 12px; }
.privacy-section p { color: #718096; line-height: 1.8; margin-bottom: 12px; }
.privacy-section ul { padding-left: 20px; }
.privacy-section li { color: #718096; line-height: 1.8; margin-bottom: 6px; }
.privacy-rights { list-style: none; padding: 0; }
.privacy-rights li { display: flex; flex-wrap: wrap; align-items: center; gap: 12px; padding: 16px; background: #f7fafc; border-radius: 8px; margin-bottom: 12px; }
.privacy-rights li i { font-size: 1.2rem; color: #667eea; width: 24px; text-align: center; flex-shrink: 0; }
.privacy-rights li strong { min-width: 100px; color: #2d3748; }
.privacy-rights li span { flex: 1; min-width: 200px; color: #718096; font-size: 0.9rem; }
.privacy-contact { list-style: none; padding: 0; }
.privacy-contact li { display: inline-flex; align-items: center; gap: 6px; margin-right: 20px; }
.privacy-contact li i { color: #667eea; }
.btn-danger { border-color: #e53e3e; color: #e53e3e; }
.btn-danger:hover { background: #e53e3e; color: #fff; }
@media (max-width: 768px) {
    .privacy-rights li { flex-direction: column; align-items: flex-start; }
    .privacy-rights li span { min-width: auto; }
}
</style>

<!--#include file="../includes/footer.asp"-->
