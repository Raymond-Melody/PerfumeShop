<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()

' 处理表单提交
Dim successMsg
successMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        ' 不显示成功消息，让页面显示错误
        successMsg = ""
    Else
        ' 这里可以保存联系信息到数据库或发送邮件
        successMsg = T("contact_success_msg", Empty)
    End If
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()
%>
<!--#include file="includes/header.asp"-->

<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp"><% If FEATURE_I18N Then %><%= T("breadcrumb_home", Empty) %><% Else %>首页<% End If %></a>
        <span class="separator">/</span>
        <span><% If FEATURE_I18N Then %><%= T("contact_breadcrumb", Empty) %><% Else %>联系我们<% End If %></span>
    </div>
</div>

<div class="container">
    <div class="contact-page">
        <div class="contact-header">
            <h1><% If FEATURE_I18N Then %><%= T("contact_header_title", Empty) %><% Else %>联系我们<% End If %></h1>
            <p><% If FEATURE_I18N Then %><%= T("contact_header_desc", Empty) %><% Else %>有任何问题或建议，欢迎与我们联系<% End If %></p>
        </div>

        <div class="contact-content">
            <!-- 联系方式 -->
            <div class="contact-info">
                <div class="info-card">
                    <div class="info-icon"><i class="fas fa-phone"></i></div>
                    <h3><% If FEATURE_I18N Then %><%= T("contact_info_phone", Empty) %><% Else %>客服热线<% End If %></h3>
                    <p><%= SITE_PHONE %></p>
                    <span><% If FEATURE_I18N Then %><%= T("contact_info_hours", Empty) %><% Else %>周一至周日 9:00-21:00<% End If %></span>
                </div>
                <div class="info-card">
                    <div class="info-icon"><i class="fas fa-envelope"></i></div>
                    <h3><% If FEATURE_I18N Then %><%= T("contact_info_email", Empty) %><% Else %>电子邮箱<% End If %></h3>
                    <p><%= SITE_EMAIL %></p>
                    <span><% If FEATURE_I18N Then %><%= T("contact_info_email_hint", Empty) %><% Else %>工作日24小时内回复<% End If %></span>
                </div>
                <div class="info-card">
                    <div class="info-icon"><i class="fab fa-weixin"></i></div>
                    <h3><% If FEATURE_I18N Then %><%= T("contact_info_wechat", Empty) %><% Else %>官方微信<% End If %></h3>
                    <p>PerfumeCustom</p>
                    <span><% If FEATURE_I18N Then %><%= T("contact_info_wechat_hint", Empty) %><% Else %>扫码关注获取优惠<% End If %></span>
                </div>
                <div class="info-card">
                    <div class="info-icon"><i class="fas fa-map-marker-alt"></i></div>
                    <h3><% If FEATURE_I18N Then %><%= T("contact_info_address", Empty) %><% Else %>公司地址<% End If %></h3>
                    <p><% If FEATURE_I18N Then %><%= T("contact_info_addr_text", Empty) %><% Else %>上海市静安区南京西路<% End If %></p>
                    <span><% If FEATURE_I18N Then %><%= T("contact_info_addr_hint", Empty) %><% Else %>香氛定制体验中心<% End If %></span>
                </div>
            </div>

            <!-- 联系表单 -->
            <div class="contact-form-section">
                <h2><% If FEATURE_I18N Then %><%= T("contact_form_title", Empty) %><% Else %>给我们留言<% End If %></h2>
                
                <% If successMsg <> "" Then %>
                <div class="alert alert-success">
                    <i class="fas fa-check-circle"></i> <%= successMsg %>
                </div>
                <% End If %>
                
                <form method="post" class="contact-form">
                    <%= GetCSRFTokenField() %>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="name"><% If FEATURE_I18N Then %><%= T("contact_form_name", Empty) %><% Else %>您的姓名 *<% End If %></label>
                            <input type="text" id="name" name="name" required>
                        </div>
                        <div class="form-group">
                            <label for="email"><% If FEATURE_I18N Then %><%= T("contact_form_email", Empty) %><% Else %>电子邮箱 *<% End If %></label>
                            <input type="email" id="email" name="email" required>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="phone"><% If FEATURE_I18N Then %><%= T("contact_form_phone", Empty) %><% Else %>联系电话<% End If %></label>
                            <input type="tel" id="phone" name="phone">
                        </div>
                        <div class="form-group">
                            <label for="subject"><% If FEATURE_I18N Then %><%= T("contact_form_subject", Empty) %><% Else %>主题<% End If %></label>
                            <select id="subject" name="subject">
                                <option value="咨询"><% If FEATURE_I18N Then %><%= T("contact_form_subject_consult", Empty) %><% Else %>产品咨询<% End If %></option>
                                <option value="定制"><% If FEATURE_I18N Then %><%= T("contact_form_subject_custom", Empty) %><% Else %>定制服务<% End If %></option>
                                <option value="售后"><% If FEATURE_I18N Then %><%= T("contact_form_subject_after", Empty) %><% Else %>售后服务<% End If %></option>
                                <option value="合作"><% If FEATURE_I18N Then %><%= T("contact_form_subject_biz", Empty) %><% Else %>商务合作<% End If %></option>
                                <option value="其他"><% If FEATURE_I18N Then %><%= T("contact_form_subject_other", Empty) %><% Else %>其他<% End If %></option>
                            </select>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="message"><% If FEATURE_I18N Then %><%= T("contact_form_message", Empty) %><% Else %>留言内容 *<% End If %></label>
                        <textarea id="message" name="message" rows="5" required></textarea>
                    </div>
                    <button type="submit" class="btn btn-primary btn-lg">
                        <i class="fas fa-paper-plane"></i> <% If FEATURE_I18N Then %><%= T("contact_form_send", Empty) %><% Else %>发送留言<% End If %>
                    </button>
                </form>
            </div>
        </div>
    </div>
</div>

<style>
.contact-page {
    padding: 40px 0;
}

.contact-header {
    text-align: center;
    margin-bottom: 50px;
}

.contact-header h1 {
    font-size: 36px;
    color: var(--primary-color);
    margin-bottom: 10px;
}

.contact-header p {
    color: var(--text-light);
    font-size: 16px;
}

.contact-content {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 50px;
}

.contact-info {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
}

.info-card {
    background: #fff;
    padding: 30px;
    border-radius: var(--radius-lg);
    text-align: center;
    box-shadow: var(--shadow);
    transition: var(--transition);
}

.info-card:hover {
    transform: translateY(-5px);
    box-shadow: var(--shadow-lg);
}

.info-icon {
    width: 60px;
    height: 60px;
    background: var(--accent-color);
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    margin: 0 auto 15px;
}

.info-icon i {
    font-size: 24px;
    color: var(--primary-color);
}

.info-card h3 {
    font-size: 16px;
    margin-bottom: 10px;
}

.info-card p {
    font-weight: 600;
    color: var(--primary-color);
    margin-bottom: 5px;
}

.info-card span {
    font-size: 12px;
    color: var(--text-muted);
}

.contact-form-section {
    background: #fff;
    padding: 40px;
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow);
}

.contact-form-section h2 {
    font-size: 22px;
    margin-bottom: 25px;
    color: var(--primary-color);
}

.contact-form .form-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
}

.contact-form .form-group {
    margin-bottom: 20px;
}

.contact-form label {
    display: block;
    margin-bottom: 8px;
    font-weight: 500;
}

.contact-form input,
.contact-form select,
.contact-form textarea {
    width: 100%;
    padding: 12px 15px;
    border: 1px solid var(--border-color);
    border-radius: var(--radius);
    outline: none;
    transition: var(--transition);
}

.contact-form input:focus,
.contact-form select:focus,
.contact-form textarea:focus {
    border-color: var(--primary-color);
}

@media (max-width: 992px) {
    .contact-content {
        grid-template-columns: 1fr;
    }
}

@media (max-width: 768px) {
    .contact-info {
        grid-template-columns: 1fr;
    }
    
    .contact-form .form-row {
        grid-template-columns: 1fr;
    }
}
</style>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
