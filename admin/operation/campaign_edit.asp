<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' 获取活动ID（编辑模式）
Dim campaignId, isEditMode
campaignId = Request.QueryString("id")
isEditMode = False

If campaignId <> "" And IsNumeric(campaignId) Then
    isEditMode = True
End If

' 初始化变量
Dim campaignName, description, campaignType, startDate, endDate, discountValue, minPurchase, isActive
campaignName = ""
description = ""
campaignType = "discount"
startDate = Date()
endDate = DateAdd("d", 30, Date())
discountValue = ""
minPurchase = ""
isActive = True

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' 处理POST请求
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面重试"
    Else
        ' 获取表单数据
        campaignName = Trim(Request.Form("campaign_name"))
        description = Trim(Request.Form("description"))
        campaignType = Request.Form("campaign_type")
        startDate = Request.Form("start_date")
        endDate = Request.Form("end_date")
        discountValue = Request.Form("discount_value")
        minPurchase = Request.Form("min_purchase")
        isActive = (Request.Form("is_active") = "1")
        
        ' 验证必填字段
        If campaignName = "" Then
            errorMsg = "请输入活动名称"
        ElseIf campaignType = "" Then
            errorMsg = "请选择活动类型"
        ElseIf startDate = "" Or endDate = "" Then
            errorMsg = "请选择开始和结束日期"
        ElseIf Not IsDate(startDate) Or Not IsDate(endDate) Then
            errorMsg = "日期格式无效"
        ElseIf CDate(startDate) > CDate(endDate) Then
            errorMsg = "结束日期不能早于开始日期"
        Else
            ' 安全处理SQL字符串
            Dim safeCampaignName, safeDescription
            safeCampaignName = SafeSQL(campaignName)
            safeDescription = SafeSQL(description)
            
            Dim sql, result
            
            If isEditMode Then
                ' 更新现有活动
                ' 转换日期格式
                Dim accessStartDate, accessEndDate
                If IsDate(startDate) Then accessStartDate = CDate(startDate) Else accessStartDate = Date()
                If IsDate(endDate) Then accessEndDate = CDate(endDate) Else accessEndDate = Date()
                
                sql = "UPDATE MarketingCampaigns SET " & _
                      "CampaignName = '" & safeCampaignName & "', " & _
                      "Description = '" & safeDescription & "', " & _
                      "CampaignType = '" & campaignType & "', " & _
                      "StartDate = #" & Month(accessStartDate) & "/" & Day(accessStartDate) & "/" & Year(accessStartDate) & "#, " & _
                      "EndDate = #" & Month(accessEndDate) & "/" & Day(accessEndDate) & "/" & Year(accessEndDate) & "#, " & _
                      "DiscountValue = " & IIf(discountValue <> "", discountValue, "0") & ", " & _
                      "MinPurchase = " & IIf(minPurchase <> "", minPurchase, "0") & ", " & _
                      "IsActive = " & IIf(isActive, "True", "False") & " " & _
                      "WHERE CampaignID = " & campaignId
                
                result = ExecuteNonQuery(sql)
                
                If result Then
                    Call LogAdminAction("编辑营销活动", "operation", "MarketingCampaigns", campaignId, safeCampaignName)
                    Response.Redirect "marketing.asp?msg=updated"
                    Response.End
                Else
                    errorMsg = "更新失败: " & Session("LastDBError")
                End If
            Else
                ' 创建新活动
                ' 转换日期格式
                If IsDate(startDate) Then accessStartDate = CDate(startDate) Else accessStartDate = Date()
                If IsDate(endDate) Then accessEndDate = CDate(endDate) Else accessEndDate = Date()
                
                sql = "INSERT INTO MarketingCampaigns (CampaignName, Description, CampaignType, StartDate, EndDate, DiscountValue, MinPurchase, IsActive, ParticipantCount, TotalSales) " & _
                      "VALUES ('" & safeCampaignName & "', '" & safeDescription & "', '" & campaignType & "', #" & Month(accessStartDate) & "/" & Day(accessStartDate) & "/" & Year(accessStartDate) & "#, #" & Month(accessEndDate) & "/" & Day(accessEndDate) & "/" & Year(accessEndDate) & "#, " & _
                      IIf(discountValue <> "", discountValue, "0") & ", " & IIf(minPurchase <> "", minPurchase, "0") & ", " & _
                      IIf(isActive, "True", "False") & ", 0, 0)"
                
                
                result = ExecuteNonQuery(sql)
                
                If result Then
                    Dim newId
                    newId = GetLastInsertID("MarketingCampaigns")
                    Call LogAdminAction("创建营销活动", "operation", "MarketingCampaigns", newId, safeCampaignName)
                    Response.Redirect "marketing.asp?msg=created"
                    Response.End
                Else
                    errorMsg = "创建失败: " & Session("LastDBError")
                End If
            End If
        End If
    End If
ElseIf isEditMode Then
    ' 编辑模式：加载现有数据
    Dim rsCampaign
    Set rsCampaign = ExecuteQuery("SELECT * FROM MarketingCampaigns WHERE CampaignID = " & campaignId)
    
    If Not rsCampaign Is Nothing And Not rsCampaign.EOF Then
        campaignName = rsCampaign("CampaignName")
        description = rsCampaign("Description")
        campaignType = rsCampaign("CampaignType")
        startDate = rsCampaign("StartDate")
        endDate = rsCampaign("EndDate")
        discountValue = rsCampaign("DiscountValue")
        minPurchase = rsCampaign("MinPurchase")
        isActive = rsCampaign("IsActive")
        rsCampaign.Close
    Else
        errorMsg = "活动不存在"
        isEditMode = False
    End If
    Set rsCampaign = Nothing
End If

Call LogAdminAction(IIf(isEditMode, "编辑营销活动页面", "创建营销活动页面"), "operation", "MarketingCampaigns", campaignId, "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title><%= IIf(isEditMode, "编辑", "创建") %>营销活动 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .form-container { max-width: 800px; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; font-weight: 500; color: #333; }
        .form-group label .required { color: #e74c3c; margin-left: 4px; }
        .form-control { width: 100%; padding: 12px 15px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; box-sizing: border-box; }
        .form-control:focus { outline: none; border-color: #667eea; }
        textarea.form-control { min-height: 100px; resize: vertical; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .form-actions { display: flex; gap: 15px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #f0f0f0; }
        .alert { padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .alert-error { background: #ffebee; color: #c62828; border: 1px solid #ffcdd2; }
        .alert-success { background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; }
        .checkbox-group { display: flex; align-items: center; gap: 10px; }
        .checkbox-group input[type="checkbox"] { width: 20px; height: 20px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-bullhorn"></i> <%= IIf(isEditMode, "编辑", "创建") %>营销活动</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <a href="marketing.asp">营销活动</a> / <span><%= IIf(isEditMode, "编辑活动", "创建活动") %></span>
            </div>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= errorMsg %></div>
        <% End If %>
        
        <div class="form-container">
            <form method="post" action="">
                <%= GetCSRFTokenField() %>
                
                <div class="form-group">
                    <label>活动名称 <span class="required">*</span></label>
                    <input type="text" name="campaign_name" class="form-control" value="<%= SafeOutput(campaignName) %>" required>
                </div>
                
                <div class="form-group">
                    <label>活动描述</label>
                    <textarea name="description" class="form-control" placeholder="请输入活动描述（可选）"><%= SafeOutput(description) %></textarea>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>活动类型 <span class="required">*</span></label>
                        <select name="campaign_type" class="form-control" required>
                            <option value="discount" <%= IIf(campaignType="discount", "selected", "") %>>折扣</option>
                            <option value="coupon" <%= IIf(campaignType="coupon", "selected", "") %>>满减</option>
                            <option value="gift" <%= IIf(campaignType="gift", "selected", "") %>>赠品</option>
                            <option value="other" <%= IIf(campaignType="other", "selected", "") %>>其他</option>
                        </select>
                    </div>
                    
                    <div class="form-group">
                        <label>活动状态</label>
                        <div class="checkbox-group">
                            <input type="checkbox" name="is_active" value="1" <%= IIf(isActive, "checked", "") %>>
                            <span>启用活动</span>
                        </div>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>开始日期 <span class="required">*</span></label>
                        <input type="date" name="start_date" class="form-control" value="<%= FormatDateField(startDate) %>" required>
                    </div>
                    
                    <div class="form-group">
                        <label>结束日期 <span class="required">*</span></label>
                        <input type="date" name="end_date" class="form-control" value="<%= FormatDateField(endDate) %>" required>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>优惠值</label>
                        <input type="number" name="discount_value" class="form-control" value="<%= SafeOutput(discountValue) %>" step="0.01" placeholder="折扣率或优惠金额">
                    </div>
                    
                    <div class="form-group">
                        <label>最低消费金额</label>
                        <input type="number" name="min_purchase" class="form-control" value="<%= SafeOutput(minPurchase) %>" step="0.01" placeholder="0表示无限制">
                    </div>
                </div>
                
                <div class="form-actions">
                    <button type="submit" class="admin-btn admin-btn-primary">
                        <i class="fas fa-save"></i> <%= IIf(isEditMode, "保存修改", "创建活动") %>
                    </button>
                    <a href="marketing.asp" class="admin-btn admin-btn-secondary">
                        <i class="fas fa-times"></i> 取消
                    </a>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
