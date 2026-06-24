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

' ========== 安全工具函数 ==========
Function SafeCount(val)
    If IsNull(val) Or val = "" Then
        SafeCount = 0
    ElseIf IsNumeric(val) Then
        SafeCount = CLng(val)
    Else
        SafeCount = 0
    End If
End Function

' ========== JS字符串安全转义函数 ==========
Function SafeJSString(str)
    If IsNull(str) Then
        SafeJSString = ""
        Exit Function
    End If
    Dim s
    s = CStr(str)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, "'", "\'")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    SafeJSString = s
End Function

' ========== 格式化日期函数 ==========
Function FormatDateTimeVal(val)
    If IsNull(val) Or val = "" Then
        FormatDateTimeVal = ""
    Else
        On Error Resume Next
        FormatDateTimeVal = FormatDateTime(val, 2)
        On Error GoTo 0
    End If
End Function

' ========== SafeNum 函数 ==========
Function SafeNum(val)
    On Error Resume Next
    If IsNull(val) Or val = "" Then
        SafeNum = 0
    ElseIf IsNumeric(val) Then
        SafeNum = CDbl(val)
    Else
        SafeNum = 0
    End If
    On Error GoTo 0
End Function

' ========== IIF 函数 ==========
Function IIF(cond, tVal, fVal)
    If cond Then IIF = tVal Else IIF = fVal
End Function

' ========== 处理审核操作 ==========
Dim action, productId, status, message
action = Request.Form("action")
productId = Request.Form("productId")
status = Request.Form("status")
message = ""

If action = "updateStatus" And IsNumeric(productId) Then
    Dim updateSql
    updateSql = "UPDATE Products SET ReviewStatus = '" & SafeSQL(status) & "', IsActive = " & IIF(status = "Approved", 1, 0) & " WHERE ProductID = " & CLng(productId)
    If ExecuteNonQuery(updateSql) Then
        message = "审核状态已更新！"
    Else
        message = "更新失败：" & Session("LastDBError")
    End If
End If

' ========== 筛选参数 ==========
Dim filterStatus
filterStatus = Request.QueryString("filter")

' ========== 统计数量 ==========
Dim totalCount, pendingCount, approvedCount, rejectedCount
totalCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType = 'KOL'"))
pendingCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType = 'KOL' AND ReviewStatus = 'Pending'"))
approvedCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType = 'KOL' AND ReviewStatus = 'Approved'"))
rejectedCount = SafeCount(GetScalar("SELECT COUNT(*) FROM Products WHERE ProductType = 'KOL' AND ReviewStatus = 'Rejected'"))

' ========== 查询KOL产品 ==========
Dim sqlWhere, rsReviews
sqlWhere = "WHERE p.ProductType = 'KOL'"
If filterStatus = "Pending" Or filterStatus = "Approved" Or filterStatus = "Rejected" Then
    sqlWhere = sqlWhere & " AND p.ReviewStatus = '" & SafeSQL(filterStatus) & "'"
End If

Set rsReviews = ExecuteQuery("SELECT p.*, u.Username FROM Products p LEFT JOIN Users u ON p.KOLID = u.UserID " & sqlWhere & " ORDER BY p.CreatedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>KOL产品审核 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
            <link rel="stylesheet" href="/css/design-tokens.css">
            <link rel="stylesheet" href="/css/buttons.css">
    <style>
        body {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            color: #e0e0e0;
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
        }
        
        .main-content {
            margin-left: 250px;
            padding: 30px;
            min-height: 100vh;
        }
        
        /* 页面头部 */
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .page-title {
            font-size: 24px;
            color: #fff;
            margin: 0;
        }
        .page-title i {
            color: #00bcd4;
            margin-right: 10px;
        }
        .breadcrumb {
            font-size: 14px;
            color: #888;
        }
        .breadcrumb a {
            color: #00bcd4;
            text-decoration: none;
        }
        .breadcrumb a:hover {
            text-decoration: underline;
        }
        
        /* 统计卡片 */
        .stats-section {
            margin-bottom: 25px;
        }
        .stats-cards {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
        }
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            cursor: pointer;
        }
        .stat-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(0,188,212,0.2);
        }
        .stat-card.active {
            border-color: #00bcd4;
            box-shadow: 0 0 15px rgba(0,188,212,0.3);
        }
        .stat-value {
            font-size: 32px;
            font-weight: 700;
            color: #fff;
        }
        .stat-label {
            font-size: 12px;
            color: #888;
            margin-top: 5px;
            text-transform: uppercase;
        }
        .stat-card.total .stat-value { color: #00bcd4; }
        .stat-card.pending .stat-value { color: #ffc107; }
        .stat-card.approved .stat-value { color: #4caf50; }
        .stat-card.rejected .stat-value { color: #f44336; }
        
        /* 提示消息 */
        .alert {
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .alert-success {
            background: rgba(76,175,80,0.1);
            border: 1px solid rgba(76,175,80,0.3);
            color: #4caf50;
        }
        .alert-error {
            background: rgba(244,67,54,0.1);
            border: 1px solid rgba(244,67,54,0.3);
            color: #f44336;
        }
        
        /* 产品卡片网格 */
        .product-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        }
        .product-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            border: 1px solid rgba(255,255,255,0.05);
            overflow: hidden;
            transition: all 0.3s ease;
        }
        .product-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.4);
            border-color: rgba(0,188,212,0.2);
        }
        .product-image {
            width: 100%;
            aspect-ratio: 1 / 1;
            object-fit: cover;
            background: #1a1a2e;
            display: block;
            border-radius: 8px;
        }
        .product-body {
            padding: 20px;
        }
        .product-name {
            font-size: 16px;
            font-weight: 600;
            color: #e8e8e8;
            margin-bottom: 8px;
            line-height: 1.4;
        }
        .product-kol {
            font-size: 13px;
            color: #00bcd4;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .product-kol i {
            font-size: 12px;
        }
        .product-desc {
            font-size: 12px;
            color: #999;
            margin-bottom: 12px;
            line-height: 1.5;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }
        
        /* 配方详情 */
        .recipe-section {
            background: rgba(0,0,0,0.2);
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 12px;
        }
        .recipe-title {
            font-size: 11px;
            color: #888;
            text-transform: uppercase;
            margin-bottom: 8px;
            letter-spacing: 0.5px;
        }
        .recipe-item {
            font-size: 12px;
            color: #c0c0c0;
            padding: 3px 0;
            border-bottom: 1px solid rgba(255,255,255,0.03);
        }
        .recipe-item:last-child {
            border-bottom: none;
        }
        .recipe-item .note-type {
            font-weight: 500;
        }
        .recipe-item .note-type.top { color: #ffc107; }
        .recipe-item .note-type.middle { color: #ce93d8; }
        .recipe-item .note-type.base { color: #80cbc4; }
        .recipe-empty {
            font-size: 12px;
            color: #888;
            font-style: italic;
        }
        
        /* 状态与操作 */
        .product-footer {
            padding: 15px 20px;
            border-top: 1px solid rgba(255,255,255,0.05);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .status-badge {
            display: inline-block;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 500;
        }
        .status-pending {
            background: rgba(255,193,7,0.2);
            color: #ffc107;
        }
        .status-approved {
            background: rgba(76,175,80,0.2);
            color: #4caf50;
        }
        .status-rejected {
            background: rgba(244,67,54,0.2);
            color: #f44336;
        }
        
        /* 按钮 */
        .action-btns {
            display: flex;
            gap: 8px;
        }
        /* .admin-btn 样式已由 /css/buttons.css Section 2 & 4 & 5 统一管理 */
        .no-action {
            font-size: 12px;
            color: #888;
        }
        
        /* 空状态 */
        .empty-state {
            text-align: center;
            padding: 80px 20px;
            color: #888;
            grid-column: 1 / -1;
        }
        .empty-state i {
            font-size: 64px;
            margin-bottom: 20px;
            color: #888;
        }
        .empty-state h3 {
            font-size: 18px;
            margin-bottom: 10px;
            color: #aaa;
        }
        
        /* 通用香调颜色 */
        .note-type { font-weight: 500; }
        .note-type.top { color: #ffc107; }
        .note-type.middle { color: #ce93d8; }
        .note-type.base { color: #80cbc4; }
        
        /* 详情弹窗样式 */
        .detail-section { margin-bottom: 20px; padding: 15px; background: rgba(255,255,255,0.05); border-radius: 8px; }
        .detail-section-title { color: #00bcd4; font-size: 14px; margin-bottom: 12px; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 8px; font-weight: 600; }
        .detail-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .detail-table th { text-align: left; color: #999; font-weight: normal; padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.05); }
        .detail-table td { color: #e0e0e0; padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.05); }
        .detail-table tr:last-child th, .detail-table tr:last-child td { border-bottom: none; }
        .detail-table tr:hover { background: rgba(255,255,255,0.03); }
        .detail-info-row { display: flex; margin-bottom: 10px; }
        .detail-info-label { color: #999; width: 120px; flex-shrink: 0; font-size: 13px; }
        .detail-info-value { color: #e0e0e0; font-size: 13px; }
        .ratio-bar { height: 6px; background: rgba(255,255,255,0.1); border-radius: 3px; margin-top: 4px; overflow: hidden; }
        .ratio-bar-fill { height: 100%; border-radius: 3px; }
        .ratio-bar-fill.top { background: #ffc107; }
        .ratio-bar-fill.middle { background: #ce93d8; }
        .ratio-bar-fill.base { background: #80cbc4; }
        .detail-warning { color: #f44336; font-size: 12px; margin-top: 6px; }
        .detail-badge { display: inline-block; padding: 2px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .detail-badge.pending { background: rgba(255,193,7,0.2); color: #ffc107; }
        .detail-badge.approved { background: rgba(76,175,80,0.2); color: #4caf50; }
        .detail-badge.rejected { background: rgba(244,67,54,0.2); color: #f44336; }
        .detail-empty { color: #888; font-style: italic; font-size: 13px; padding: 8px 0; }
        .detail-image { width: 100px; height: 100px; object-fit: cover; border-radius: 8px; background: #1a1a2e; margin-right: 15px; flex-shrink: 0; }
        .detail-header-flex { display: flex; align-items: flex-start; }
        .admin-btn-info {
            background: rgba(0,188,212,0.15);
            color: #00bcd4;
            border: 1px solid rgba(0,188,212,0.4);
        }
        .admin-btn-info:hover {
            background: rgba(0,188,212,0.25);
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .product-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .main-content { margin-left: 0; padding: 15px; }
            .stats-cards { grid-template-columns: repeat(2, 1fr); }
            .product-grid { grid-template-columns: 1fr; }
        }
        
        /* ===== 弹窗颜色覆盖（覆盖 admin.css 的深蓝色） ===== */
        .admin-modal-content .admin-modal-title,
        .admin-modal-content h1,
        .admin-modal-content h2,
        .admin-modal-content h3,
        .admin-modal-content h4,
        .admin-modal-content h5,
        .admin-modal-content h6 {
            color: #ffffff !important;
        }
        .admin-modal-content .admin-modal-close {
            color: #bbb !important;
        }
        .admin-modal-content .admin-modal-close:hover {
            color: #fff !important;
        }
        .admin-modal-content .admin-form-label,
        .admin-modal-content .form-label,
        .admin-modal-content label {
            color: #e0e0e0 !important;
        }
        .admin-modal-content .admin-form-control,
        .admin-modal-content input,
        .admin-modal-content select,
        .admin-modal-content textarea {
            color: #fff !important;
            background-color: rgba(255,255,255,0.08) !important;
            border-color: rgba(255,255,255,0.15) !important;
        }
        .admin-modal-content input::placeholder,
        .admin-modal-content select::placeholder,
        .admin-modal-content textarea::placeholder {
            color: #999 !important;
        }
        .admin-modal-content small,
        .admin-modal-content .form-text,
        .admin-modal-content .text-muted {
            color: #aaa !important;
        }
        .admin-modal-content .admin-modal-footer {
            background: rgba(0,0,0,0.2) !important;
            border-top-color: rgba(255,255,255,0.1) !important;
        }
        .admin-modal-content .detail-section-title {
            color: #00bcd4 !important;
        }
        .admin-modal-content .detail-info-label,
        .admin-modal-content .detail-table th {
            color: #999 !important;
        }
        .admin-modal-content .detail-info-value,
        .admin-modal-content .detail-table td {
            color: #e0e0e0 !important;
        }
        .admin-modal-content .detail-warning {
            color: #f44336 !important;
        }
        .admin-modal-content .detail-empty {
            color: #888 !important;
        }

        /* 弹窗基础样式补充 */
        .admin-modal-content {
            background: linear-gradient(135deg, #1e1e2e 0%, #2d2d44 100%) !important;
            border: 1px solid rgba(255,255,255,0.1) !important;
            max-width: 800px;
            max-height: 90vh;
            overflow-y: auto;
        }
        .admin-modal-footer {
            display: flex !important;
            justify-content: space-between !important;
            align-items: center !important;
            padding: 15px 20px !important;
            border-top: 1px solid rgba(255,255,255,0.1) !important;
            background: rgba(0,0,0,0.2) !important;
        }
        .admin-modal-footer .admin-btn-success {
            background: linear-gradient(135deg, #4caf50, #45a049) !important;
            color: #fff !important;
            border: none !important;
            padding: 8px 20px !important;
            border-radius: 6px !important;
            cursor: pointer !important;
        }
        .admin-modal-footer .admin-btn-success:hover {
            background: linear-gradient(135deg, #45a049, #3d8b40) !important;
        }
        .admin-modal-footer .admin-btn-danger {
            background: linear-gradient(135deg, #f44336, #d32f2f) !important;
            color: #fff !important;
            border: none !important;
            padding: 8px 20px !important;
            border-radius: 6px !important;
            cursor: pointer !important;
        }
        .admin-modal-footer .admin-btn-danger:hover {
            background: linear-gradient(135deg, #d32f2f, #b71c1c) !important;
        }
        @media (max-width: 768px) {
            .admin-modal-content { max-width: 95% !important; margin: 10px !important; }
        }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2 class="page-title"><i class="fas fa-user-check"></i> KOL产品审核</h2>
                <div class="breadcrumb">
                    <a href="index.asp">技术中心</a> / <span>KOL产品审核</span>
                </div>
            </div>
        </div>
        
        <% If message <> "" Then %>
        <div class="alert alert-success">
            <i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(message) %>
        </div>
        <% End If %>
        
        <!-- 统计仪表板 -->
        <div class="stats-section">
            <div class="stats-cards">
                <div class="stat-card total <%= IIF(filterStatus = "", "active", "") %>" onclick="location.href='kol_reviews.asp'">
                    <div class="stat-value"><%= totalCount %></div>
                    <div class="stat-label">总数</div>
                </div>
                <div class="stat-card pending <%= IIF(filterStatus = "Pending", "active", "") %>" onclick="location.href='kol_reviews.asp?filter=Pending'">
                    <div class="stat-value"><%= pendingCount %></div>
                    <div class="stat-label">待审核</div>
                </div>
                <div class="stat-card approved <%= IIF(filterStatus = "Approved", "active", "") %>" onclick="location.href='kol_reviews.asp?filter=Approved'">
                    <div class="stat-value"><%= approvedCount %></div>
                    <div class="stat-label">已通过</div>
                </div>
                <div class="stat-card rejected <%= IIF(filterStatus = "Rejected", "active", "") %>" onclick="location.href='kol_reviews.asp?filter=Rejected'">
                    <div class="stat-value"><%= rejectedCount %></div>
                    <div class="stat-label">已驳回</div>
                </div>
            </div>
        </div>
        
        <!-- 产品卡片网格 -->
        <div class="product-grid">
            <script>var productDetails = {};</script>
            <%
            If Not rsReviews Is Nothing Then
                If rsReviews.EOF Then
            %>
            <div class="empty-state">
                <i class="fas fa-clipboard-check"></i>
                <h3>暂无KOL产品</h3>
                <p>当前筛选条件下没有需要显示的产品</p>
            </div>
            <%
                Else
                    Do While Not rsReviews.EOF
                        Dim pId, curStatus, sClass
                        pId = rsReviews("ProductID")
                        
                        curStatus = ""
                        On Error Resume Next
                        curStatus = rsReviews("ReviewStatus") & ""
                        On Error GoTo 0
                        
                        sClass = ""
                        Dim statusText
                        Select Case curStatus
                            Case "Pending"
                                sClass = "status-pending"
                                statusText = "待审核"
                            Case "Approved"
                                sClass = "status-approved"
                                statusText = "已通过"
                            Case "Rejected"
                                sClass = "status-rejected"
                                statusText = "已驳回"
                            Case Else
                                sClass = "status-pending"
                                statusText = "待审核"
                        End Select
            %>
            <%
            ' 查询KOL邮箱
            Dim kolEmail
            kolEmail = ""
            On Error Resume Next
            Dim rsKOL
            Set rsKOL = ExecuteQuery("SELECT Email FROM Users WHERE UserID = " & rsReviews("KOLID"))
            If Not rsKOL Is Nothing Then
                If Not rsKOL.EOF Then
                    kolEmail = rsKOL("Email") & ""
                End If
                rsKOL.Close
                Set rsKOL = Nothing
            End If
            On Error GoTo 0
            %>
            <script>
            productDetails[<%=pId%>] = {
                basicInfo: {
                    productName: "<%=SafeJSString(rsReviews("ProductName")&"")%>",
                    description: "<%=SafeJSString(rsReviews("Description")&"")%>",
                    basePrice: <%=SafeNum(rsReviews("BasePrice"))%>,
                    imageURL: "<%=SafeJSString(rsReviews("ImageURL")&"")%>",
                    kolName: "<%=SafeJSString(rsReviews("Username")&"")%>",
                    kolEmail: "<%=SafeJSString(kolEmail)%>",
                    engravable: <%=IIF(Not IsNull(rsReviews("Engravable")) And rsReviews("Engravable") <> 0, "true", "false")%>,
                    engravingPrice: <%=SafeNum(rsReviews("EngravingPrice"))%>,
                    isActive: <%=IIF(Not IsNull(rsReviews("IsActive")) And rsReviews("IsActive") <> 0, "true", "false")%>,
                    createdAt: "<%=SafeJSString(FormatDateTimeVal(rsReviews("CreatedAt")))%>",
                    reviewStatus: "<%=SafeJSString(curStatus)%>"
                },
                noteRatios: [
                    <%
                    Dim rsDetailRatios, ratioIdx
                    ratioIdx = 0
                    On Error Resume Next
                    Set rsDetailRatios = ExecuteQuery("SELECT pnr.Percentage, fn.NoteName, fn.NoteType FROM (ProductNoteRatios AS pnr LEFT JOIN FragranceNotes AS fn ON pnr.NoteID = fn.NoteID) WHERE pnr.ProductID = " & pId & " ORDER BY fn.NoteType, pnr.Percentage DESC")
                    On Error GoTo 0
                    If Not rsDetailRatios Is Nothing Then
                        Do While Not rsDetailRatios.EOF
                            If ratioIdx > 0 Then Response.Write ","
                    %>
                    { noteName: "<%=SafeJSString(rsDetailRatios("NoteName")&"")%>", noteType: "<%=SafeJSString(rsDetailRatios("NoteType")&"")%>", percentage: <%=SafeNum(rsDetailRatios("Percentage"))%> }
                    <%
                            ratioIdx = ratioIdx + 1
                            rsDetailRatios.MoveNext
                        Loop
                        rsDetailRatios.Close
                        Set rsDetailRatios = Nothing
                    End If
                    %>
                ],
                volumes: [
                    <%
                    Dim rsDetailVolumes, volIdx
                    volIdx = 0
                    On Error Resume Next
                    Set rsDetailVolumes = ExecuteQuery("SELECT pvp.VolumeID, pvp.Price, v.VolumeML, v.VolumeName, v.PriceMultiplier, v.IsActive FROM ProductVolumePrices pvp INNER JOIN Volumes v ON pvp.VolumeID = v.VolumeID WHERE pvp.ProductID = " & CLng(pId) & " AND v.IsActive <> 0 ORDER BY v.VolumeML")
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                    On Error GoTo 0
                    If Not rsDetailVolumes Is Nothing Then
                        Do While Not rsDetailVolumes.EOF
                            If volIdx > 0 Then Response.Write ","
                    %>
                    { volumeName: "<%=SafeJSString(rsDetailVolumes("VolumeName")&"")%>", volumeML: <%=SafeNum(rsDetailVolumes("VolumeML"))%>, price: <%=SafeNum(rsDetailVolumes("Price"))%>, priceMultiplier: <%=SafeNum(rsDetailVolumes("PriceMultiplier"))%> }
                    <%
                            volIdx = volIdx + 1
                            rsDetailVolumes.MoveNext
                        Loop
                        rsDetailVolumes.Close
                        Set rsDetailVolumes = Nothing
                    End If
                    %>
                ],
                bottles: [
                    <%
                    Dim rsDetailBottles, bottleIdx
                    bottleIdx = 0
                    On Error Resume Next
                    Set rsDetailBottles = ExecuteQuery("SELECT bs.BottleName AS StyleName, bs.Description AS StyleDesc, bs.PriceAddition AS Surcharge FROM ProductBottleStyles AS pbs INNER JOIN BottleStyles AS bs ON pbs.BottleID = bs.BottleID WHERE pbs.ProductID = " & CLng(pId) & " AND bs.IsActive <> 0")
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                    On Error GoTo 0
                    If Not rsDetailBottles Is Nothing Then
                        Do While Not rsDetailBottles.EOF
                            If bottleIdx > 0 Then Response.Write ","
                    %>
                    { styleName: "<%=SafeJSString(rsDetailBottles("StyleName")&"")%>", styleDesc: "<%=SafeJSString(rsDetailBottles("StyleDesc")&"")%>", surcharge: <%=SafeNum(rsDetailBottles("Surcharge"))%> }
                    <%
                            bottleIdx = bottleIdx + 1
                            rsDetailBottles.MoveNext
                        Loop
                        rsDetailBottles.Close
                        Set rsDetailBottles = Nothing
                    End If
                    %>
                ],
                recipe: {
                    <%
                    Dim hasRecipe
                    hasRecipe = False
                    If Not IsNull(rsReviews("RecipeID")) And (rsReviews("RecipeID")&"") <> "" And IsNumeric(rsReviews("RecipeID")) Then
                        If CLng(rsReviews("RecipeID")) > 0 Then
                            Dim rsDetailRecipe
                            On Error Resume Next
                            Set rsDetailRecipe = ExecuteQuery("SELECT RecipeName, RecipeCode, Description, ProductType, ReviewStatus FROM Recipes WHERE RecipeID = " & CLng(rsReviews("RecipeID")))
                            If Err.Number <> 0 Then
                                Err.Clear
                            End If
                            On Error GoTo 0
                            If Not rsDetailRecipe Is Nothing Then
                                If Not rsDetailRecipe.EOF Then
                                    hasRecipe = True
                    %>
                    recipeName: "<%=SafeJSString(rsDetailRecipe("RecipeName")&"")%>",
                    recipeCode: "<%=SafeJSString(rsDetailRecipe("RecipeCode")&"")%>",
                    recipeDesc: "<%=SafeJSString(rsDetailRecipe("Description")&"")%>",
                    productType: "<%=SafeJSString(rsDetailRecipe("ProductType")&"")%>",
                    reviewStatus: "<%=SafeJSString(rsDetailRecipe("ReviewStatus")&"")%>",
                    hasRecipe: true
                    <%
                                End If
                                rsDetailRecipe.Close
                                Set rsDetailRecipe = Nothing
                            End If
                        End If
                    End If
                    If Not hasRecipe Then
                    %>
                    hasRecipe: false
                    <%
                    End If
                    %>
                }
            };
            </script>
            <div class="product-card">
                <img src="<%= HTMLEncode(rsReviews("ImageURL") & "") %>" class="product-image" alt="产品图片" onerror="this.src='/images/placeholder.jpg'">
                <div class="product-body">
                    <div class="product-name"><%= HTMLEncode(rsReviews("ProductName") & "") %></div>
                    <div class="product-kol">
                        <i class="fas fa-user"></i> <%= HTMLEncode(rsReviews("Username") & "") %>
                    </div>
                    <div class="product-desc">
                        <%= Left(HTMLEncode(rsReviews("Description") & ""), 60) %><%
                        If Len(rsReviews("Description") & "") > 60 Then Response.Write "..."
                        %>
                    </div>
                    
                    <div class="recipe-section">
                        <div class="recipe-title"><i class="fas fa-flask"></i> 配方详情</div>
                        <%
                        Dim rsRatios
                        Set rsRatios = ExecuteQuery("SELECT r.*, n.NoteName FROM ProductNoteRatios r LEFT JOIN FragranceNotes n ON r.NoteID = n.NoteID WHERE r.ProductID = " & pId)
                        If Not rsRatios Is Nothing Then
                            If rsRatios.EOF Then
                        %>
                        <div class="recipe-empty">暂无配方数据</div>
                        <%
                            Else
                                Do While Not rsRatios.EOF
                                    Dim noteTypeDisplay, noteTypeClass
                                    noteTypeDisplay = rsRatios("NoteType") & ""
                                    noteTypeClass = ""
                                    If noteTypeDisplay = "前调" Then
                                        noteTypeClass = "top"
                                    ElseIf noteTypeDisplay = "中调" Then
                                        noteTypeClass = "middle"
                                    ElseIf noteTypeDisplay = "后调" Then
                                        noteTypeClass = "base"
                                    End If
                        %>
                        <div class="recipe-item">
                            <span class="note-type <%= noteTypeClass %>">[<%= HTMLEncode(noteTypeDisplay) %>]</span>
                            <%= HTMLEncode(rsRatios("NoteName") & "") %>
                            <span style="color:#888;">(<%= rsRatios("Percentage") %>%)</span>
                        </div>
                        <%
                                    rsRatios.MoveNext
                                Loop
                            End If
                            rsRatios.Close
                            Set rsRatios = Nothing
                        End If
                        %>
                    </div>
                </div>
                
                <div class="product-footer">
                    <span class="status-badge <%= sClass %>"><%= statusText %></span>
                    <div class="action-btns">
                        <button type="button" class="admin-btn admin-btn-info" onclick="openDetailModal(<%=pId%>)"><i class="fas fa-eye"></i> 查看详情</button>
                        <% If curStatus = "Pending" Then %>
                        <form method="post" style="display:inline;">
                            <input type="hidden" name="action" value="updateStatus">
                            <input type="hidden" name="productId" value="<%= pId %>">
                            <input type="hidden" name="status" value="Approved">
                            <button type="submit" class="admin-btn admin-btn-success"><i class="fas fa-check"></i> 通过</button>
                        </form>
                        <form method="post" style="display:inline;">
                            <input type="hidden" name="action" value="updateStatus">
                            <input type="hidden" name="productId" value="<%= pId %>">
                            <input type="hidden" name="status" value="Rejected">
                            <button type="submit" class="admin-btn admin-btn-danger"><i class="fas fa-times"></i> 驳回</button>
                        </form>
                        <% Else %>
                        <span class="no-action">无需操作</span>
                        <% End If %>
                    </div>
                </div>
            </div>
            <%
                        rsReviews.MoveNext
                    Loop
                End If
                rsReviews.Close
                Set rsReviews = Nothing
            End If
            %>
        </div>
    </div>
    
    <!-- 产品详情弹窗 -->
    <div class="admin-modal" id="detailModal">
        <div class="admin-modal-content" style="max-width:800px;">
            <div class="admin-modal-header">
                <h3 class="admin-modal-title"><i class="fas fa-clipboard-list"></i> 产品审核详情</h3>
                <span class="admin-modal-close" onclick="closeDetailModal()">&times;</span>
            </div>
            <div class="admin-modal-body" id="detailBody">
                <!-- 动态内容 -->
            </div>
            <div class="admin-modal-footer" id="detailFooter">
                <div id="detailReviewActions" style="display:none;">
                    <button class="admin-btn admin-btn-success" onclick="approveFromDetail()">
                        <i class="fas fa-check"></i> 通过
                    </button>
                    <button class="admin-btn admin-btn-danger" onclick="rejectFromDetail()" style="margin-left:8px;">
                        <i class="fas fa-times"></i> 驳回
                    </button>
                </div>
                <button class="admin-btn" onclick="closeDetailModal()">关闭</button>
            </div>
        </div>
    </div>
    
    <script>
    var currentDetailProductId = null;

    function openDetailModal(pid) {
        currentDetailProductId = pid;
        var data = productDetails[pid];
        if (!data) return;
        var html = '';
        var bi = data.basicInfo;
        
        // 区块1：产品基本信息
        html += '<div class="detail-section">';
        html += '<div class="detail-section-title"><i class="fas fa-info-circle"></i> 产品基本信息</div>';
        html += '<div class="detail-header-flex">';
        if (bi.imageURL && bi.imageURL !== '') {
            html += '<img src="' + escapeHtml(bi.imageURL) + '" class="detail-image" alt="产品图片" onerror="this.src=\'/images/placeholder.jpg\'">';
        }
        html += '<div style="flex:1;">';
        html += '<div class="detail-info-row"><span class="detail-info-label">产品名称：</span><span class="detail-info-value">' + escapeHtml(bi.productName) + '</span></div>';
        html += '<div class="detail-info-row"><span class="detail-info-label">描述：</span><span class="detail-info-value">' + escapeHtml(bi.description || '无') + '</span></div>';
        html += '<div class="detail-info-row"><span class="detail-info-label">基础价格：</span><span class="detail-info-value">$' + bi.basePrice.toFixed(2) + '</span></div>';
        html += '<div class="detail-info-row"><span class="detail-info-label">KOL推荐人：</span><span class="detail-info-value">' + escapeHtml(bi.kolName) + (bi.kolEmail ? ' (' + escapeHtml(bi.kolEmail) + ')' : '') + '</span></div>';
        html += '<div class="detail-info-row"><span class="detail-info-label">创建时间：</span><span class="detail-info-value">' + escapeHtml(bi.createdAt) + '</span></div>';
        html += '<div class="detail-info-row"><span class="detail-info-label">审核状态：</span><span class="detail-info-value">' + getStatusBadge(bi.reviewStatus) + '</span></div>';
        html += '</div></div></div>';
        
        // 区块2：香调配比
        html += '<div class="detail-section">';
        html += '<div class="detail-section-title"><i class="fas fa-flask"></i> 香调配比</div>';
        if (data.noteRatios && data.noteRatios.length > 0) {
            var totalRatio = 0;
            for (var i = 0; i < data.noteRatios.length; i++) {
                totalRatio += data.noteRatios[i].percentage;
            }
            html += '<table class="detail-table">';
            html += '<tr><th>香调名称</th><th>类型</th><th style="width:150px;">配比</th></tr>';
            var typeOrder = {'前调':1,'中调':2,'后调':3};
            var sorted = data.noteRatios.slice().sort(function(a,b){
                var ta = typeOrder[a.noteType] || 4;
                var tb = typeOrder[b.noteType] || 4;
                if (ta !== tb) return ta - tb;
                return b.percentage - a.percentage;
            });
            for (var i = 0; i < sorted.length; i++) {
                var nr = sorted[i];
                var typeClass = '';
                if (nr.noteType === '前调') typeClass = 'top';
                else if (nr.noteType === '中调') typeClass = 'middle';
                else if (nr.noteType === '后调') typeClass = 'base';
                html += '<tr>';
                html += '<td>' + escapeHtml(nr.noteName) + '</td>';
                html += '<td><span class="note-type ' + typeClass + '">' + escapeHtml(nr.noteType) + '</span></td>';
                html += '<td>';
                html += '<span style="font-weight:500;">' + nr.percentage + '%</span>';
                html += '<div class="ratio-bar"><div class="ratio-bar-fill ' + typeClass + '" style="width:' + nr.percentage + '%;"></div></div>';
                html += '</td>';
                html += '</tr>';
            }
            html += '</table>';
            if (Math.abs(totalRatio - 100) > 0.5) {
                html += '<div class="detail-warning"><i class="fas fa-exclamation-triangle"></i> 配比总和为 ' + totalRatio.toFixed(1) + '%，不等于100%</div>';
            } else {
                html += '<div style="color:#4caf50;font-size:12px;margin-top:6px;"><i class="fas fa-check-circle"></i> 配比总和：' + totalRatio.toFixed(1) + '%</div>';
            }
        } else {
            html += '<div class="detail-empty">暂无香调配比数据</div>';
        }
        html += '</div>';
        
        // 区块3：容量与定价
        html += '<div class="detail-section">';
        html += '<div class="detail-section-title"><i class="fas fa-bottle-droplet"></i> 容量与定价</div>';
        if (data.volumes && data.volumes.length > 0) {
            html += '<table class="detail-table">';
            html += '<tr><th>容量</th><th>价格</th></tr>';
            for (var i = 0; i < data.volumes.length; i++) {
                var vol = data.volumes[i];
                var displayPrice = vol.price;
                if (displayPrice <= 0 && bi.basePrice > 0 && vol.priceMultiplier > 0) {
                    displayPrice = bi.basePrice * vol.priceMultiplier;
                }
                html += '<tr><td>' + escapeHtml(vol.volumeName || (vol.volumeML + 'ml')) + '</td><td>$' + displayPrice.toFixed(2) + '</td></tr>';
            }
            html += '</table>';
        } else {
            html += '<div class="detail-empty">未设置容量价格</div>';
        }
        html += '</div>';
        
        // 区块4：瓶型选择
        html += '<div class="detail-section">';
        html += '<div class="detail-section-title"><i class="fas fa-wine-bottle"></i> 瓶型选择</div>';
        if (data.bottles && data.bottles.length > 0) {
            html += '<table class="detail-table">';
            html += '<tr><th>瓶型名称</th><th>描述</th><th>附加费用</th></tr>';
            for (var i = 0; i < data.bottles.length; i++) {
                var bot = data.bottles[i];
                html += '<tr><td>' + escapeHtml(bot.styleName) + '</td><td>' + escapeHtml(bot.styleDesc || '无') + '</td><td>$' + bot.surcharge.toFixed(2) + '</td></tr>';
            }
            html += '</table>';
        } else {
            html += '<div class="detail-empty">未设置瓶型</div>';
        }
        html += '</div>';
        
        // 区块5：关联配方
        html += '<div class="detail-section">';
        html += '<div class="detail-section-title"><i class="fas fa-book"></i> 关联配方</div>';
        if (data.recipe && data.recipe.hasRecipe) {
            html += '<div class="detail-info-row"><span class="detail-info-label">配方名称：</span><span class="detail-info-value">' + escapeHtml(data.recipe.recipeName) + '</span></div>';
            html += '<div class="detail-info-row"><span class="detail-info-label">配方编码：</span><span class="detail-info-value">' + escapeHtml(data.recipe.recipeCode) + '</span></div>';
            html += '<div class="detail-info-row"><span class="detail-info-label">配方描述：</span><span class="detail-info-value">' + escapeHtml(data.recipe.recipeDesc || '无') + '</span></div>';
            html += '<div class="detail-info-row"><span class="detail-info-label">产品类型：</span><span class="detail-info-value">' + escapeHtml(data.recipe.productType) + '</span></div>';
            html += '<div class="detail-info-row"><span class="detail-info-label">配方审核状态：</span><span class="detail-info-value">' + getStatusBadge(data.recipe.reviewStatus) + '</span></div>';
        } else {
            html += '<div class="detail-empty">未关联配方</div>';
        }
        html += '</div>';
        
        // 区块6：其他设置
        html += '<div class="detail-section">';
        html += '<div class="detail-section-title"><i class="fas fa-cog"></i> 其他设置</div>';
        html += '<div class="detail-info-row"><span class="detail-info-label">支持刻字：</span><span class="detail-info-value">' + (bi.engravable ? '是' : '否') + '</span></div>';
        if (bi.engravable) {
            html += '<div class="detail-info-row"><span class="detail-info-label">刻字附加价格：</span><span class="detail-info-value">$' + bi.engravingPrice.toFixed(2) + '</span></div>';
        }
        html += '<div class="detail-info-row"><span class="detail-info-label">产品激活：</span><span class="detail-info-value">' + (bi.isActive ? '已激活' : '未激活') + '</span></div>';
        html += '</div>';
        
        document.getElementById('detailBody').innerHTML = html;

        // 控制审核按钮显示
        var reviewActions = document.getElementById('detailReviewActions');
        if (bi.reviewStatus === 'Pending') {
            reviewActions.style.display = 'block';
        } else {
            reviewActions.style.display = 'none';
        }

        document.getElementById('detailModal').style.display = 'block';
        document.body.style.overflow = 'hidden';
    }
    
    function closeDetailModal() {
        document.getElementById('detailModal').style.display = 'none';
        document.body.style.overflow = '';
    }

    function approveFromDetail() {
        if (!currentDetailProductId) return;
        if (!confirm('确定通过该产品的审核吗？')) return;
        var form = document.createElement('form');
        form.method = 'POST';
        form.action = '';

        var fields = {
            'action': 'updateStatus',
            'productId': currentDetailProductId,
            'status': 'Approved'
        };
        for (var key in fields) {
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = key;
            input.value = fields[key];
            form.appendChild(input);
        }
        document.body.appendChild(form);
        form.submit();
    }

    function rejectFromDetail() {
        if (!currentDetailProductId) return;
        if (!confirm('确定驳回该产品吗？')) return;
        var form = document.createElement('form');
        form.method = 'POST';
        form.action = '';

        var fields = {
            'action': 'updateStatus',
            'productId': currentDetailProductId,
            'status': 'Rejected'
        };
        for (var key in fields) {
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = key;
            input.value = fields[key];
            form.appendChild(input);
        }
        document.body.appendChild(form);
        form.submit();
    }

    function escapeHtml(text) {
        if (text === null || text === undefined) return '';
        var div = document.createElement('div');
        div.appendChild(document.createTextNode(String(text)));
        return div.innerHTML;
    }
    
    function getStatusBadge(status) {
        if (status === 'Approved') return '<span class="detail-badge approved">已通过</span>';
        if (status === 'Rejected') return '<span class="detail-badge rejected">已驳回</span>';
        return '<span class="detail-badge pending">待审核</span>';
    }
    
    // 点击弹窗外部关闭
    document.getElementById('detailModal').addEventListener('click', function(e) {
        if (e.target === this) closeDetailModal();
    });
    
    // ESC键关闭
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && document.getElementById('detailModal').style.display === 'block') {
            closeDetailModal();
        }
    });
    </script>
</body>
</html>
<% Call CloseConnection() %>
