<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/dal.asp"-->
<!--#include file="../../includes/audit_utils.asp"-->
<%
Call OpenConnection()
Call EnsureAuditLogTable()

Dim action, msg, msgType
action = Request.QueryString("action")
msg = Request.QueryString("msg")
If msg <> "" Then msgType = Request.QueryString("type") Else msgType = "" End If

' 处理删除图片操作
If Request.ServerVariables("REQUEST_METHOD") = "POST" And Request.Form("action") = "delete_image" Then
    If Not ValidateCSRFToken() Then
        Response.Write "<script>alert('安全验证失败'); history.back();</script>"
        Response.End
    End If
    
    Dim imgId : imgId = Request.Form("img_id")
    Dim imgPath : imgPath = Trim(Request.Form("img_path"))
    
    If IsNumeric(imgId) And imgId <> "" Then
        Dim delParams(0)
        delParams(0) = Array("@ImageID", DAL_adInteger, 0, CLng(imgId))
        
        ' 删除数据库记录
        If DAL_Execute("DELETE FROM ProductImages WHERE ImageID=@ImageID", delParams) >= 0 Then
            ' 尝试删除物理文件
            If imgPath <> "" Then
                Dim fs : Set fs = CreateObject("Scripting.FileSystemObject")
                Dim fullPath : fullPath = Server.MapPath(imgPath)
                If fs.FileExists(fullPath) Then
                    fs.DeleteFile fullPath, True
                End If
                Set fs = Nothing
            End If
            Call AuditLog(AUDIT_ACTION_DELETE, AUDIT_TARGET_PRODUCT, CLng(imgId), "图片删除", "已删除图片记录 (ID: " & imgId & ") 路径: " & imgPath)
            msg = "图片已删除"
            msgType = "success"
        Else
            msg = "图片删除失败"
            msgType = "error"
        End If
    End If
End If

' 获取统计
Dim totalImages, orphanImages, totalSize
totalImages = DAL_GetScalar("SELECT COUNT(*) FROM ProductImages", Null, 0)
totalSize = DAL_GetScalar("SELECT ISNULL(SUM(ImageSize), 0) FROM ProductImages", Null, 0)
Dim orphanColor
orphanImages = DAL_GetScalar("SELECT COUNT(*) FROM ProductImages pi LEFT JOIN Products p ON pi.ProductID = p.ProductID WHERE p.ProductID IS NULL", Null, 0)
If orphanImages > 0 Then orphanColor = "#ff6b6b" Else orphanColor = "#81c784"

' 分页参数
Dim page, pageSize, pageInfo
page = CInt(Request.QueryString("page"))
If page < 1 Then page = 1
pageSize = 20

' 图片列表（含分页）
Dim rs, imgSql, imgParams
imgSql = "SELECT pi.*, p.ProductName FROM ProductImages pi LEFT JOIN Products p ON pi.ProductID = p.ProductID ORDER BY pi.ImageID DESC"
Set rs = DAL_GetListPaged(imgSql, Null, page, pageSize, pageInfo)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>产品图片管理 - 系统管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; }
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 25px; border-radius: 12px; text-align: center; }
        .stat-value { font-size: 28px; font-weight: bold; color: #fff; }
        .stat-label { color: #888; font-size: 14px; margin-top: 5px; }
        .img-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 15px; }
        .img-item { background: #2d2d44; border-radius: 8px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .img-item img { width: 100%; height: 140px; object-fit: cover; background: #1a1a2e; }
        .img-info { padding: 10px; }
        .img-info .product-name { font-size: 12px; color: #aaa; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .img-info .img-actions { margin-top: 8px; }
        .img-info .img-actions a { font-size: 12px; color: #ff6b6b; text-decoration: none; }
        .img-info .img-actions a:hover { text-decoration: underline; }
        .alert { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
        .alert-success { background: rgba(76, 175, 80, 0.15); color: #81c784; border: 1px solid rgba(76, 175, 80, 0.3); }
        .alert-error { background: rgba(244, 67, 54, 0.15); color: #e57373; border: 1px solid rgba(244, 67, 54, 0.3); }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-images"></i> 产品图片管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">系统中心</a> / <span>图片管理</span>
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%= msgType %>">
            <i class="fas <%= IIF(msgType="success","fa-check-circle","fa-exclamation-circle") %>"></i> <%= Server.HTMLEncode(msg) %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value"><%= totalImages %></div>
                <div class="stat-label">图片总数</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= Round(totalSize / 1024, 1) %> KB</div>
                <div class="stat-label">总大小</div>
            </div>
            <div class="stat-card">
                <div class="stat-value" style="color: <%= orphanColor %>"><%= orphanImages %></div>
                <div class="stat-label">孤立图片</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= pageInfo("totalPages") %></div>
                <div class="stat-label">分页数</div>
            </div>
        </div>
        
        <div class="dashboard-card">
            <h3><i class="fas fa-list"></i> 图片列表</h3>
            <div class="img-grid">
                <% If Not rs Is Nothing And Not rs.EOF Then
                    Do While Not rs.EOF
                        Dim imgSrc : imgSrc = rs("ImageURL")
                        If IsNull(imgSrc) Or imgSrc = "" Then imgSrc = "/images/default-product.svg"
                %>
                <div class="img-item">
                    <img src="<%= Server.HTMLEncode(imgSrc) %>" alt="产品图片" loading="lazy">
                    <div class="img-info">
                        <div class="product-name"><i class="fas fa-tag"></i> <%= Server.HTMLEncode(rs("ProductName")) %></div>
                        <div class="product-name">#<%= rs("ImageID") %></div>
                        <div class="img-actions">
                            <form method="post" action="image_management.asp" onsubmit="return confirm('确定删除此图片？')" style="display:inline">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" name="action" value="delete_image">
                                <input type="hidden" name="img_id" value="<%= rs("ImageID") %>">
                                <input type="hidden" name="img_path" value="<%= Server.HTMLEncode(imgSrc) %>">
                                <a href="#" onclick="this.closest('form').submit(); return false;"><i class="fas fa-trash"></i> 删除</a>
                            </form>
                        </div>
                    </div>
                </div>
                <%  rs.MoveNext
                    Loop
                Else %>
                <div style="grid-column: 1/-1; text-align: center; padding: 40px; color: #888;">
                    <i class="fas fa-images" style="font-size: 48px; margin-bottom: 15px; display: block;"></i>
                    暂无图片数据
                </div>
                <% End If %>
            </div>
            
            <% If pageInfo("totalPages") > 1 Then %>
            <div class="pagination" style="margin-top: 20px; text-align: center;">
                <% If pageInfo("hasPrev") Then %>
                <a href="image_management.asp?page=<%= pageInfo("currentPage") - 1 %>" class="btn btn-sm">上一页</a>
                <% End If %>
                <span style="color: #888; padding: 0 15px;">第 <%= pageInfo("currentPage") %> / <%= pageInfo("totalPages") %> 页</span>
                <% If pageInfo("hasNext") Then %>
                <a href="image_management.asp?page=<%= pageInfo("currentPage") + 1 %>" class="btn btn-sm">下一页</a>
                <% End If %>
            </div>
            <% End If %>
        </div>
    </div>
</body>
</html>
<%
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Call CloseConnection()
%>