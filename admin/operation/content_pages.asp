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

Function GetScalar(sql)
    Dim rs, val : val = ""
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                val = rs(0)
                rs.Close
            End If
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing : GetScalar = val
End Function

' 自动创建 ContentPages 表
On Error Resume Next
conn.Execute "SELECT TOP 1 * FROM ContentPages WHERE 1=0"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE ContentPages (" & _
        "PageID INT IDENTITY(1,1) PRIMARY KEY," & _
        "Title NVARCHAR(200) NOT NULL," & _
        "Slug NVARCHAR(200) UNIQUE," & _
        "Content NTEXT," & _
        "MetaDescription NVARCHAR(500)," & _
        "IsPublished BIT DEFAULT 0," & _
        "SortOrder INT DEFAULT 0," & _
        "UpdatedBy INT NULL," & _
        "UpdatedAt DATETIME DEFAULT GETDATE()," & _
        "CreatedAt DATETIME DEFAULT GETDATE()" & _
        ")"
    
    ' 插入默认页面
    conn.Execute "INSERT INTO ContentPages (Title, Slug, Content, IsPublished, SortOrder) VALUES ('关于我们', 'about', '<h2>关于我们</h2><p>欢迎来到香氛定制电商平台...</p>', 1, 1)"
    conn.Execute "INSERT INTO ContentPages (Title, Slug, Content, IsPublished, SortOrder) VALUES ('联系我们', 'contact', '<h2>联系我们</h2><p>如有任何问题请与我们联系...</p>', 1, 2)"
    conn.Execute "INSERT INTO ContentPages (Title, Slug, Content, IsPublished, SortOrder) VALUES ('隐私政策', 'privacy', '<h2>隐私政策</h2><p>我们重视您的隐私...</p>', 1, 3)"
    conn.Execute "INSERT INTO ContentPages (Title, Slug, Content, IsPublished, SortOrder) VALUES ('服务条款', 'terms', '<h2>服务条款</h2><p>使用本平台即表示同意以下条款...</p>', 1, 4)"
    conn.Execute "INSERT INTO ContentPages (Title, Slug, Content, IsPublished, SortOrder) VALUES ('配送说明', 'shipping', '<h2>配送说明</h2><p>全国顺丰包邮，预计3-5个工作日送达...</p>', 1, 5)"
End If
On Error GoTo 0

Dim msg : msg = ""
Dim msgType : msgType = ""
Dim editMode : editMode = False
Dim editPage
Set editPage = Nothing

' 处理 POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action : action = Request.Form("action")
    
    If action = "save" Then
        Dim pageID : pageID = Request.Form("pageID")
        Dim title : title = Request.Form("title")
        Dim slug : slug = Request.Form("slug")
        Dim content : content = Request.Form("content")
        Dim metaDesc : metaDesc = Request.Form("metaDescription")
        Dim isPublished : isPublished = IIf(Request.Form("isPublished")="1",1,0)
        Dim sortOrder : sortOrder = CInt("0" & Request.Form("sortOrder"))
        
        If pageID = "" Or pageID = "0" Then
            ' 新增
            conn.Execute "INSERT INTO ContentPages (Title, Slug, Content, MetaDescription, IsPublished, SortOrder, UpdatedBy) VALUES (" & _
                "'" & Replace(title,"'","''") & "', '" & Replace(slug,"'","''") & "', '" & Replace(content,"'","''") & "', " & _
                "'" & Replace(metaDesc,"'","''") & "', " & isPublished & ", " & sortOrder & ", " & Session("AdminID") & ")"
            msg = "页面已创建"
        Else
            ' 更新
            conn.Execute "UPDATE ContentPages SET Title='" & Replace(title,"'","''") & "', " & _
                "Slug='" & Replace(slug,"'","''") & "', Content='" & Replace(content,"'","''") & "', " & _
                "MetaDescription='" & Replace(metaDesc,"'","''") & "', IsPublished=" & isPublished & ", " & _
                "SortOrder=" & sortOrder & ", UpdatedBy=" & Session("AdminID") & ", UpdatedAt=GETDATE() " & _
                "WHERE PageID=" & CInt(pageID)
            msg = "页面已更新"
        End If
        msgType = "success"
    
    ElseIf action = "delete" Then
        conn.Execute "DELETE FROM ContentPages WHERE PageID=" & CInt(Request.Form("pageID"))
        msg = "页面已删除"
        msgType = "success"
    End If
End If

' 编辑模式
Dim editID : editID = 0
If Request.QueryString("edit") <> "" And Request.QueryString("edit") <> "new" Then
    If IsNumeric(Request.QueryString("edit")) Then
        editMode = True
        editID = CInt(Request.QueryString("edit"))
        Set editPage = ExecuteQuery("SELECT * FROM ContentPages WHERE PageID=" & editID)
        If editPage Is Nothing Then
            editMode = False
        ElseIf editPage.EOF Then
            editPage.Close
            Set editPage = Nothing
            editMode = False
        End If
    End If
End If

' 获取页面列表
Dim rsPages
Set rsPages = ExecuteQuery("SELECT * FROM ContentPages ORDER BY SortOrder ASC, Title ASC")
Dim pagesCountLabel : pagesCountLabel = "0"
If Not rsPages Is Nothing Then
    If Not rsPages.EOF Then pagesCountLabel = ""
End If
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>内容页面管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI',Arial,sans-serif; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .page-header { margin-bottom: 25px; }
        .page-title { color: #fff; font-size: 24px; margin: 0 0 8px; }
        .breadcrumb { color: #888; font-size: 13px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        .msg { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; font-weight: 500; }
        .msg-success { background: rgba(76,175,80,0.15); color: #4CAF50; border: 1px solid rgba(76,175,80,0.3); }
        .card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 15px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 16px; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 12px; background: rgba(0,188,212,0.06); border-bottom: 1px solid rgba(255,255,255,0.06); font-size: 13px; color: #999; }
        td { padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.03); }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .badge-published { background: rgba(76,175,80,0.2); color: #A5D6A7; }
        .badge-draft { background: rgba(158,158,158,0.2); color: #BDBDBD; }
        .slug-text { font-family: 'Consolas',monospace; font-size: 12px; color: #00bcd4; }
        .editor-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        @media (max-width: 1024px) { .editor-row { grid-template-columns: 1fr; } }
        @media (max-width: 768px) { .main-content { margin-left: 0; } }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; color: #999; font-size: 13px; margin-bottom: 5px; }
        .form-group input, .form-group textarea, .form-group select { width: 100%; padding: 10px; border: 1px solid rgba(255,255,255,0.12); border-radius: 8px; background: #1a1a2e; color: #e0e0e0; font-size: 14px; }
        .form-group input:focus, .form-group textarea:focus { border-color: #00bcd4; outline: none; }
        .form-group textarea { min-height: 300px; font-family: 'Consolas',monospace; font-size: 13px; line-height: 1.6; resize: vertical; }
        .form-actions { display: flex; gap: 10px; justify-content: flex-end; }
        .empty { text-align: center; padding: 40px; color: #666; }
        .content-preview { background: #1a1a2e; padding: 15px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.06); max-height: 350px; overflow-y: auto; font-size: 14px; line-height: 1.8; }
        .content-preview h2, .content-preview h3 { color: #00bcd4; }
    </style>
</head>
<body data-theme="operation-dark">
<!--#include file="includes/nav.asp"-->
<div class="main-content">
    <div class="page-header">
        <h2 class="page-title"><i class="fas fa-file-alt"></i> 内容页面管理</h2>
        <div class="breadcrumb"><a href="index.asp">运营中心</a> / 内容页面管理</div>
    </div>

    <% If msg <> "" Then %>
    <div class="msg msg-<%= msgType %>"><i class="fas fa-info-circle"></i> <%= msg %></div>
    <% End If %>

    <% If editMode Then
       If IsObject(editPage) Then
       If Not editPage Is Nothing Then %>
    <% If Not editPage.EOF Then %>
    <!-- 编辑面板 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-edit"></i> 编辑页面</div>
        <div class="card-body">
            <form method="post">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="pageID" value="<%= editPage("PageID") %>">
                <div class="editor-row">
                    <div>
                        <div class="form-group">
                            <label>页面标题</label>
                            <input type="text" name="title" value="<%= editPage("Title") %>" required>
                        </div>
                        <div class="form-group">
                            <label>URL标识 (Slug)</label>
                            <input type="text" name="slug" value="<%= editPage("Slug") %>" required>
                        </div>
                        <div style="display:flex;gap:15px;">
                            <div class="form-group" style="flex:1;">
                                <label>排序</label>
                                <input type="number" name="sortOrder" value="<%= editPage("SortOrder") %>" min="0" style="max-width:100px;">
                            </div>
                            <div class="form-group" style="flex:1;">
                                <label>状态</label>
                                <select name="isPublished">
                                    <option value="1" <%= IIf(CBool(editPage("IsPublished") And 1),"selected","") %>>已发布</option>
                                    <option value="0" <%= IIf(Not CBool(editPage("IsPublished") And 1),"selected","") %>>草稿</option>
                                </select>
                            </div>
                        </div>
                        <div class="form-group">
                            <label>Meta描述</label>
                            <input type="text" name="metaDescription" value="<%= editPage("MetaDescription") %>" placeholder="用于SEO的页面描述...">
                        </div>
                    </div>
                    <div>
                        <label style="display:block;color:#999;font-size:13px;margin-bottom:5px;">预览</label>
                        <div class="content-preview">
                            <%= editPage("Content") %>
                        </div>
                    </div>
                </div>
                <div class="form-group">
                    <label>页面内容 (HTML)</label>
                    <textarea name="content"><%= editPage("Content") %></textarea>
                </div>
                <div class="form-actions">
                    <a href="content_pages.asp" class="btn btn-ghost">取消</a>
                    <button type="submit" class="btn btn-success"><i class="fas fa-save"></i> 保存页面</button>
                </div>
            </form>
        </div>
    </div>
    <% End If
       End If
       End If %>
    <% End If %>

    <!-- 新建按钮 -->
    <div style="margin-bottom:20px;">
        <a href="?edit=new" class="btn btn-primary"><i class="fas fa-plus"></i> 新建页面</a>
    </div>

    <% If editMode = False Or editID = 0 Then
    ' Show new page form if ?edit=new
    If Request.QueryString("edit") = "new" Then %>
    <div class="card">
        <div class="card-header"><i class="fas fa-plus-circle"></i> 新建页面</div>
        <div class="card-body">
            <form method="post">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="pageID" value="0">
                <div class="editor-row">
                    <div>
                        <div class="form-group"><label>页面标题</label><input type="text" name="title" required></div>
                        <div class="form-group"><label>URL标识 (Slug)</label><input type="text" name="slug" required></div>
                        <div style="display:flex;gap:15px;">
                            <div class="form-group" style="flex:1;"><label>排序</label><input type="number" name="sortOrder" value="0" min="0" style="max-width:100px;"></div>
                            <div class="form-group" style="flex:1;"><label>状态</label><select name="isPublished"><option value="0">草稿</option><option value="1">已发布</option></select></div>
                        </div>
                        <div class="form-group"><label>Meta描述</label><input type="text" name="metaDescription" placeholder="SEO描述..."></div>
                    </div>
                    <div></div>
                </div>
                <div class="form-group"><label>页面内容 (HTML)</label><textarea name="content"><p>在此输入页面内容...</p></textarea></div>
                <div class="form-actions">
                    <a href="content_pages.asp" class="btn btn-ghost">取消</a>
                    <button type="submit" class="btn btn-success"><i class="fas fa-save"></i> 创建页面</button>
                </div>
            </form>
        </div>
    </div>
    <% End If %>

    <!-- 页面列表 -->
    <div class="card">
        <div class="card-header"><i class="fas fa-list"></i> 页面列表 (<%= pagesCountLabel %>)</div>
        <div class="card-body">
            <table>
                <thead><tr><th>排序</th><th>标题</th><th>Slug</th><th>状态</th><th>更新时间</th><th>操作</th></tr></thead>
                <tbody>
                    <% If Not rsPages Is Nothing Then
                    Dim hasPages : hasPages = False
                    Do While Not rsPages.EOF
                        hasPages = True %>
                    <tr>
                        <td><%= rsPages("SortOrder") %></td>
                        <td><strong><%= rsPages("Title") %></strong></td>
                        <td><span class="slug-text">/<%= rsPages("Slug") %></span></td>
                        <td>
                            <span class="badge <%= IIf(CBool(rsPages("IsPublished") And 1),"badge-published","badge-draft") %>">
                                <%= IIf(CBool(rsPages("IsPublished") And 1),"已发布","草稿") %>
                            </span>
                        </td>
                        <td style="font-size:13px;color:#999;"><%= rsPages("UpdatedAt") %></td>
                        <td>
                            <a href="?edit=<%= rsPages("PageID") %>" class="btn btn-primary" style="padding:5px 10px;font-size:11px;"><i class="fas fa-edit"></i> 编辑</a>
                            <form method="post" style="display:inline;" onsubmit="return confirm('确认删除该页面？')">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="pageID" value="<%= rsPages("PageID") %>">
                                <button type="submit" class="btn btn-danger" style="padding:5px 10px;font-size:11px;"><i class="fas fa-trash"></i></button>
                            </form>
                        </td>
                    </tr>
                    <% 
                        rsPages.MoveNext
                    Loop
                    rsPages.Close
                    If Not hasPages Then %>
                    <tr><td colspan="6" class="empty"><i class="fas fa-file"></i>暂无内容页面</td></tr>
                    <% End If
                    Else %>
                    <tr><td colspan="6" class="empty"><i class="fas fa-file"></i>暂无内容页面</td></tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
    <% End If %>
</div>
</body>
</html>
<%
If IsObject(editPage) Then
    If Not editPage Is Nothing Then
        If editPage.State = 1 Then editPage.Close
        Set editPage = Nothing
    End If
End If
Call CloseConnection()
%>
