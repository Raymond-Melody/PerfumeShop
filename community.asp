<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<%
If Not FEATURE_COMMUNITY Then Response.Redirect "/index.asp"
Call OpenConnection()

Dim comMsg, comMsgType, comView, comPostID, comType
comMsg = "" : comMsgType = ""
comView = Request.QueryString("view")
comPostID = Request.QueryString("id")
comType = Request.QueryString("type")
If comType = "" Then comType = "discussion"

' 发帖
If Request.Form("action") = "create_post" And Session("UserID") <> "" Then
    Dim pUserID : pUserID = CLng(Session("UserID"))
    Dim pTitle : pTitle = Trim(Request.Form("title"))
    Dim pContent : pContent = Trim(Request.Form("content"))
    Dim pType : pType = Request.Form("post_type")
    Dim pTags : pTags = Trim(Request.Form("tags"))
    Dim pTop : pTop = Trim(Request.Form("top_notes"))
    Dim pMid : pMid = Trim(Request.Form("mid_notes"))
    Dim pBase : pBase = Trim(Request.Form("base_notes"))

    If pTitle = "" Or pContent = "" Then
        comMsg = "请填写标题和内容" : comMsgType = "error"
    Else
        Dim fragJson : fragJson = ""
        If pTop <> "" Or pMid <> "" Or pBase <> "" Then
            fragJson = "{""top"":""" & Replace(pTop, """", "\""") & """,""middle"":""" & Replace(pMid, """", "\""") & """,""base"":""" & Replace(pBase, """", "\""") & """}"
        End If
        DAL_Execute "INSERT INTO CommunityPosts (UserID, Title, Content, PostType, FragranceNotes, Tags) VALUES (@UID,@Title,@Content,@Type,@Frag,@Tags)", _
            Array(Array("@UID", DAL_adInteger, 0, pUserID), _
                  Array("@Title", DAL_adVarWChar, 200, pTitle), _
                  Array("@Content", DAL_adVarWChar, 4000, pContent), _
                  Array("@Type", DAL_adVarWChar, 20, pType), _
                  Array("@Frag", DAL_adVarWChar, 500, fragJson), _
                  Array("@Tags", DAL_adVarWChar, 300, pTags))
        comMsg = "发布成功！" : comMsgType = "success"
    End If
End If

' 评论
If Request.Form("action") = "add_comment" And Session("UserID") <> "" Then
    Dim cPostID : cPostID = Request.Form("post_id")
    Dim cContent : cContent = Trim(Request.Form("content"))
    Dim cParentID : cParentID = Request.Form("parent_id")
    If cParentID = "" Then cParentID = Null
    If cContent <> "" And IsNumeric(cPostID) Then
        DAL_Execute "INSERT INTO PostComments (PostID, UserID, ParentCommentID, Content) VALUES (@PID,@UID,@PCID,@Content)", _
            Array(Array("@PID", DAL_adInteger, 0, CLng(cPostID)), _
                  Array("@UID", DAL_adInteger, 0, CLng(Session("UserID"))), _
                  Array("@PCID", DAL_adInteger, 0, cParentID), _
                  Array("@Content", DAL_adVarWChar, 1000, cContent))
        DAL_Execute "UPDATE CommunityPosts SET CommentCount = CommentCount + 1 WHERE PostID = @PID", _
            Array(Array("@PID", DAL_adInteger, 0, CLng(cPostID)))
        comMsg = "评论成功！"
    End If
End If

' 点赞
If Request.QueryString("action") = "like" And Session("UserID") <> "" Then
    Dim lPostID : lPostID = Request.QueryString("id")
    If IsNumeric(lPostID) Then
        Dim liked : liked = DAL_GetScalar("SELECT COUNT(*) FROM PostLikes WHERE PostID=@PID AND UserID=@UID", _
            Array(Array("@PID", DAL_adInteger, 0, CLng(lPostID)), Array("@UID", DAL_adInteger, 0, CLng(Session("UserID")))), 0)
        If CLng(liked) = 0 Then
            DAL_Execute "INSERT INTO PostLikes (PostID, UserID) VALUES (@PID,@UID)", _
                Array(Array("@PID", DAL_adInteger, 0, CLng(lPostID)), Array("@UID", DAL_adInteger, 0, CLng(Session("UserID"))))
            DAL_Execute "UPDATE CommunityPosts SET LikeCount = LikeCount + 1 WHERE PostID = @PID", _
                Array(Array("@PID", DAL_adInteger, 0, CLng(lPostID)))
        End If
    End If
End If

' 搜索
Dim comSearch : comSearch = Trim(Request.QueryString("search"))
Dim comSQL, comWhere
comWhere = "WHERE p.IsActive = 1"
If comView = "detail" And IsNumeric(comPostID) Then
    comWhere = "WHERE p.PostID = " & CLng(comPostID)
Else
    If comSearch <> "" Then
        comWhere = comWhere & " AND (p.Title LIKE '%" & SafeSQL(comSearch) & "%' OR p.Content LIKE '%" & SafeSQL(comSearch) & "%' OR p.Tags LIKE '%" & SafeSQL(comSearch) & "%')"
    End If
    comWhere = comWhere & " AND p.PostType = '" & SafeSQL(comType) & "'"
End If
comSQL = "SELECT p.*, u.Username FROM CommunityPosts p LEFT JOIN Users u ON p.UserID = u.UserID " & comWhere & " ORDER BY p.IsPinned DESC, p.CreatedAt DESC"

Dim rsPosts : Set rsPosts = conn.Execute(comSQL)

Function PostTypeLabel(pt)
    Select Case LCase(pt)
        Case "recipe": PostTypeLabel = "配方分享"
        Case "review": PostTypeLabel = "香评"
        Case Else: PostTypeLabel = "讨论"
    End Select
End Function

Function Stars(rating)
    Dim s, i
    For i = 1 To 5
        If i <= rating Then s = s & "<i class='fas fa-star'></i>" Else s = s & "<i class='far fa-star'></i>"
    Next
    Stars = s
End Function
%>
<!--#include file="includes/header.asp"-->

<section class="community-hero">
    <div class="container text-center">
        <h1><i class="fas fa-users"></i> 香氛社区</h1>
        <p>分享你的香氛故事，发现更多香气灵感</p>
        <% If Session("UserID") <> "" Then %>
        <a href="#createPost" class="btn btn-primary btn-lg" onclick="document.getElementById('createPost').scrollIntoView({behavior:'smooth'});return false;">
            <i class="fas fa-feather-alt"></i> 发布帖子
        </a>
        <% End If %>
    </div>
</section>

<div class="container community-page">
    <% If comMsg <> "" Then %>
    <div class="alert <%= IIf(comMsgType="error","alert-error","alert-success") %>">
        <i class="fas fa-info-circle"></i> <%= comMsg %>
    </div>
    <% End If %>

    <!-- 搜索与分类 -->
    <div class="community-toolbar">
        <div class="community-tabs">
            <a href="?type=discussion<%= IIf(comSearch<>"","&search=" & Server.URLEncode(comSearch),"") %>" class="tab <%= IIf(comType="discussion","active","") %>">
                <i class="fas fa-comments"></i> 讨论区
            </a>
            <a href="?type=review<%= IIf(comSearch<>"","&search=" & Server.URLEncode(comSearch),"") %>" class="tab <%= IIf(comType="review","active","") %>">
                <i class="fas fa-star"></i> 香评墙
            </a>
            <a href="?type=recipe<%= IIf(comSearch<>"","&search=" & Server.URLEncode(comSearch),"") %>" class="tab <%= IIf(comType="recipe","active","") %>">
                <i class="fas fa-flask"></i> 配方分享
            </a>
        </div>
        <form class="community-search" method="get">
            <input type="hidden" name="type" value="<%= Server.HTMLEncode(comType) %>">
            <input type="text" name="search" value="<%= Server.HTMLEncode(comSearch) %>" placeholder="搜索帖子...">
            <button type="submit"><i class="fas fa-search"></i></button>
        </form>
    </div>

    <% If comView = "detail" And IsNumeric(comPostID) Then %>
        <!-- 帖子详情 -->
        <%
        If Not rsPosts Is Nothing And Not rsPosts.EOF Then
            Dim dID, dUID, dTitle, dContent, dType, dFrag, dTags, dLikes, dComments, dViews, dPinned, dUser, dDate
            dID = rsPosts("PostID")
            dUID = rsPosts("UserID")
            dTitle = rsPosts("Title")
            dContent = rsPosts("Content")
            dType = rsPosts("PostType")
            dFrag = rsPosts("FragranceNotes")
            dTags = rsPosts("Tags")
            dLikes = rsPosts("LikeCount")
            dComments = rsPosts("CommentCount")
            dViews = rsPosts("ViewCount")
            dPinned = rsPosts("IsPinned")
            dUser = rsPosts("Username")
            dDate = rsPosts("CreatedAt")
            conn.Execute "UPDATE CommunityPosts SET ViewCount = ViewCount + 1 WHERE PostID = " & dID
        %>
        <div class="post-detail">
            <div class="post-detail-header">
                <div class="post-badges">
                    <span class="post-type-badge"><%= PostTypeLabel(dType) %></span>
                    <% If dPinned Then %><span class="post-pin-badge">置顶</span><% End If %>
                </div>
                <h2><%= Server.HTMLEncode(dTitle) %></h2>
                <div class="post-meta">
                    <span><i class="fas fa-user-circle"></i> <%= Server.HTMLEncode(dUser) %></span>
                    <span><i class="far fa-clock"></i> <%= FormatDateTime(dDate, 1) %></span>
                    <span><i class="far fa-eye"></i> <%= dViews+1 %></span>
                </div>
            </div>

            <% If dFrag <> "" And dFrag <> "null" Then %>
            <div class="fragrance-display">
                <%
                Dim topN, midN, baseN
                topN = "" : midN = "" : baseN = ""
                If InStr(dFrag, """top""") > 0 Then
                    topN = Mid(dFrag, InStr(dFrag, """top"":""") + 8)
                    topN = Left(topN, InStr(topN, """") - 1)
                    midN = Mid(dFrag, InStr(dFrag, """middle"":""") + 11)
                    midN = Left(midN, InStr(midN, """") - 1)
                    baseN = Mid(dFrag, InStr(dFrag, """base"":""") + 9)
                    baseN = Left(baseN, InStr(baseN, """") - 1)
                End If
                %>
                <div class="fragrance-structure">
                    <div class="frag-note">
                        <span class="frag-label top">前调</span>
                        <span class="frag-val"><%= IIf(topN="","—",Server.HTMLEncode(topN)) %></span>
                    </div>
                    <div class="frag-note">
                        <span class="frag-label mid">中调</span>
                        <span class="frag-val"><%= IIf(midN="","—",Server.HTMLEncode(midN)) %></span>
                    </div>
                    <div class="frag-note">
                        <span class="frag-label base">后调</span>
                        <span class="frag-val"><%= IIf(baseN="","—",Server.HTMLEncode(baseN)) %></span>
                    </div>
                </div>
            </div>
            <% End If %>

            <div class="post-content">
                <%= Replace(Server.HTMLEncode(dContent), vbCrLf, "<br>") %>
            </div>

            <% If dTags <> "" Then %>
            <div class="post-tags">
                <%
                tagArr = Split(dTags, ",")
                For Each t In tagArr
                    t = Trim(t)
                    If t <> "" Then
                %>
                <a href="?type=<%= Server.URLEncode(comType) %>&search=<%= Server.URLEncode(t) %>" class="tag-badge">#<%= Server.HTMLEncode(t) %></a>
                <%  End If
                Next %>
            </div>
            <% End If %>

            <div class="post-actions">
                <% If Session("UserID") <> "" Then %>
                <a href="?action=like&id=<%= dID %>&view=detail&type=<%= Server.URLEncode(comType) %>" class="btn-like-action">
                    <i class="far fa-heart"></i> 点赞 (<%= dLikes %>)
                </a>
                <% Else %>
                <span class="btn-like-action disabled"><i class="far fa-heart"></i> 点赞 (<%= dLikes %>)</span>
                <% End If %>
                <span class="btn-like-action"><i class="far fa-comment"></i> 评论 (<%= dComments %>)</span>
            </div>
        </div>

        <!-- 评论区 -->
        <div class="comments-section">
            <h3><i class="far fa-comments"></i> 评论 (<%= dComments %>)</h3>

            <% 
            Dim rsComments : Set rsComments = DAL_GetList("SELECT c.*, u.Username FROM PostComments c LEFT JOIN Users u ON c.UserID = u.UserID WHERE c.PostID = @PID AND c.IsActive = 1 ORDER BY c.CreatedAt ASC", _
                                                Array(Array("@PID", DAL_adInteger, 0, dID)))
            If rsComments Is Nothing Or rsComments.EOF Then %>
            <div class="no-comments">暂无评论，快来抢沙发吧！</div>
            <% Else
                Do While Not rsComments.EOF %>
                <div class="comment-item">
                    <div class="comment-avatar"><i class="fas fa-user-circle"></i></div>
                    <div class="comment-body">
                        <div class="comment-info">
                            <strong><%= Server.HTMLEncode(rsComments("Username")) %></strong>
                            <span class="comment-time"><%= FormatDateTime(rsComments("CreatedAt"), 1) %></span>
                        </div>
                        <p><%= Server.HTMLEncode(rsComments("Content")) %></p>
                    </div>
                </div>
            <%
                    rsComments.MoveNext
                Loop
            End If
            If Not rsComments Is Nothing And rsComments.State = 1 Then rsComments.Close : Set rsComments = Nothing
            %>
        </div>

        <!-- 发表评论 -->
        <% If Session("UserID") <> "" Then %>
        <div class="comment-form-section">
            <h4>发表评论</h4>
            <form method="post" class="comment-form">
                <input type="hidden" name="action" value="add_comment">
                <input type="hidden" name="post_id" value="<%= dID %>">
                <textarea name="content" rows="3" placeholder="写下你的想法..." required></textarea>
                <button type="submit" class="btn btn-primary"><i class="fas fa-paper-plane"></i> 发表评论</button>
            </form>
        </div>
        <% End If %>
        <%
        End If
        %>
    <% Else %>
        <!-- 帖子列表 -->
        <div class="posts-list">
            <%
            If rsPosts Is Nothing Or rsPosts.EOF Then
            %>
            <div class="empty-state">
                <i class="fas fa-comments"></i>
                <p>暂无帖子，快来发布第一条吧！</p>
            </div>
            <%
            Else
                Do While Not rsPosts.EOF
                    Dim pID, pUID, pFrag, pLikes, pComments, pViews, pPinned, pUser, pDate
                    pID = rsPosts("PostID")
                    pUID = rsPosts("UserID")
                    pTitle = rsPosts("Title")
                    pContent = rsPosts("Content")
                    pType = rsPosts("PostType")
                    pFrag = rsPosts("FragranceNotes")
                    pTags = rsPosts("Tags")
                    pLikes = rsPosts("LikeCount")
                    pComments = rsPosts("CommentCount")
                    pViews = rsPosts("ViewCount")
                    pPinned = rsPosts("IsPinned")
                    pUser = rsPosts("Username")
                    pDate = rsPosts("CreatedAt")

                    Dim pPreview : pPreview = Left(Server.HTMLEncode(pContent), 150)
                    If Len(pContent) > 150 Then pPreview = pPreview & "..."
            %>
            <div class="post-card">
                <div class="post-card-header">
                    <div class="post-badges">
                        <span class="post-type-badge"><%= PostTypeLabel(pType) %></span>
                        <% If pPinned Then %><span class="post-pin-badge">置顶</span><% End If %>
                    </div>
                    <% If pTags <> "" Then %>
                    <div class="post-tags-mini">
                        <%
                        Dim ptArr, pta
                        ptArr = Split(pTags, ",")
                        For Each pta In ptArr
                            pta = Trim(pta)
                            If pta <> "" Then
                        %>
                        <a href="?type=<%= Server.URLEncode(comType) %>&search=<%= Server.URLEncode(pta) %>" class="tag-badge">#<%= Server.HTMLEncode(pta) %></a>
                        <%  End If
                        Next %>
                    </div>
                    <% End If %>
                </div>
                <h3 class="post-card-title">
                    <a href="?view=detail&id=<%= pID %>&type=<%= Server.URLEncode(comType) %>"><%= Server.HTMLEncode(pTitle) %></a>
                </h3>
                <p class="post-card-preview"><%= pPreview %></p>
                <div class="post-card-footer">
                    <div class="post-card-meta">
                        <span><i class="fas fa-user-circle"></i> <%= Server.HTMLEncode(pUser) %></span>
                        <span><i class="far fa-clock"></i> <%= FormatDateTime(pDate, 1) %></span>
                    </div>
                    <div class="post-card-stats">
                        <span><i class="far fa-eye"></i> <%= pViews %></span>
                        <span><i class="far fa-heart"></i> <%= pLikes %></span>
                        <span><i class="far fa-comment"></i> <%= pComments %></span>
                    </div>
                </div>
            </div>
            <%
                    rsPosts.MoveNext
                Loop
            End If
            %>
        </div>
    <% End If %>

    <!-- 发帖表单 -->
    <% If Session("UserID") <> "" Then %>
    <div class="create-post-section" id="createPost">
        <h3><i class="fas fa-feather-alt"></i> 发布帖子</h3>
        <form method="post" class="create-post-form">
            <input type="hidden" name="action" value="create_post">
            <div class="form-row-2">
                <div class="form-group">
                    <label>类型</label>
                    <select name="post_type">
                        <option value="discussion">讨论</option>
                        <option value="review">香评</option>
                        <option value="recipe">配方分享</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>标题 *</label>
                    <input type="text" name="title" placeholder="输入标题..." required>
                </div>
            </div>
            <div class="form-group">
                <label>内容 *</label>
                <textarea name="content" rows="5" placeholder="分享你的香氛体验..." required></textarea>
            </div>
            <div class="form-group">
                <label>标签 (逗号分隔)</label>
                <input type="text" name="tags" placeholder="如: 花香,玫瑰,夏季">
            </div>
            <div class="fragrance-form-section">
                <h5><i class="fas fa-flask"></i> 配方详情 (配方分享时填写)</h5>
                <div class="form-row-3">
                    <div class="form-group">
                        <label>前调</label>
                        <input type="text" name="top_notes" placeholder="如: 佛手柑,柠檬">
                    </div>
                    <div class="form-group">
                        <label>中调</label>
                        <input type="text" name="mid_notes" placeholder="如: 玫瑰,茉莉">
                    </div>
                    <div class="form-group">
                        <label>后调</label>
                        <input type="text" name="base_notes" placeholder="如: 檀香,琥珀">
                    </div>
                </div>
            </div>
            <button type="submit" class="btn btn-primary"><i class="fas fa-paper-plane"></i> 发布</button>
        </form>
    </div>
    <% Else %>
    <div class="create-post-section">
        <div class="login-reminder">
            <i class="fas fa-lock"></i>
            <p>登录后即可发帖、评论和点赞</p>
            <a href="/user/login.asp" class="btn btn-primary">立即登录</a>
        </div>
    </div>
    <% End If %>
</div>

<style>
.community-hero {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: #fff; padding: 50px 0 30px; text-align: center;
}
.community-hero h1 { font-size: 2rem; margin: 0 0 8px; }
.community-hero p { font-size: 1rem; opacity: 0.9; margin-bottom: 20px; }
.community-page { max-width: 860px; margin: 30px auto; padding: 0 20px; }

.community-toolbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 12px; }
.community-tabs { display: flex; gap: 4px; }
.tab { padding: 8px 18px; border-radius: 20px; text-decoration: none; color: #666; font-size: 14px; transition: all 0.2s; background: #f0f0f0; }
.tab:hover, .tab.active { background: #667eea; color: #fff; }
.tab i { margin-right: 4px; }
.community-search { display: flex; }
.community-search input { padding: 8px 14px; border: 1px solid #ddd; border-radius: 20px 0 0 20px; outline: none; min-width: 180px; }
.community-search button { padding: 8px 14px; background: #667eea; color: #fff; border: none; border-radius: 0 20px 20px 0; cursor: pointer; }

.post-card { background: #fff; border-radius: 12px; padding: 22px; margin-bottom: 16px; box-shadow: 0 1px 6px rgba(0,0,0,0.06); transition: box-shadow 0.2s; }
.post-card:hover { box-shadow: 0 3px 12px rgba(0,0,0,0.1); }
.post-card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; flex-wrap: wrap; gap: 8px; }
.post-badges { display: flex; gap: 6px; }
.post-type-badge { font-size: 11px; background: #e8eaf6; color: #5c6bc0; padding: 2px 10px; border-radius: 10px; }
.post-pin-badge { font-size: 11px; background: #ffcdd2; color: #c62828; padding: 2px 10px; border-radius: 10px; }
.tag-badge { font-size: 11px; color: #667eea; text-decoration: none; background: #f0f0ff; padding: 2px 8px; border-radius: 10px; margin-right: 4px; }
.tag-badge:hover { background: #667eea; color: #fff; }
.post-card-title { margin: 0 0 8px; font-size: 1.1rem; }
.post-card-title a { color: #333; text-decoration: none; }
.post-card-title a:hover { color: #667eea; }
.post-card-preview { font-size: 13px; color: #888; line-height: 1.6; margin: 0 0 14px; }
.post-card-footer { display: flex; justify-content: space-between; align-items: center; }
.post-card-meta { font-size: 12px; color: #aaa; display: flex; gap: 14px; }
.post-card-stats { font-size: 12px; color: #aaa; display: flex; gap: 14px; }

/* 帖子详情 */
.post-detail { background: #fff; border-radius: 12px; padding: 30px; box-shadow: 0 1px 6px rgba(0,0,0,0.06); margin-bottom: 20px; }
.post-detail-header { margin-bottom: 20px; }
.post-detail-header h2 { margin: 10px 0 8px; font-size: 1.4rem; }
.post-meta { display: flex; gap: 16px; font-size: 13px; color: #999; }
.fragrance-display { background: linear-gradient(135deg, #f3e5f5, #e8eaf6); border-radius: 10px; padding: 16px 20px; margin-bottom: 16px; }
.fragrance-structure { display: flex; gap: 20px; }
.frag-note { flex: 1; }
.frag-label { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; margin-right: 6px; color: #fff; }
.frag-label.top { background: #7cb342; }
.frag-label.mid { background: #fb8c00; }
.frag-label.base { background: #8e24aa; }
.frag-val { font-size: 13px; color: #555; }
.post-content { font-size: 15px; line-height: 1.8; color: #444; margin-bottom: 16px; }
.post-tags { margin-bottom: 16px; }
.post-actions { display: flex; gap: 16px; padding-top: 16px; border-top: 1px solid #f0f0f0; }
.btn-like-action { color: #999; text-decoration: none; font-size: 14px; display: inline-flex; align-items: center; gap: 4px; }
.btn-like-action:hover { color: #667eea; }
.btn-like-action.disabled { opacity: 0.5; cursor: not-allowed; }

.comments-section { background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 1px 6px rgba(0,0,0,0.06); margin-bottom: 20px; }
.comments-section h3 { font-size: 1rem; margin: 0 0 16px; color: #555; }
.comment-item { display: flex; gap: 12px; padding: 12px 0; border-bottom: 1px solid #f5f5f5; }
.comment-item:last-child { border-bottom: none; }
.comment-avatar { font-size: 1.8rem; color: #ccc; flex-shrink: 0; }
.comment-info { display: flex; gap: 10px; align-items: center; margin-bottom: 4px; }
.comment-info strong { font-size: 13px; }
.comment-time { font-size: 11px; color: #bbb; }
.comment-body p { margin: 0; font-size: 14px; color: #555; line-height: 1.5; }
.no-comments { text-align: center; padding: 30px; color: #ccc; }

.comment-form-section { background: #fafafa; border-radius: 12px; padding: 20px; margin-bottom: 20px; }
.comment-form-section h4 { font-size: 1rem; margin: 0 0 12px; }
.comment-form textarea { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 8px; resize: vertical; outline: none; font-size: 14px; }
.comment-form textarea:focus { border-color: #667eea; }
.comment-form button { margin-top: 10px; }

.create-post-section { background: #fafafa; border-radius: 12px; padding: 24px; margin-top: 30px; }
.create-post-section h3 { font-size: 1.1rem; margin: 0 0 16px; color: #555; }
.create-post-form .form-group { margin-bottom: 14px; }
.create-post-form .form-group label { display: block; margin-bottom: 4px; font-size: 13px; color: #888; }
.create-post-form .form-group input, .create-post-form .form-group textarea, .create-post-form .form-group select {
    width: 100%; padding: 10px 14px; border: 1px solid #ddd; border-radius: 8px; outline: none; font-size: 14px; color: #333;
}
.create-post-form .form-group input:focus, .create-post-form .form-group textarea:focus { border-color: #667eea; }
.form-row-2 { display: grid; grid-template-columns: 140px 1fr; gap: 14px; }
.form-row-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; }
.fragrance-form-section { background: #fff; border-radius: 8px; padding: 16px; margin-bottom: 14px; border: 1px dashed #ddd; }
.fragrance-form-section h5 { margin: 0 0 10px; font-size: 13px; color: #888; }

.login-reminder { text-align: center; padding: 20px; }
.login-reminder i { font-size: 2rem; color: #ddd; display: block; margin-bottom: 10px; }
.login-reminder p { color: #999; margin-bottom: 14px; }

.alert { padding: 14px 20px; border-radius: 8px; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
.alert-success { background: #d4edda; color: #155724; }
.alert-error { background: #f8d7da; color: #721c24; }
.empty-state { text-align: center; padding: 60px 20px; color: #ccc; }
.empty-state i { font-size: 3rem; display: block; margin-bottom: 14px; }

@media (max-width: 600px) {
    .form-row-2, .form-row-3 { grid-template-columns: 1fr; }
    .fragrance-structure { flex-direction: column; gap: 10px; }
    .post-card-footer { flex-direction: column; align-items: flex-start; gap: 8px; }
}
</style>

<!--#include file="includes/footer.asp"-->
<%
If Not rsPosts Is Nothing Then
    If rsPosts.State = 1 Then rsPosts.Close
    Set rsPosts = Nothing
End If
Call CloseConnection()
%>
