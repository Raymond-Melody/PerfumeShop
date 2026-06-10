<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查用户是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("URL"))
End If

Call OpenConnection()

' 处理表单提交
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        Response.Write "<script>alert('安全验证失败，请刷新页面重试'); history.back();</script>"
        Response.End
    End If
    
    Dim fullName, email, phone, address, city, postalCode
    fullName = SafeSQL(Trim(Request.Form("fullName")))
    email = SafeSQL(Trim(Request.Form("email")))
    phone = SafeSQL(Trim(Request.Form("phone")))
    address = SafeSQL(Trim(Request.Form("address")))
    city = SafeSQL(Trim(Request.Form("city")))
    postalCode = SafeSQL(Trim(Request.Form("postalCode")))
    
    ' 更新用户信息
    Dim sql
    sql = "UPDATE Users SET FullName = '" & fullName & "', Email = '" & email & "', Phone = '" & phone & "', Address = '" & address & "', City = '" & city & "', PostalCode = '" & postalCode & "' WHERE UserID = " & Session("UserID")
    
    If ExecuteNonQuery(sql) Then
        ' 刷新Session信息
        Session("FullName") = fullName
        Response.Write "<script>alert('设置已保存'); location.href='settings.asp';</script>"
    Else
        Response.Write "<script>alert('保存失败'); location.href='settings.asp';</script>"
    End If
End If

Dim userId, rsUser
userId = Session("UserID")

' 获取用户信息
Set rsUser = ExecuteQuery("SELECT * FROM Users WHERE UserID = " & userId & " AND IsActive <> 0")
If rsUser Is Nothing Or rsUser.EOF Then
    Response.Redirect "/user/login.asp"
End If

' 确保CSRF令牌存在
Call EnsureCSRFToken()
%>
<!--#include file="../includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/user/index.asp">个人中心</a>
        <span class="separator">/</span>
        <span>账户设置</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!-- 侧边栏 -->
        <aside class="user-sidebar">
            <div class="user-profile">
                <h3><%= HTMLEncode(Session("Username")) %></h3>
                <p><%= HTMLEncode(Session("Email")) %></p>
            </div>
            
            <nav class="user-nav">
                <a href="/user/index.asp"><i class="fas fa-home"></i> 个人中心</a>
                <a href="/user/orders.asp"><i class="fas fa-list"></i> 我的订单</a>
                <a href="/user/settings.asp" class="active"><i class="fas fa-user-edit"></i> 账户设置</a>
                <a href="/user/addresses.asp"><i class="fas fa-map-marker-alt"></i> 收货地址</a>
                <a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a>
                <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
            </nav>
        </aside>
        
        <!-- 主内容 -->
        <div class="user-main">
            <div class="user-card">
                <h2 class="card-title"><i class="fas fa-cog"></i> 账户设置</h2>
                
                <form method="post" class="form-horizontal" action="settings.asp">
                    <%= GetCSRFTokenField() %>
                    <div class="form-group">
                        <label for="fullName">姓名</label>
                        <input type="text" id="fullName" name="fullName" value="<%= HTMLEncode(rsUser("FullName")) %>" placeholder="请输入您的姓名">
                    </div>
                    
                    <div class="form-group">
                        <label for="email">邮箱</label>
                        <input type="email" id="email" name="email" value="<%= HTMLEncode(rsUser("Email")) %>" placeholder="请输入邮箱地址">
                    </div>
                    
                    <div class="form-group">
                        <label for="phone">手机号</label>
                        <input type="tel" id="phone" name="phone" value="<%= HTMLEncode(rsUser("Phone")) %>" placeholder="请输入手机号">
                    </div>
                    
                    <div class="form-group">
                        <label for="address">地址</label>
                        <textarea id="address" name="address" placeholder="请输入收货地址"><%= HTMLEncode(rsUser("Address")) %></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label for="city">城市</label>
                        <input type="text" id="city" name="city" value="<%= HTMLEncode(rsUser("City")) %>" placeholder="请输入所在城市">
                    </div>
                    
                    <div class="form-group">
                        <label for="postalCode">邮政编码</label>
                        <input type="text" id="postalCode" name="postalCode" value="<%= HTMLEncode(rsUser("PostalCode")) %>" placeholder="请输入邮政编码">
                    </div>
                    
                    <div class="form-actions">
                        <button type="submit" class="btn btn-primary">保存设置</button>
                        <button type="button" class="btn btn-text" onclick="history.back()">取消</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<script>
function removeAvatar() {
    alert('头像已删除');
}
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>