<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp"
    Response.End
End If
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim userId
userId = Session("UserID")

' 检查用户角色
Dim rsUser
Set rsUser = ExecuteQuery("SELECT UserRole FROM Users WHERE UserID = " & userId)
If rsUser Is Nothing Or rsUser.EOF Or rsUser("UserRole") <> "KOL" Then
    Response.Write "<script>alert('您不是KOL用户，无法访问此页面'); location.href='/user/index.asp';</script>"
    Response.End
End If

' 处理表单提交
Dim action, message
action = Request.Form("action")
message = ""

If action = "submit" Then
    Dim pName, pDesc, pImage
    pName = SafeSQL(Request.Form("productName"))
    pDesc = SafeSQL(Request.Form("description"))
    pImage = SafeSQL(Request.Form("image"))
    
    If pName = "" Then
        message = "✗ 商品名称不能为空"
    Else
        Dim insertSql, newProductId
        insertSql = "INSERT INTO Products (ProductName, Description, ImageURL, ProductType, KOLID, ReviewStatus, IsActive, BasePrice, Category) VALUES (" & _
            "'" & pName & "', '" & pDesc & "', '" & pImage & "', 'KOL', " & userId & ", 'Pending', 0, 0, 'KOL推荐')"
        
        If ExecuteNonQuery(insertSql) Then
            ' 获取ID
            Dim rsId
            Set rsId = ExecuteQuery("SELECT SCOPE_IDENTITY()")
            newProductId = rsId(0)
            rsId.Close
            
            ' 保存配比
            Dim noteIds, nId, nType, nPercent
            noteIds = Request.Form("selectedNotes") ' 逗号分隔的NoteID
            If noteIds <> "" Then
                Dim noteArr
                noteArr = Split(noteIds, ",")
                For Each nId In noteArr
                    If IsNumeric(nId) Then
                        nType = Request.Form("noteType_" & nId)
                        nPercent = Request.Form("notePercent_" & nId)
                        If IsNumeric(nPercent) Then
                            ExecuteNonQuery("INSERT INTO ProductNoteRatios (ProductID, NoteID, NoteType, Percentage) VALUES (" & _
                                newProductId & ", " & CLng(nId) & ", '" & SafeSQL(nType) & "', " & CLng(nPercent) & ")")
                        End If
                    End If
                Next
            End If
            message = "✓ 商品提交成功，请等待管理员审核！"
        Else
            message = "✗ 提交失败：" & Session("LastDBError")
        End If
    End If
End If

' 获取我的推荐列表
Dim rsMyProducts
Set rsMyProducts = ExecuteQuery("SELECT * FROM Products WHERE KOLID = " & userId & " ORDER BY CreatedAt DESC")

' 从SiteSettings获取香调最小比例配置
Dim minTopPercent, minMiddlePercent, minBasePercent
minTopPercent = 10
minMiddlePercent = 10
minBasePercent = 10

Dim rsMinPercent
Set rsMinPercent = ExecuteQuery("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey IN ('MinTopPercent', 'MinMiddlePercent', 'MinBasePercent')")
If Not rsMinPercent Is Nothing Then
    Do While Not rsMinPercent.EOF
        Select Case rsMinPercent("SettingKey")
            Case "MinTopPercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minTopPercent = CInt(rsMinPercent("SettingValue"))
            Case "MinMiddlePercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minMiddlePercent = CInt(rsMinPercent("SettingValue"))
            Case "MinBasePercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minBasePercent = CInt(rsMinPercent("SettingValue"))
        End Select
        rsMinPercent.MoveNext
    Loop
    rsMinPercent.Close
End If
Set rsMinPercent = Nothing

' 获取基香用于选择
Dim rsNotes
Set rsNotes = ExecuteQuery("SELECT * FROM FragranceNotes WHERE IsActive <> 0 ORDER BY NoteType, NoteName")
%>
<!--#include file="../includes/header.asp"-->

<div class="container">
    <div class="user-center">
        <!-- 侧边栏 -->
        <aside class="user-sidebar">
            <div class="user-profile">
                <h3><%= HTMLEncode(Session("Username")) %></h3>
                <p><span class="status-badge" style="background:#eb2f96;">KOL 推荐官</span></p>
            </div>
            <nav class="user-nav">
                <a href="/user/index.asp"><i class="fas fa-home"></i> 个人中心</a>
                <a href="/user/kol_products.asp" class="active"><i class="fas fa-star"></i> KOL推荐管理</a>
                <a href="/user/orders.asp"><i class="fas fa-list"></i> 我的订单</a>
                <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
            </nav>
        </aside>

        <!-- 主内容 -->
        <div class="user-main">
            <div class="welcome-section">
                <h1>KOL 推荐管理</h1>
                <p>在这里创作您的专属香氛方案并推荐给粉丝</p>
            </div>

            <% If message <> "" Then %>
                <div class="alert <%= IIF(InStr(message, "✓") > 0, "alert-success", "alert-error") %>" style="margin-bottom:20px; padding:15px; border-radius:4px; <%= IIF(InStr(message, "✓") > 0, "background:#d4edda;color:#155724;", "background:#f8d7da;color:#721c24;") %>">
                    <%= message %>
                </div>
            <% End If %>

            <!-- 提交新推荐按钮 -->
            <div style="margin-bottom: 20px; text-align: right;">
                <button class="btn btn-primary" onclick="toggleAddForm()"><i class="fas fa-plus"></i> 创建新推荐</button>
            </div>

            <!-- 添加表单 (默认隐藏) -->
            <div id="addFormBox" style="display:none; background: #fff; padding: 25px; border-radius: 8px; box-shadow: 0 2px 15px rgba(0,0,0,0.1); margin-bottom: 30px;">
                <h2 style="margin-bottom: 20px; border-bottom: 2px solid #eee; padding-bottom: 10px;">设计您的专属香氛</h2>
                <form method="post" id="kolForm">
                    <input type="hidden" name="action" value="submit">
                    <input type="hidden" name="selectedNotes" id="selectedNotesInput" value="">
                    
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label style="display:block; margin-bottom:5px; font-weight:bold;">方案名称 *</label>
                        <input type="text" name="productName" class="form-control" placeholder="给您的独家配方起个好听的名字" required style="width:100%; padding:10px; border:1px solid #ddd;">
                    </div>
                    
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label style="display:block; margin-bottom:5px; font-weight:bold;">推荐语 / 灵感来源</label>
                        <textarea name="description" class="form-control" rows="3" placeholder="告诉粉丝这个味道背后的故事..." style="width:100%; padding:10px; border:1px solid #ddd;"></textarea>
                    </div>

                    <div class="form-group" style="margin-bottom: 15px;">
                        <label style="display:block; margin-bottom:5px; font-weight:bold;">方案封面图片 URL</label>
                        <input type="text" name="image" class="form-control" value="/images/default-product.svg" style="width:100%; padding:10px; border:1px solid #ddd;">
                    </div>

                    <!-- 香调选择逻辑 -->
                    <h3 style="font-size: 16px; margin: 25px 0 10px 0;"><i class="fas fa-flask"></i> 配方组合 (前中后调配比)</h3>
                    <div class="note-selection-grid" style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px;">
                        <!-- 前调 -->
                        <div class="note-col">
                            <h4 style="border-bottom: 1px solid #eee; padding-bottom:5px; margin-bottom:10px; color:#f39c12;">前调</h4>
                            <div style="max-height: 200px; overflow-y: auto;">
                                <% 
                                If Not rsNotes Is Nothing Then
                                    rsNotes.Filter = "NoteType='前调'"
                                    Do While Not rsNotes.EOF
                                %>
                                <label style="display:block; margin-bottom:8px; font-size:13px; cursor:pointer;">
                                    <input type="checkbox" class="note-cb" data-id="<%= rsNotes("NoteID") %>" data-type="前调" onclick="handleNoteClick(this)"> <%= HTMLEncode(rsNotes("NoteName")) %>
                                </label>
                                <div id="ratio_<%= rsNotes("NoteID") %>" style="display:none; margin-left:20px; margin-bottom:10px;">
                                    <input type="number" class="note-ratio-input" data-id="<%= rsNotes("NoteID") %>" min="1" max="90" value="30" style="width:60px; padding:2px 5px;" onchange="syncInputs()"> %
                                    <input type="hidden" name="noteType_<%= rsNotes("NoteID") %>" value="前调">
                                    <input type="hidden" name="notePercent_<%= rsNotes("NoteID") %>" id="input_percent_<%= rsNotes("NoteID") %>" value="30">
                                </div>
                                <%
                                        rsNotes.MoveNext
                                    Loop
                                End If
                                %>
                            </div>
                        </div>
                        <!-- 中调 -->
                        <div class="note-col">
                            <h4 style="border-bottom: 1px solid #eee; padding-bottom:5px; margin-bottom:10px; color:#e74c3c;">中调</h4>
                            <div style="max-height: 200px; overflow-y: auto;">
                                <% 
                                If Not rsNotes Is Nothing Then
                                    rsNotes.Filter = "NoteType='中调'"
                                    Do While Not rsNotes.EOF
                                %>
                                <label style="display:block; margin-bottom:8px; font-size:13px; cursor:pointer;">
                                    <input type="checkbox" class="note-cb" data-id="<%= rsNotes("NoteID") %>" data-type="中调" onclick="handleNoteClick(this)"> <%= HTMLEncode(rsNotes("NoteName")) %>
                                </label>
                                <div id="ratio_<%= rsNotes("NoteID") %>" style="display:none; margin-left:20px; margin-bottom:10px;">
                                    <input type="number" class="note-ratio-input" data-id="<%= rsNotes("NoteID") %>" min="1" max="90" value="40" style="width:60px; padding:2px 5px;" onchange="syncInputs()"> %
                                    <input type="hidden" name="noteType_<%= rsNotes("NoteID") %>" value="中调">
                                    <input type="hidden" name="notePercent_<%= rsNotes("NoteID") %>" id="input_percent_<%= rsNotes("NoteID") %>" value="40">
                                </div>
                                <%
                                        rsNotes.MoveNext
                                    Loop
                                End If
                                %>
                            </div>
                        </div>
                        <!-- 后调 -->
                        <div class="note-col">
                            <h4 style="border-bottom: 1px solid #eee; padding-bottom:5px; margin-bottom:10px; color:#3498db;">后调</h4>
                            <div style="max-height: 200px; overflow-y: auto;">
                                <% 
                                If Not rsNotes Is Nothing Then
                                    rsNotes.Filter = "NoteType='后调'"
                                    Do While Not rsNotes.EOF
                                %>
                                <label style="display:block; margin-bottom:8px; font-size:13px; cursor:pointer;">
                                    <input type="checkbox" class="note-cb" data-id="<%= rsNotes("NoteID") %>" data-type="后调" onclick="handleNoteClick(this)"> <%= HTMLEncode(rsNotes("NoteName")) %>
                                </label>
                                <div id="ratio_<%= rsNotes("NoteID") %>" style="display:none; margin-left:20px; margin-bottom:10px;">
                                    <input type="number" class="note-ratio-input" data-id="<%= rsNotes("NoteID") %>" min="1" max="90" value="30" style="width:60px; padding:2px 5px;" onchange="syncInputs()"> %
                                    <input type="hidden" name="noteType_<%= rsNotes("NoteID") %>" value="后调">
                                    <input type="hidden" name="notePercent_<%= rsNotes("NoteID") %>" id="input_percent_<%= rsNotes("NoteID") %>" value="30">
                                </div>
                                <%
                                        rsNotes.MoveNext
                                    Loop
                                    rsNotes.Filter = ""
                                End If
                                %>
                            </div>
                        </div>
                    </div>

                    <div style="margin-top: 20px; padding: 15px; background: #fffbe6; border: 1px solid #ffe58f; border-radius: 4px;">
                        <span id="totalRatioLabel">当前总比例: <strong>0%</strong></span> (总比例必须等于 100% 才能提交)
                    </div>

                    <div style="margin-top: 30px; text-align: center;">
                        <button type="button" class="btn btn-outline" onclick="toggleAddForm()" style="margin-right:15px;">取消</button>
                        <button type="submit" id="submitBtn" class="btn btn-primary" disabled>提交审核</button>
                    </div>
                </form>
            </div>

            <!-- 列表部分 -->
            <div class="recent-orders">
                <div class="section-header">
                    <h2>我的推荐方案</h2>
                </div>
                
                <div class="table-responsive" style="background: white; border-radius: 8px; overflow: hidden;">
                    <table style="width: 100%; border-collapse: collapse;">
                        <thead>
                            <tr style="background: #f8f9fa; border-bottom: 2px solid #eee;">
                                <th style="padding: 15px; text-align: left;">方案</th>
                                <th style="padding: 15px; text-align: center;">状态</th>
                                <th style="padding: 15px; text-align: center;">上架情况</th>
                                <th style="padding: 15px; text-align: center;">日期</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% If Not rsMyProducts Is Nothing And Not rsMyProducts.EOF Then %>
                                <% Do While Not rsMyProducts.EOF %>
                                <tr style="border-bottom: 1px solid #eee;">
                                    <td style="padding: 15px;">
                                        <div style="display: flex; align-items: center;">
                                            <img src="<%= rsMyProducts("ImageURL") %>" style="width: 50px; height: 50px; border-radius: 4px; margin-right: 15px; object-fit: cover;">
                                            <div>
                                                <strong><%= HTMLEncode(rsMyProducts("ProductName")) %></strong>
                                                <p style="font-size: 12px; color: #999; margin: 0;"><%= Left(HTMLEncode(rsMyProducts("Description") & ""), 30) %>...</p>
                                            </div>
                                        </div>
                                    </td>
                                    <td style="padding: 15px; text-align: center;">
                                        <%
                                        Dim rStatus, rClass
                                        rStatus = rsMyProducts("ReviewStatus") & ""
                                        Select Case rStatus
                                            Case "Pending": rClass = "background:#f39c12; color:white; padding:2px 8px; border-radius:10px; font-size:12px;"; rStatus = "待审核"
                                            Case "Approved": rClass = "background:#27ae60; color:white; padding:2px 8px; border-radius:10px; font-size:12px;"; rStatus = "通过"
                                            Case "Rejected": rClass = "background:#e74c3c; color:white; padding:2px 8px; border-radius:10px; font-size:12px;"; rStatus = "驳回"
                                        End Select
                                        %>
                                        <span style="<%= rClass %>"><%= rStatus %></span>
                                    </td>
                                    <td style="padding: 15px; text-align: center;">
                                        <%= IIF(rsMyProducts("IsActive"), "<span style='color:green;'>已上架</span>", "<span style='color:gray;'>未上架</span>") %>
                                    </td>
                                    <td style="padding: 15px; text-align: center; color: #999; font-size: 12px;">
                                        <%= SafeFormatDateTime(rsMyProducts("CreatedAt"), 2) %>
                                    </td>
                                </tr>
                                <% rsMyProducts.MoveNext %>
                                <% Loop %>
                            <% Else %>
                                <tr>
                                    <td colspan="4" style="padding: 40px; text-align: center; color: #999;">您还没有创建任何推荐方案</td>
                                </tr>
                            <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

<script>
    // 从SiteSettings读取的最小比例配置
    var minTopPercent = <%= minTopPercent %>;
    var minMiddlePercent = <%= minMiddlePercent %>;
    var minBasePercent = <%= minBasePercent %>;

    function toggleAddForm() {
        var box = document.getElementById('addFormBox');
        box.style.display = box.style.display === 'none' ? 'block' : 'none';
        if (box.style.display === 'block') box.scrollIntoView({ behavior: 'smooth' });
    }

    function handleNoteClick(cb) {
        var id = cb.getAttribute('data-id');
        var ratioBox = document.getElementById('ratio_' + id);
        ratioBox.style.display = cb.checked ? 'block' : 'none';
        calculateTotal();
    }

    function syncInputs() {
        var inputs = document.querySelectorAll('.note-ratio-input');
        inputs.forEach(function(input) {
            var id = input.getAttribute('data-id');
            document.getElementById('input_percent_' + id).value = input.value;
        });
        calculateTotal();
    }

    function calculateTotal() {
        var total = 0;
        var selectedIds = [];
        var cbs = document.querySelectorAll('.note-cb:checked');
        var topTotal = 0, middleTotal = 0, baseTotal = 0;
        
        cbs.forEach(function(cb) {
            var id = cb.getAttribute('data-id');
            var noteType = cb.getAttribute('data-type');
            selectedIds.push(id);
            var ratioInput = document.querySelector('.note-ratio-input[data-id="' + id + '"]');
            var val = parseInt(ratioInput.value || 0);
            total += val;
            
            // 按调性统计
            if (noteType === '前调') topTotal += val;
            else if (noteType === '中调') middleTotal += val;
            else if (noteType === '后调') baseTotal += val;
        });

        document.getElementById('selectedNotesInput').value = selectedIds.join(',');
        var label = document.getElementById('totalRatioLabel');
        
        // 验证各调性最小比例
        var errorMsg = '';
        if (topTotal > 0 && topTotal < minTopPercent) errorMsg += '前调至少需要' + minTopPercent + '%；';
        if (middleTotal > 0 && middleTotal < minMiddlePercent) errorMsg += '中调至少需要' + minMiddlePercent + '%；';
        if (baseTotal > 0 && baseTotal < minBasePercent) errorMsg += '后调至少需要' + minBasePercent + '%；';
        
        if (errorMsg) {
            label.innerHTML = '当前总比例: <strong style="color:red">' + total + '%</strong> <span style="color:#e74c3c;font-size:12px;">' + errorMsg + '</span>';
            document.getElementById('submitBtn').disabled = true;
        } else {
            label.innerHTML = '当前总比例: <strong style="' + (total === 100 ? 'color:green' : 'color:red') + '">' + total + '%</strong>';
            document.getElementById('submitBtn').disabled = (total !== 100 || selectedIds.length === 0);
        }
    }
</script>

<!--#include file="../includes/footer.asp"-->
<%
If Not rsMyProducts Is Nothing Then rsMyProducts.Close: Set rsMyProducts = Nothing
If Not rsUser Is Nothing Then rsUser.Close: Set rsUser = Nothing
Call CloseConnection()
%>
