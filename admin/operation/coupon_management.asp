<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/promotion_engine.asp"-->
<%
Call OpenConnection()

Dim action, actionMsg, actionResult
action = Request.QueryString("action")
If action = "" Then action = Request.Form("action")
actionMsg = ""
actionResult = True

' 新增/编辑优惠券
If action = "save" And Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim editId, code, name, cType, value, minSpend, maxDiscount, validFrom, validTo, totalQty, isActive, firstOrder, category, productId, description, terms
    
    editId = Request.Form("edit_id")
    code = Trim(Request.Form("code"))
    name = Trim(Request.Form("name"))
    cType = Request.Form("type")
    value = CDbl(Request.Form("value"))
    minSpend = CDbl(Request.Form("min_spend"))
    maxDiscount = CDbl(Request.Form("max_discount"))
    validFrom = Request.Form("valid_from")
    validTo = Request.Form("valid_to")
    totalQty = Request.Form("total_qty")
    If totalQty = "" Then totalQty = "0"
    isActive = IIf(Request.Form("is_active") = "1", "1", "0")
    firstOrder = IIf(Request.Form("first_order") = "1", "1", "0")
    category = Trim(Request.Form("category"))
    productId = Request.Form("product_id")
    description = Trim(Request.Form("description"))
    terms = Trim(Request.Form("terms"))
    
    If code = "" Or name = "" Or cType = "" Then
        actionMsg = "请填写优惠码和名称"
        actionResult = False
    Else
        On Error Resume Next
        If editId <> "" And IsNumeric(editId) Then
            ' 更新
            conn.Execute "UPDATE Coupons SET CouponName='" & SafeSQL(name) & "', CouponType='" & SafeSQL(cType) & _
                       "', DiscountValue=" & value & ", MinSpend=" & minSpend & ", MaxDiscount=" & maxDiscount & _
                       ", ValidFrom='" & SafeSQL(validFrom) & "', ValidTo='" & SafeSQL(validTo) & _
                       "', TotalQty=" & totalQty & ", IsActive=" & isActive & _
                       ", FirstOrderOnly=" & firstOrder & ", ApplicableCategory='" & SafeSQL(category) & _
                       "', ApplicableProductID=" & IIf(productId="" Or Not IsNumeric(productId), "NULL", productId) & _
                       ", Description='" & SafeSQL(description) & "', Terms='" & SafeSQL(terms) & _
                       "', UpdatedAt=GETDATE() WHERE CouponID=" & editId
            If Err.Number = 0 Then actionMsg = "更新成功" Else actionMsg = "更新失败: " & Err.Description : actionResult = False
        Else
            ' 新增
            conn.Execute "INSERT INTO Coupons (CouponCode, CouponName, CouponType, DiscountValue, MinSpend, MaxDiscount, ValidFrom, ValidTo, TotalQty, IsActive, FirstOrderOnly, ApplicableCategory, ApplicableProductID, Description, Terms) VALUES ('" & _
                       SafeSQL(UCase(code)) & "','" & SafeSQL(name) & "','" & SafeSQL(cType) & "'," & value & "," & minSpend & "," & maxDiscount & ",'" & _
                       SafeSQL(validFrom) & "','" & SafeSQL(validTo) & "'," & totalQty & "," & isActive & "," & firstOrder & ",'" & _
                       SafeSQL(category) & "'," & IIf(productId="" Or Not IsNumeric(productId), "NULL", productId) & ",'" & _
                       SafeSQL(description) & "','" & SafeSQL(terms) & "')"
            If Err.Number = 0 Then actionMsg = "创建成功" Else actionMsg = "创建失败: " & Err.Description : actionResult = False
        End If
        On Error GoTo 0
    End If
End If

' 删除
If action = "delete" Then
    Dim delId
    delId = Request.QueryString("id")
    If IsNumeric(delId) Then
        conn.Execute "DELETE FROM Coupons WHERE CouponID = " & delId
        actionMsg = "已删除"
    End If
End If

' 统计
Dim totalCoupons, activeCoupons, totalIssued, totalUsed
totalCoupons = GetScalar("SELECT COUNT(*) FROM Coupons")
activeCoupons = GetScalar("SELECT COUNT(*) FROM Coupons WHERE IsActive = 1")
totalIssued = GetScalar("SELECT COUNT(*) FROM UserCoupons")
totalUsed = GetScalar("SELECT COUNT(*) FROM UserCoupons WHERE Status = 'used'")

' 获取编辑中的券
Dim editCoupon, editCouponId, editCouponData
editCouponId = Request.QueryString("edit_id")
Set editCouponData = Nothing
If editCouponId <> "" And IsNumeric(editCouponId) Then
    Dim rsEdit
    Set rsEdit = conn.Execute("SELECT * FROM Coupons WHERE CouponID = " & editCouponId)
    If Not rsEdit Is Nothing And Not rsEdit.EOF Then
        Set editCouponData = rsEdit
    End If
End If

' 所有券列表
Dim rsAllCoupons
Set rsAllCoupons = PE_CouponGetAll()
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>优惠券管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { color: #e0e0e0; padding: 24px; }
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .page-title { font-size: 22px; color: #fff; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #ff8f00; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-value { font-size: 28px; font-weight: bold; color: #ff8f00; }
        .stat-label { font-size: 13px; color: #888; margin-top: 4px; }
        
        .panel { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; padding: 24px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 24px; }
        .panel h3 { color: #fff; margin: 0 0 16px; font-size: 18px; display: flex; align-items: center; gap: 8px; }
        .panel h3 i { color: #ff8f00; }
        
        .form-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 12px; }
        .form-group { display: flex; flex-direction: column; gap: 4px; }
        .form-group label { font-size: 12px; color: #888; font-weight: 500; }
        .form-group input, .form-group select, .form-group textarea {
            padding: 8px 12px; border: 1px solid #3a3a4a; border-radius: 6px; background: #1a1a2e; color: #e0e0e0; font-size: 13px;
        }
        .form-group input:focus, .form-group select:focus, .form-group textarea:focus { border-color: #ff8f00; outline: none; }
        .form-group textarea { resize: vertical; min-height: 60px; }
        
        .btn { padding: 8px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; }
        .btn-primary { background: linear-gradient(135deg, #ff8f00, #f57c00); color: #fff; }
        .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(255,143,0,0.3); }
        .btn-danger { background: #c62828; color: #fff; }
        .btn-sm { padding: 4px 12px; font-size: 11px; }
        
        .coupon-table { width: 100%; border-collapse: collapse; }
        .coupon-table th { text-align: left; padding: 10px 12px; background: rgba(0,0,0,0.2); color: #888; font-size: 11px; text-transform: uppercase; }
        .coupon-table td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .coupon-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; }
        .badge-active { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .badge-inactive { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        .badge-fixed { background: rgba(255,87,34,0.2); color: #FF5722; }
        .badge-percent { background: rgba(156,39,176,0.2); color: #CE93D8; }
        .badge-freeship { background: rgba(33,150,243,0.2); color: #64B5F6; }
        
        .alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #ef9a9a; border: 1px solid rgba(244,67,54,0.3); }
        
        @media (max-width: 768px) { .stats-grid { grid-template-columns: repeat(2, 1fr); } }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-ticket-alt"></i> 优惠券管理</h2>
        </div>
        
        <% If actionMsg <> "" Then %>
        <div class="alert <% If actionResult Then %>alert-success<% Else %>alert-error<% End If %>">
            <i class="fas fa-<% If actionResult Then %>check-circle<% Else %>exclamation-circle<% End If %>"></i> <%= actionMsg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value"><%= totalCoupons %></div>
                <div class="stat-label">总券模板</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= activeCoupons %></div>
                <div class="stat-label">活跃中</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= totalIssued %></div>
                <div class="stat-label">总发放</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= totalUsed %></div>
                <div class="stat-label">已使用</div>
            </div>
        </div>
        
        <!-- 创建/编辑表单 -->
        <div class="panel">
            <h3><i class="fas fa-<% If Not editCouponData Is Nothing Then %>edit<% Else %>plus-circle<% End If %>"></i> <% If Not editCouponData Is Nothing Then %>编辑优惠券<% Else %>创建优惠券<% End If %></h3>
            <form method="post">
                <% If Not editCouponData Is Nothing Then %>
                <input type="hidden" name="edit_id" value="<%= editCouponData("CouponID") %>">
                <% End If %>
                <div class="form-row">
                    <div class="form-group">
                        <label>优惠码 *</label>
                        <input type="text" name="code" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("CouponCode"), "") %>" placeholder="如: SUMMER20" style="text-transform:uppercase;" required <% If Not editCouponData Is Nothing Then %>readonly<% End If %>>
                    </div>
                    <div class="form-group">
                        <label>券名称 *</label>
                        <input type="text" name="name" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("CouponName"), "") %>" placeholder="如: 夏日大促券" required>
                    </div>
                    <div class="form-group">
                        <label>券类型 *</label>
                        <select name="type" required>
                            <% 
                            Dim selType : selType = IIf(Not editCouponData Is Nothing, editCouponData("CouponType"), "fixed")
                            %>
                            <option value="fixed"<% If selType = "fixed" Then %> selected<% End If %>>满减券</option>
                            <option value="percentage"<% If selType = "percentage" Then %> selected<% End If %>>折扣券</option>
                            <option value="free_shipping"<% If selType = "free_shipping" Then %> selected<% End If %>>免邮券</option>
                            <option value="gift"<% If selType = "gift" Then %> selected<% End If %>>礼品券</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>券面额</label>
                        <input type="number" name="value" step="0.01" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("DiscountValue"), "0") %>" required>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>最低消费</label>
                        <input type="number" name="min_spend" step="0.01" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("MinSpend"), "0") %>">
                    </div>
                    <div class="form-group">
                        <label>最大优惠(折扣券用)</label>
                        <input type="number" name="max_discount" step="0.01" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("MaxDiscount"), "0") %>">
                    </div>
                    <div class="form-group">
                        <label>发行总量(0=不限)</label>
                        <input type="number" name="total_qty" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("TotalQty"), "0") %>">
                    </div>
                    <div class="form-group">
                        <label>&nbsp;</label>
                        <label style="display:flex;align-items:center;gap:6px;font-size:13px;">
                            <input type="checkbox" name="is_active" value="1"<% If editCouponData Is Nothing Or editCouponData("IsActive") Then %> checked<% End If %>> 启用
                            <input type="checkbox" name="first_order" value="1"<% If Not editCouponData Is Nothing And editCouponData("FirstOrderOnly") Then %> checked<% End If %>> 仅首单
                        </label>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>生效日期</label>
                        <input type="date" name="valid_from" value="<%= IIf(Not editCouponData Is Nothing, Left(editCouponData("ValidFrom"), 10), Left(Now(), 10)) %>" required>
                    </div>
                    <div class="form-group">
                        <label>失效日期</label>
                        <input type="date" name="valid_to" value="<%= IIf(Not editCouponData Is Nothing, Left(editCouponData("ValidTo"), 10), DateAdd("yyyy", 1, Now())) %>" required>
                    </div>
                    <div class="form-group">
                        <label>限定品类</label>
                        <input type="text" name="category" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("ApplicableCategory"), "") %>" placeholder="留空=全场通用">
                    </div>
                    <div class="form-group">
                        <label>限定产品ID</label>
                        <input type="number" name="product_id" value="<%= IIf(Not editCouponData Is Nothing And Not IsNull(editCouponData("ApplicableProductID")), editCouponData("ApplicableProductID"), "") %>" placeholder="留空=全场通用">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group" style="grid-column: span 2;">
                        <label>描述</label>
                        <input type="text" name="description" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("Description"), "") %>" placeholder="简短描述">
                    </div>
                    <div class="form-group" style="grid-column: span 2;">
                        <label>使用条款</label>
                        <input type="text" name="terms" value="<%= IIf(Not editCouponData Is Nothing, editCouponData("Terms"), "") %>" placeholder="详细条款">
                    </div>
                </div>
                <div style="margin-top:12px;display:flex;gap:8px;">
                    <button type="submit" name="action" value="save" class="btn btn-primary">
                        <i class="fas fa-save"></i> <% If Not editCouponData Is Nothing Then %>更新<% Else %>创建<% End If %>优惠券
                    </button>
                    <% If Not editCouponData Is Nothing Then %>
                    <a href="coupon_management.asp" class="btn" style="background:#555;color:#fff;">取消编辑</a>
                    <% End If %>
                </div>
            </form>
        </div>
        
        <!-- 优惠券列表 -->
        <div class="panel">
            <h3><i class="fas fa-list"></i> 优惠券列表</h3>
            <table class="coupon-table">
                <thead>
                    <tr><th>ID</th><th>优惠码</th><th>名称</th><th>类型</th><th>面额</th><th>门槛</th><th>发行/已用</th><th>有效期</th><th>状态</th><th>操作</th></tr>
                </thead>
                <tbody>
                    <% If Not rsAllCoupons Is Nothing Then
                        Do While Not rsAllCoupons.EOF
                            Dim cid, ccode, cname, ctype, cval, cmin, cqty, cuqty, cfrom, cto, cactive
                            cid = rsAllCoupons("CouponID")
                            ccode = rsAllCoupons("CouponCode")
                            cname = rsAllCoupons("CouponName")
                            ctype = rsAllCoupons("CouponType")
                            cval = CDbl(rsAllCoupons("DiscountValue"))
                            cmin = CDbl(rsAllCoupons("MinSpend"))
                            cqty = rsAllCoupons("TotalQty")
                            cuqty = rsAllCoupons("UsedQty")
                            cfrom = rsAllCoupons("ValidFrom")
                            cto = rsAllCoupons("ValidTo")
                            cactive = CBool(rsAllCoupons("IsActive"))
                            
                            Dim tBadge
                            Select Case ctype
                                Case "fixed": tBadge = "<span class='badge badge-fixed'>满减</span>"
                                Case "percentage": tBadge = "<span class='badge badge-percent'>折扣</span>"
                                Case "free_shipping": tBadge = "<span class='badge badge-freeship'>免邮</span>"
                                Case Else: tBadge = "<span class='badge badge-inactive'>" & ctype & "</span>"
                            End Select
                    %>
                    <tr>
                        <td><%= cid %></td>
                        <td><strong style="color:#ff8f00;"><%= ccode %></strong></td>
                        <td><%= cname %></td>
                        <td><%= tBadge %></td>
                        <td><%= FormatNumber(cval, 0) %></td>
                        <td><%= FormatNumber(cmin, 0) %></td>
                        <td><%= cuqty %>/<%= IIf(cqty = 0, "∞", cqty) %></td>
                        <td style="font-size:11px;"><%= Left(cfrom, 10) %><br>~<%= Left(cto, 10) %></td>
                        <td><span class="badge <% If cactive Then %>badge-active<% Else %>badge-inactive<% End If %>"><% If cactive Then %>启用<% Else %>禁用<% End If %></span></td>
                        <td>
                            <a href="?edit_id=<%= cid %>" class="btn btn-sm btn-primary"><i class="fas fa-edit"></i> 编辑</a>
                            <a href="?action=delete&id=<%= cid %>" class="btn btn-sm btn-danger" onclick="return confirm('确认删除 <%= ccode %>？')"><i class="fas fa-trash"></i></a>
                        </td>
                    </tr>
                    <%
                            rsAllCoupons.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
<%
If Not rsAllCoupons Is Nothing Then rsAllCoupons.Close : Set rsAllCoupons = Nothing
If Not editCouponData Is Nothing Then
    editCouponData.Close
    Set editCouponData = Nothing
End If
Call CloseConnection()
%>
