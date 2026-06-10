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

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' 处理POST请求
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        errorMsg = "安全验证失败，请刷新页面重试"
    Else
        Dim action
        action = Request.Form("action")
        
        Select Case action
            Case "create"
                ' 创建优惠券
                Dim couponCode, discountType, discountValue, minPurchaseC, startDateC, endDateC, usageLimit
                couponCode = Trim(Request.Form("coupon_code"))
                discountType = Request.Form("discount_type")
                discountValue = Request.Form("discount_value")
                minPurchaseC = Request.Form("min_purchase")
                startDateC = Request.Form("start_date")
                endDateC = Request.Form("end_date")
                usageLimit = Request.Form("usage_limit")
                
                If couponCode = "" Then
                    errorMsg = "请输入优惠券代码"
                ElseIf discountValue = "" Then
                    errorMsg = "请输入优惠值"
                Else
                    Dim sqlCreate, startDateSQL, endDateSQL
                    ' 处理日期格式
                    If startDateC <> "" Then
                        startDateSQL = "'" & startDateC & "'"
                    Else
                        startDateSQL = "NULL"
                    End If
                    If endDateC <> "" Then
                        endDateSQL = "'" & endDateC & "'"
                    Else
                        endDateSQL = "NULL"
                    End If
                    
                    sqlCreate = "INSERT INTO Coupons (CouponCode, DiscountType, DiscountValue, MinPurchase, StartDate, EndDate, UsageLimit, UsedCount, IsActive) " & _
                                "VALUES ('" & SafeSQL(couponCode) & "', '" & discountType & "', " & discountValue & ", " & _
                                IIf(minPurchaseC <> "", minPurchaseC, "0") & ", " & startDateSQL & ", " & endDateSQL & ", " & _
                                IIf(usageLimit <> "", usageLimit, "0") & ", 0, 1)"
                    
                    If ExecuteNonQuery(sqlCreate) Then
                        Call LogAdminAction("创建优惠券", "operation", "Coupons", "", couponCode)
                        successMsg = "优惠券创建成功"
                    Else
                        errorMsg = "创建失败: " & Session("LastDBError")
                    End If
                End If
                
            Case "update"
                ' 更新优惠券
                Dim couponId, isActiveU
                couponId = Request.Form("coupon_id")
                discountType = Request.Form("discount_type")
                discountValue = Request.Form("discount_value")
                minPurchaseC = Request.Form("min_purchase")
                startDateC = Request.Form("start_date")
                endDateC = Request.Form("end_date")
                usageLimit = Request.Form("usage_limit")
                isActiveU = (Request.Form("is_active") = "1")
                
                If couponId <> "" And IsNumeric(couponId) Then
                    Dim sqlUpdate
                    ' 处理日期格式
                    If startDateC <> "" Then
                        startDateSQL = "'" & startDateC & "'"
                    Else
                        startDateSQL = "NULL"
                    End If
                    If endDateC <> "" Then
                        endDateSQL = "'" & endDateC & "'"
                    Else
                        endDateSQL = "NULL"
                    End If
                    
                    sqlUpdate = "UPDATE Coupons SET " & _
                                "DiscountType = '" & discountType & "', " & _
                                "DiscountValue = " & discountValue & ", " & _
                                "MinPurchase = " & IIf(minPurchaseC <> "", minPurchaseC, "0") & ", " & _
                                "StartDate = " & startDateSQL & ", " & _
                                "EndDate = " & endDateSQL & ", " & _
                                "UsageLimit = " & IIf(usageLimit <> "", usageLimit, "0") & ", " & _
                                "IsActive = " & IIf(isActiveU, "1", "0") & " " & _
                                "WHERE CouponID = " & couponId
                    
                    If ExecuteNonQuery(sqlUpdate) Then
                        Call LogAdminAction("编辑优惠券", "operation", "Coupons", couponId, "")
                        successMsg = "优惠券更新成功"
                    Else
                        errorMsg = "更新失败: " & Session("LastDBError")
                    End If
                End If
                
            Case "delete"
                ' 删除优惠券
                couponId = Request.Form("coupon_id")
                If couponId <> "" And IsNumeric(couponId) Then
                    Dim sqlDelete
                    sqlDelete = "DELETE FROM Coupons WHERE CouponID = " & couponId
                    
                    If ExecuteNonQuery(sqlDelete) Then
                        Call LogAdminAction("删除优惠券", "operation", "Coupons", couponId, "")
                        successMsg = "优惠券删除成功"
                    Else
                        errorMsg = "删除失败: " & Session("LastDBError")
                    End If
                End If
        End Select
    End If
End If

' 获取优惠券列表
Dim rsCoupons
Set rsCoupons = ExecuteQuery("SELECT * FROM Coupons ORDER BY CouponID DESC")

' 获取统计
Dim totalCoupons, activeCoupons, totalUsed
totalCoupons = GetScalar("SELECT COUNT(*) FROM Coupons")
activeCoupons = GetScalar("SELECT COUNT(*) FROM Coupons WHERE EndDate >= CAST(GETDATE() AS DATE) AND IsActive = 1")
totalUsed = GetScalar("SELECT SUM(UsedCount) FROM Coupons")

Call LogAdminAction("查看优惠券管理", "operation", "Coupons", "", "")
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
        .stats-cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); text-align: center; }
        .stat-card i { font-size: 36px; color: #667eea; margin-bottom: 10px; }
        .stat-card h3 { font-size: 32px; margin: 10px 0; color: #333; }
        .stat-card p { color: #666; margin: 0; }
        .coupons-table { width: 100%; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .coupons-table th { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; text-align: left; }
        .coupons-table td { padding: 15px; border-bottom: 1px solid #f0f0f0; }
        .coupons-table tr:hover { background: #f8f9fa; }
        .coupon-code { font-family: monospace; font-weight: 600; color: #667eea; }
        .discount-type { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .type-percentage { background: #e3f2fd; color: #1976d2; }
        .type-fixed { background: #e8f5e9; color: #2e7d32; }
        .coupon-status { padding: 4px 12px; border-radius: 12px; font-size: 12px; }
        .status-active { background: #e8f5e9; color: #2e7d32; }
        .status-inactive { background: #ffebee; color: #c62828; }
        .btn-small { padding: 6px 15px; border-radius: 6px; text-decoration: none; font-size: 13px; cursor: pointer; border: none; }
        .btn-edit { background: #667eea; color: white; }
        .btn-delete { background: #e74c3c; color: white; }
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; }
        .modal-content { background: white; width: 90%; max-width: 500px; margin: 50px auto; padding: 30px; border-radius: 12px; }
        .modal-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #f0f0f0; }
        .modal-header h3 { margin: 0; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; margin-bottom: 5px; font-weight: 500; }
        .form-control { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 6px; box-sizing: border-box; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .alert { padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .alert-error { background: #ffebee; color: #c62828; border: 1px solid #ffcdd2; }
        .alert-success { background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-ticket-alt"></i> 优惠券管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>优惠券管理</span>
            </div>
        </div>
        
        <% If errorMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-exclamation-circle"></i> <%= errorMsg %></div>
        <% End If %>
        <% If successMsg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= successMsg %></div>
        <% End If %>
        
        <div class="stats-cards">
            <div class="stat-card">
                <i class="fas fa-ticket-alt"></i>
                <h3><%= totalCoupons %></h3>
                <p>总优惠券数</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-check-circle"></i>
                <h3><%= activeCoupons %></h3>
                <p>活跃优惠券</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-chart-bar"></i>
                <h3><%= totalUsed %></h3>
                <p>已使用次数</p>
            </div>
        </div>
        
        <div style="margin-bottom: 20px;">
            <button onclick="showCreateModal()" class="admin-btn admin-btn-primary">
                <i class="fas fa-plus"></i> 新建优惠券
            </button>
        </div>
        
        <table class="coupons-table">
            <thead>
                <tr>
                    <th>优惠券代码</th>
                    <th>类型</th>
                    <th>优惠值</th>
                    <th>最低消费</th>
                    <th>有效期</th>
                    <th>使用限制</th>
                    <th>已使用</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsCoupons Is Nothing Then %>
                <% Do While Not rsCoupons.EOF %>
                <tr>
                    <td class="coupon-code"><%= rsCoupons("CouponCode") %></td>
                    <td>
                        <% If rsCoupons("DiscountType") = "percentage" Then %>
                        <span class="discount-type type-percentage">百分比</span>
                        <% Else %>
                        <span class="discount-type type-fixed">固定金额</span>
                        <% End If %>
                    </td>
                    <td>
                        <% If rsCoupons("DiscountType") = "percentage" Then %>
                        <%= rsCoupons("DiscountValue") %>%
                        <% Else %>
                        ¥<%= FormatNumber(CDbl("0" & rsCoupons("DiscountValue")), 2) %>
                        <% End If %>
                    </td>
                    <td>¥<%= FormatNumber(CDbl("0" & rsCoupons("MinPurchase")), 2) %></td>
                    <td style="font-size: 13px; color: #666;">
                        <%= FormatDateField(rsCoupons("StartDate")) %><br>~ <%= FormatDateField(rsCoupons("EndDate")) %>
                    </td>
                    <td><%= IIf(rsCoupons("UsageLimit") > 0, rsCoupons("UsageLimit"), "无限制") %></td>
                    <td><%= rsCoupons("UsedCount") %></td>
                    <td>
                        <% If rsCoupons("IsActive") = True And rsCoupons("EndDate") >= Date() Then %>
                        <span class="coupon-status status-active">有效</span>
                        <% Else %>
                        <span class="coupon-status status-inactive">无效</span>
                        <% End If %>
                    </td>
                    <td>
                        <button onclick="showEditModal(<%= rsCoupons("CouponID") %>, '<%= SafeOutput(rsCoupons("CouponCode")) %>', '<%= rsCoupons("DiscountType") %>', <%= CDbl("0" & rsCoupons("DiscountValue")) %>, <%= CDbl("0" & rsCoupons("MinPurchase")) %>, '<%= FormatDateField(rsCoupons("StartDate")) %>', '<%= FormatDateField(rsCoupons("EndDate")) %>', <%= rsCoupons("UsageLimit") %>, <%= IIf(rsCoupons("IsActive")=True, "true", "false") %>)" class="btn-small btn-edit">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <form method="post" style="display:inline;" onsubmit="return confirm('确定要删除此优惠券吗？');">
                            <%= GetCSRFTokenField() %>
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="coupon_id" value="<%= rsCoupons("CouponID") %>">
                            <button type="submit" class="btn-small btn-delete">
                                <i class="fas fa-trash"></i> 删除
                            </button>
                        </form>
                    </td>
                </tr>
                <% rsCoupons.MoveNext %>
                <% Loop %>
                <% rsCoupons.Close %>
                <% End If %>
            </tbody>
        </table>
    </div>
    
    <!-- 创建优惠券弹窗 -->
    <div id="createModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-plus"></i> 新建优惠券</h3>
            </div>
            <form method="post">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="create">
                
                <div class="form-group">
                    <label>优惠券代码 *</label>
                    <input type="text" name="coupon_code" class="form-control" required placeholder="例如: SUMMER2024">
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>优惠类型</label>
                        <select name="discount_type" class="form-control">
                            <option value="percentage">百分比折扣</option>
                            <option value="fixed">固定金额</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>优惠值 *</label>
                        <input type="number" name="discount_value" class="form-control" step="0.01" required placeholder="例如: 20 或 10.5">
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>最低消费金额</label>
                        <input type="number" name="min_purchase" class="form-control" step="0.01" value="0" placeholder="0表示无限制">
                    </div>
                    <div class="form-group">
                        <label>使用次数限制</label>
                        <input type="number" name="usage_limit" class="form-control" value="0" placeholder="0表示无限制">
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>开始日期</label>
                        <input type="date" name="start_date" class="form-control" value="<%= FormatDateField(Date()) %>">
                    </div>
                    <div class="form-group">
                        <label>结束日期</label>
                        <input type="date" name="end_date" class="form-control" value="<%= FormatDateField(DateAdd("d", 30, Date())) %>">
                    </div>
                </div>
                
                <div style="display: flex; gap: 10px; margin-top: 20px;">
                    <button type="submit" class="admin-btn admin-btn-primary">创建</button>
                    <button type="button" onclick="hideCreateModal()" class="admin-btn admin-btn-secondary">取消</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 编辑优惠券弹窗 -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-edit"></i> 编辑优惠券</h3>
            </div>
            <form method="post">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="update">
                <input type="hidden" name="coupon_id" id="edit_coupon_id">
                
                <div class="form-group">
                    <label>优惠券代码</label>
                    <input type="text" id="edit_coupon_code" class="form-control" disabled>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>优惠类型</label>
                        <select name="discount_type" id="edit_discount_type" class="form-control">
                            <option value="percentage">百分比折扣</option>
                            <option value="fixed">固定金额</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>优惠值 *</label>
                        <input type="number" name="discount_value" id="edit_discount_value" class="form-control" step="0.01" required>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>最低消费金额</label>
                        <input type="number" name="min_purchase" id="edit_min_purchase" class="form-control" step="0.01">
                    </div>
                    <div class="form-group">
                        <label>使用次数限制</label>
                        <input type="number" name="usage_limit" id="edit_usage_limit" class="form-control">
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>开始日期</label>
                        <input type="date" name="start_date" id="edit_start_date" class="form-control">
                    </div>
                    <div class="form-group">
                        <label>结束日期</label>
                        <input type="date" name="end_date" id="edit_end_date" class="form-control">
                    </div>
                </div>
                
                <div class="form-group">
                    <label>
                        <input type="checkbox" name="is_active" id="edit_is_active" value="1"> 启用优惠券
                    </label>
                </div>
                
                <div style="display: flex; gap: 10px; margin-top: 20px;">
                    <button type="submit" class="admin-btn admin-btn-primary">保存</button>
                    <button type="button" onclick="hideEditModal()" class="admin-btn admin-btn-secondary">取消</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function showCreateModal() {
            document.getElementById('createModal').style.display = 'block';
        }
        
        function hideCreateModal() {
            document.getElementById('createModal').style.display = 'none';
        }
        
        function showEditModal(id, code, type, value, minPurchase, startDate, endDate, usageLimit, isActive) {
            document.getElementById('edit_coupon_id').value = id;
            document.getElementById('edit_coupon_code').value = code;
            document.getElementById('edit_discount_type').value = type;
            document.getElementById('edit_discount_value').value = value;
            document.getElementById('edit_min_purchase').value = minPurchase;
            document.getElementById('edit_start_date').value = startDate;
            document.getElementById('edit_end_date').value = endDate;
            document.getElementById('edit_usage_limit').value = usageLimit;
            document.getElementById('edit_is_active').checked = isActive;
            document.getElementById('editModal').style.display = 'block';
        }
        
        function hideEditModal() {
            document.getElementById('editModal').style.display = 'none';
        }
        
        // 点击弹窗外部关闭
        window.onclick = function(event) {
            if (event.target.className === 'modal') {
                event.target.style.display = 'none';
            }
        }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
