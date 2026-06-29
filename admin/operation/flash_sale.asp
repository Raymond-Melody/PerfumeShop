<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/dal.asp"-->
<%
Call OpenConnection()

Dim fsAction, fsActionMsg, fsActionResult
fsAction = Request.QueryString("action")
If fsAction = "" Then fsAction = Request.Form("action")
fsActionMsg = ""
fsActionResult = True

' 新增/编辑秒杀
If fsAction = "save" And Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim fsEditId, fsProductID, fsPrice, fsStock, fsLimit, fsStart, fsEnd, fsSort
    fsEditId = Request.Form("edit_id")
    fsProductID = CLng(Request.Form("product_id"))
    fsPrice = CDbl(Request.Form("flash_price"))
    fsStock = CLng(Request.Form("stock"))
    fsLimit = CLng(Request.Form("limit_per_user"))
    If fsLimit <= 0 Then fsLimit = 1
    fsStart = Request.Form("start_time")
    fsEnd = Request.Form("end_time")
    fsSort = Request.Form("sort_order")
    If fsSort = "" Or Not IsNumeric(fsSort) Then fsSort = 0

    If fsProductID <= 0 Or fsPrice <= 0 Or fsStock <= 0 Then
        fsActionMsg = "请完整填写必填字段"
        fsActionResult = False
    ElseIf CDate(fsEnd) <= CDate(fsStart) Then
        fsActionMsg = "结束时间必须晚于开始时间"
        fsActionResult = False
    Else
        On Error Resume Next
        If fsEditId <> "" And IsNumeric(fsEditId) Then
            conn.Execute "UPDATE FlashSale SET ProductID=" & fsProductID & ", FlashPrice=" & fsPrice & _
                       ", Stock=" & fsStock & ", LimitPerUser=" & fsLimit & _
                       ", StartTime='" & SafeSQL(fsStart) & "', EndTime='" & SafeSQL(fsEnd) & _
                       "', SortOrder=" & fsSort & " WHERE FlashSaleID=" & fsEditId
            If Err.Number = 0 Then fsActionMsg = "更新成功" Else fsActionMsg = "更新失败: " & Err.Description : fsActionResult = False
        Else
            conn.Execute "INSERT INTO FlashSale (ProductID, FlashPrice, Stock, LimitPerUser, StartTime, EndTime, SortOrder) VALUES (" & _
                       fsProductID & "," & fsPrice & "," & fsStock & "," & fsLimit & ",'" & _
                       SafeSQL(fsStart) & "','" & SafeSQL(fsEnd) & "'," & fsSort & ")"
            If Err.Number = 0 Then fsActionMsg = "创建成功" Else fsActionMsg = "创建失败: " & Err.Description : fsActionResult = False
        End If
        On Error GoTo 0
    End If
End If

' 删除
If fsAction = "delete" Then
    Dim fsDelId : fsDelId = Request.QueryString("id")
    If IsNumeric(fsDelId) Then
        conn.Execute "DELETE FROM FlashSale WHERE FlashSaleID = " & fsDelId
        fsActionMsg = "已删除"
    End If
End If

' 切换状态
If fsAction = "toggle" Then
    Dim fsTogId : fsTogId = Request.QueryString("id")
    If IsNumeric(fsTogId) Then
        conn.Execute "UPDATE FlashSale SET IsActive = CASE WHEN IsActive = 1 THEN 0 ELSE 1 END WHERE FlashSaleID = " & fsTogId
        fsActionMsg = "状态已切换"
    End If
End If

' 统计数据
Dim fsTotal, fsActive, fsUpcoming, fsExpired
Dim nowStr : nowStr = Now()
fsTotal = GetScalar("SELECT COUNT(*) FROM FlashSale")
fsActive = GetScalar("SELECT COUNT(*) FROM FlashSale WHERE IsActive = 1 AND '" & nowStr & "' >= StartTime AND '" & nowStr & "' <= EndTime")
fsUpcoming = GetScalar("SELECT COUNT(*) FROM FlashSale WHERE IsActive = 1 AND StartTime > '" & nowStr & "'")
fsExpired = GetScalar("SELECT COUNT(*) FROM FlashSale WHERE EndTime < '" & nowStr & "'")

' 获取编辑中的秒杀
Dim fsEditData, fsEditId
fsEditId = Request.QueryString("edit_id")
Set fsEditData = Nothing
If fsEditId <> "" And IsNumeric(fsEditId) Then
    Set fsEditData = conn.Execute("SELECT * FROM FlashSale WHERE FlashSaleID = " & fsEditId)
End If

' 所有秒杀列表
Dim rsFlash : Set rsFlash = DAL_GetList("SELECT fs.*, p.ProductName, p.BasePrice FROM FlashSale fs INNER JOIN Products p ON fs.ProductID = p.ProductID ORDER BY fs.StartTime DESC", Null)

' 产品列表（用于下拉选择）
Dim rsProducts : Set rsProducts = DAL_GetList("SELECT ProductID, ProductName, BasePrice FROM Products WHERE IsActive <> 0 ORDER BY ProductName", Null)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>秒杀管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 24px; }
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .page-title { font-size: 22px; color: #fff; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #ff416c; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); padding: 20px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.05); text-align: center; }
        .stat-value { font-size: 28px; font-weight: bold; color: #ff416c; }
        .stat-label { font-size: 13px; color: #888; margin-top: 4px; }
        
        .panel { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; padding: 24px; border: 1px solid rgba(255,255,255,0.05); margin-bottom: 24px; }
        .panel h3 { color: #fff; margin: 0 0 16px; font-size: 18px; display: flex; align-items: center; gap: 8px; }
        .panel h3 i { color: #ff416c; }
        
        .form-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 12px; }
        .form-group { display: flex; flex-direction: column; gap: 4px; }
        .form-group label { font-size: 12px; color: #888; font-weight: 500; }
        .form-group input, .form-group select, .form-group textarea {
            padding: 8px 12px; border: 1px solid #3a3a4a; border-radius: 6px; background: #1a1a2e; color: #e0e0e0; font-size: 13px;
        }
        .form-group input:focus, .form-group select:focus { border-color: #ff416c; outline: none; }
        
        .btn { padding: 8px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; text-decoration: none; display: inline-block; }
        .btn-primary { background: linear-gradient(135deg, #ff416c, #ff4b2b); color: #fff; }
        .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(255,65,108,0.3); }
        .btn-danger { background: #c62828; color: #fff; }
        .btn-sm { padding: 4px 12px; font-size: 11px; }
        .btn-outline { background: transparent; border: 1px solid #555; color: #ccc; }
        
        .flash-table { width: 100%; border-collapse: collapse; }
        .flash-table th { text-align: left; padding: 10px 12px; background: rgba(0,0,0,0.2); color: #888; font-size: 11px; text-transform: uppercase; }
        .flash-table td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .flash-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; }
        .badge-active { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .badge-inactive { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        .badge-upcoming { background: rgba(33,150,243,0.2); color: #64B5F6; }
        .badge-expired { background: rgba(244,67,54,0.2); color: #ef9a9a; }
        
        .alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #ef9a9a; border: 1px solid rgba(244,67,54,0.3); }
        
        @media (max-width: 768px) { .stats-grid { grid-template-columns: repeat(2, 1fr); } .form-row { grid-template-columns: 1fr; } }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-bolt"></i> 秒杀活动管理</h2>
        </div>
        
        <% If fsActionMsg <> "" Then %>
        <div class="alert <% If fsActionResult Then %>alert-success<% Else %>alert-error<% End If %>">
            <i class="fas fa-<% If fsActionResult Then %>check-circle<% Else %>exclamation-circle<% End If %>"></i> <%= fsActionMsg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value"><%= fsTotal %></div>
                <div class="stat-label">全部活动</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= fsActive %></div>
                <div class="stat-label">进行中</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= fsUpcoming %></div>
                <div class="stat-label">即将开始</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%= fsExpired %></div>
                <div class="stat-label">已结束</div>
            </div>
        </div>
        
        <!-- 创建/编辑表单 -->
        <div class="panel">
            <h3><i class="fas fa-<% If Not fsEditData Is Nothing And Not fsEditData.EOF Then %>edit<% Else %>plus-circle<% End If %>"></i> <% If Not fsEditData Is Nothing And Not fsEditData.EOF Then %>编辑秒杀活动<% Else %>创建秒杀活动<% End If %></h3>
            <form method="post">
                <% If Not fsEditData Is Nothing And Not fsEditData.EOF Then %>
                <input type="hidden" name="edit_id" value="<%= fsEditData("FlashSaleID") %>">
                <% End If %>
                <div class="form-row">
                    <div class="form-group">
                        <label>选择产品 *</label>
                        <select name="product_id" required>
                            <option value="">-- 请选择产品 --</option>
                            <%
                            If Not rsProducts Is Nothing Then
                                Do While Not rsProducts.EOF
                                    Dim pSelID : pSelID = rsProducts("ProductID")
                                    Dim pSelName : pSelName = rsProducts("ProductName")
                                    Dim pSelPrice : pSelPrice = rsProducts("BasePrice")
                                    Dim pSelSelected : pSelSelected = ""
                                    If Not fsEditData Is Nothing And Not fsEditData.EOF Then
                                        If CLng(fsEditData("ProductID")) = CLng(pSelID) Then pSelSelected = " selected"
                                    End If
                            %>
                            <option value="<%= pSelID %>"<%= pSelSelected %>><%= Server.HTMLEncode(pSelName) %> (&yen;<%= FormatNumber(pSelPrice, 2) %>)</option>
                            <%
                                    rsProducts.MoveNext
                                Loop
                                rsProducts.MoveFirst
                            End If
                            %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>秒杀价 *</label>
                        <input type="number" name="flash_price" step="0.01" value="<%= IIf(Not fsEditData Is Nothing And Not fsEditData.EOF, fsEditData("FlashPrice"), "") %>" placeholder="如: 99.00" required>
                    </div>
                    <div class="form-group">
                        <label>秒杀库存 *</label>
                        <input type="number" name="stock" value="<%= IIf(Not fsEditData Is Nothing And Not fsEditData.EOF, fsEditData("Stock"), "100") %>" required>
                    </div>
                    <div class="form-group">
                        <label>每人限购</label>
                        <input type="number" name="limit_per_user" value="<%= IIf(Not fsEditData Is Nothing And Not fsEditData.EOF, fsEditData("LimitPerUser"), "1") %>" min="1">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>开始时间 *</label>
                        <input type="datetime-local" name="start_time" value="<%= IIf(Not fsEditData Is Nothing And Not fsEditData.EOF, FormatDateTime(fsEditData("StartTime"), 0), FormatDateTime(Now(), 0)) %>" required>
                    </div>
                    <div class="form-group">
                        <label>结束时间 *</label>
                        <input type="datetime-local" name="end_time" value="<%= IIf(Not fsEditData Is Nothing And Not fsEditData.EOF, FormatDateTime(fsEditData("EndTime"), 0), FormatDateTime(DateAdd("d", 1, Now()), 0)) %>" required>
                    </div>
                    <div class="form-group">
                        <label>排序</label>
                        <input type="number" name="sort_order" value="<%= IIf(Not fsEditData Is Nothing And Not fsEditData.EOF, fsEditData("SortOrder"), "0") %>" min="0">
                    </div>
                </div>
                <div style="margin-top:12px;display:flex;gap:8px;">
                    <button type="submit" name="action" value="save" class="btn btn-primary">
                        <i class="fas fa-save"></i> <% If Not fsEditData Is Nothing And Not fsEditData.EOF Then %>更新<% Else %>创建<% End If %>秒杀活动
                    </button>
                    <% If Not fsEditData Is Nothing And Not fsEditData.EOF Then %>
                    <a href="flash_sale.asp" class="btn btn-outline">取消编辑</a>
                    <% End If %>
                </div>
            </form>
        </div>
        
        <!-- 秒杀列表 -->
        <div class="panel">
            <h3><i class="fas fa-list"></i> 秒杀活动列表</h3>
            <table class="flash-table">
                <thead>
                    <tr>
                        <th>ID</th><th>产品</th><th>原价</th><th>秒杀价</th><th>库存(剩余/总量)</th><th>已售</th><th>限购</th><th>时间</th><th>状态</th><th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If Not rsFlash Is Nothing Then
                        Do While Not rsFlash.EOF
                            Dim fID, fName, fBasePrice, fFlashPrice, fStock, fSold, fLimit, fStart, fEnd, fActive
                            fID = rsFlash("FlashSaleID")
                            fName = rsFlash("ProductName")
                            fBasePrice = rsFlash("BasePrice")
                            fFlashPrice = rsFlash("FlashPrice")
                            fStock = rsFlash("Stock")
                            fSold = rsFlash("SoldCount")
                            fLimit = rsFlash("LimitPerUser")
                            fStart = rsFlash("StartTime")
                            fEnd = rsFlash("EndTime")
                            fActive = CBool(rsFlash("IsActive"))
                            
                            Dim fStatus, fBadge
                            If Not fActive Then
                                fStatus = "已禁用" : fBadge = "badge-inactive"
                            ElseIf CDate(nowStr) < CDate(fStart) Then
                                fStatus = "即将开始" : fBadge = "badge-upcoming"
                            ElseIf CDate(nowStr) > CDate(fEnd) Then
                                fStatus = "已结束" : fBadge = "badge-expired"
                            Else
                                fStatus = "进行中" : fBadge = "badge-active"
                            End If
                            
                            Dim remain : remain = CLng(fStock) - CLng(fSold)
                            If remain < 0 Then remain = 0
                    %>
                    <tr>
                        <td><%= fID %></td>
                        <td><strong><%= Server.HTMLEncode(fName) %></strong></td>
                        <td>&yen;<%= FormatNumber(fBasePrice, 2) %></td>
                        <td style="color:#ff416c;font-weight:700;">&yen;<%= FormatNumber(fFlashPrice, 2) %></td>
                        <td><%= remain %>/<%= fStock %></td>
                        <td><%= fSold %></td>
                        <td><%= fLimit %></td>
                        <td style="font-size:11px;"><%= FormatDateTime(fStart, 0) %><br>~<%= FormatDateTime(fEnd, 0) %></td>
                        <td><span class="badge <%= fBadge %>"><%= fStatus %></span></td>
                        <td>
                            <a href="?edit_id=<%= fID %>" class="btn btn-sm btn-primary"><i class="fas fa-edit"></i></a>
                            <a href="?action=toggle&id=<%= fID %>" class="btn btn-sm btn-outline"><i class="fas fa-power-off"></i></a>
                            <a href="?action=delete&id=<%= fID %>" class="btn btn-sm btn-danger" onclick="return confirm('确认删除？')"><i class="fas fa-trash"></i></a>
                        </td>
                    </tr>
                    <%
                            rsFlash.MoveNext
                        Loop
                    End If
                    %>
                </tbody>
            </table>
            <% If rsFlash Is Nothing Or rsFlash.EOF Then %>
            <div class="empty-state" style="padding:40px;text-align:center;color:#888;">
                <i class="fas fa-bolt" style="font-size:2rem;margin-bottom:10px;display:block;"></i>
                <p>暂无秒杀活动，点击上方表单创建</p>
            </div>
            <% End If %>
        </div>
    </div>
</body>
</html>
<%
If Not rsFlash Is Nothing Then
    If rsFlash.State = 1 Then rsFlash.Close
    Set rsFlash = Nothing
End If
If Not rsProducts Is Nothing Then
    If rsProducts.State = 1 Then rsProducts.Close
    Set rsProducts = Nothing
End If
If Not fsEditData Is Nothing Then
    If fsEditData.State = 1 Then fsEditData.Close
    Set fsEditData = Nothing
End If
Call CloseConnection()
%>
