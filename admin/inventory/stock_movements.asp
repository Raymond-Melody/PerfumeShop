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

' 确保 StockMovements 表存在
On Error Resume Next
conn.Execute "IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='StockMovements') CREATE TABLE StockMovements (MovementID INT IDENTITY(1,1) PRIMARY KEY, ItemType NVARCHAR(30), ItemID INT, ItemName NVARCHAR(200), ItemCode NVARCHAR(100), MovementType NVARCHAR(20), Quantity DECIMAL(12,2), BeforeQty DECIMAL(12,2), AfterQty DECIMAL(12,2), Unit NVARCHAR(20), ReferenceNo NVARCHAR(100), Notes NVARCHAR(500), CreatedBy NVARCHAR(50), CreatedAt DATETIME DEFAULT GETDATE())"
Err.Clear
On Error GoTo 0

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

' ========== 筛选条件 ==========
Dim smType : smType = Request.QueryString("type")
Dim smStart : smStart = Request.QueryString("start_date")
Dim smEnd : smEnd = Request.QueryString("end_date")
Dim smSearch : smSearch = Request.QueryString("search")
Dim smPage : smPage = Request.QueryString("page")
If smPage = "" Or Not IsNumeric(smPage) Then smPage = 1 Else smPage = CInt(smPage)

Const smPageSize = 20

' 构建 WHERE 条件
Dim smWhere : smWhere = " WHERE 1=1"

If smType <> "" And smType <> "all" Then
    smWhere = smWhere & " AND MovementType='" & SafeSQL(smType) & "'"
End If

If smStart <> "" Then
    smWhere = smWhere & " AND CreatedAt >= '" & SafeSQL(smStart) & "'"
End If

If smEnd <> "" Then
    smWhere = smWhere & " AND CreatedAt < '" & SafeSQL(smEnd) & " 23:59:59'"
End If

If smSearch <> "" Then
    smWhere = smWhere & " AND (ItemName LIKE '%" & SafeSQL(smSearch) & "%' OR ItemCode LIKE '%" & SafeSQL(smSearch) & "%' OR ReferenceNo LIKE '%" & SafeSQL(smSearch) & "%')"
End If

' ========== 统计数据 ==========
Dim smTotal
smTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM StockMovements" & smWhere))

' 各种类型统计
Dim inCount, outCount, adjustCount, transferCount
inCount = SafeNum(GetScalar("SELECT COUNT(*) FROM StockMovements WHERE MovementType='In'"))
outCount = SafeNum(GetScalar("SELECT COUNT(*) FROM StockMovements WHERE MovementType='Out'"))
adjustCount = SafeNum(GetScalar("SELECT COUNT(*) FROM StockMovements WHERE MovementType='Adjust'"))
transferCount = SafeNum(GetScalar("SELECT COUNT(*) FROM StockMovements WHERE MovementType='Transfer'"))

' 本月入库/出库
Dim monthInCount, monthOutCount
monthInCount = SafeNum(GetScalar("SELECT COUNT(*) FROM StockMovements WHERE MovementType='In' AND CreatedAt >= DATEADD(month,-1,GETDATE())"))
monthOutCount = SafeNum(GetScalar("SELECT COUNT(*) FROM StockMovements WHERE MovementType='Out' AND CreatedAt >= DATEADD(month,-1,GETDATE())"))

' ========== 分页查询 ==========
Dim smOffset : smOffset = (smPage - 1) * smPageSize
Dim rsSM
On Error Resume Next
Set rsSM = conn.Execute(_
    "SELECT * FROM StockMovements" & smWhere & _
    " ORDER BY CreatedAt DESC " & _
    "OFFSET " & smOffset & " ROWS FETCH NEXT " & smPageSize & " ROWS ONLY")
If Err.Number <> 0 Then
	Err.Clear
	Set rsSM = Nothing
End If
On Error GoTo 0

' 总页数
Dim smTotalPages : smTotalPages = 1
If smTotal > 0 Then smTotalPages = Int((smTotal - 1) / smPageSize) + 1

' 构建筛选URL参数
Function BuildFilterURL(excludePage)
    Dim u : u = "stock_movements.asp?"
    If smType <> "" Then u = u & "type=" & Server.URLEncode(smType) & "&"
    If smStart <> "" Then u = u & "start_date=" & Server.URLEncode(smStart) & "&"
    If smEnd <> "" Then u = u & "end_date=" & Server.URLEncode(smEnd) & "&"
    If smSearch <> "" Then u = u & "search=" & Server.URLEncode(smSearch) & "&"
    BuildFilterURL = u
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>库存流水 - 库存管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #607D8B; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 18px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; gap: 14px; }
        .stat-icon { width: 46px; height: 46px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 20px; flex-shrink: 0; }
        .stat-icon.in { background: rgba(76,175,80,0.15); color: #4CAF50; }
        .stat-icon.out { background: rgba(244,67,54,0.15); color: #f44336; }
        .stat-icon.adjust { background: rgba(255,152,0,0.15); color: #FF9800; }
        .stat-icon.transfer { background: rgba(33,150,243,0.15); color: #2196F3; }
        .stat-info .num { font-size: 20px; font-weight: bold; }
        .stat-info .label { font-size: 11px; color: #888; }
        .stat-info .month { font-size: 10px; color: #666; }
        
        .toolbar { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; background: #2d2d44; padding: 15px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.06); }
        .toolbar label { color: #888; font-size: 12px; white-space: nowrap; }
        .toolbar select, .toolbar input[type="text"], .toolbar input[type="date"] { background: #1a1a2e; color: #e0e0e0; border: 1px solid rgba(255,255,255,0.1); padding: 8px 12px; border-radius: 6px; font-size: 13px; }
        .toolbar select:focus, .toolbar input:focus { outline: none; border-color: #00BCD4; }
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); overflow: hidden; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(96,125,139,0.15); color: #90a4ae; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; white-space: nowrap; }
        td { padding: 11px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 13px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .type-badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .type-in { background: rgba(76,175,80,0.2); color: #81c784; }
        .type-out { background: rgba(244,67,54,0.2); color: #e57373; }
        .type-adjust { background: rgba(255,152,0,0.2); color: #ffb74d; }
        .type-transfer { background: rgba(33,150,243,0.2); color: #64b5f6; }
        
        .pagination { display: flex; justify-content: center; gap: 5px; padding: 20px 0; }
        .pagination a, .pagination span { padding: 8px 14px; border-radius: 6px; font-size: 13px; text-decoration: none; color: #b0b0b0; border: 1px solid rgba(255,255,255,0.08); transition: all 0.2s; }
        .pagination a:hover { border-color: #00BCD4; color: #00BCD4; }
        .pagination span.current { background: #00BCD4; color: #1a1a2e; border-color: #00BCD4; font-weight: 600; }
        .pagination span.disabled { color: #555; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .text-right { text-align: right; }
        .qty-positive { color: #4CAF50; }
        .qty-negative { color: #f44336; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-history"></i> 库存流水</h2>
            <p style="font-size:13px;color:#888;margin-top:5px;">全品类库存变动记录</p>
        </div>
        
        <!-- 快速统计 -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon in"><i class="fas fa-arrow-down"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#4CAF50;"><%=inCount%></div>
                    <div class="label">入库记录</div>
                    <div class="month">本月 <%=monthInCount%> 条</div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon out"><i class="fas fa-arrow-up"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#f44336;"><%=outCount%></div>
                    <div class="label">出库记录</div>
                    <div class="month">本月 <%=monthOutCount%> 条</div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon adjust"><i class="fas fa-balance-scale"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#FF9800;"><%=adjustCount%></div>
                    <div class="label">调整记录</div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon transfer"><i class="fas fa-exchange-alt"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#2196F3;"><%=transferCount%></div>
                    <div class="label">调拨记录</div>
                </div>
            </div>
        </div>
        
        <!-- 筛选工具栏 -->
        <form method="get" action="stock_movements.asp" class="toolbar">
            <label><i class="fas fa-filter"></i> 类型:</label>
            <select name="type" onchange="this.form.submit()">
                <option value="all" <%=IIF(smType="" Or smType="all","selected","")%>>全部类型</option>
                <option value="In" <%=IIF(smType="In","selected","")%>>入库</option>
                <option value="Out" <%=IIF(smType="Out","selected","")%>>出库</option>
                <option value="Adjust" <%=IIF(smType="Adjust","selected","")%>>调整</option>
                <option value="Transfer" <%=IIF(smType="Transfer","selected","")%>>调拨</option>
            </select>
            <label>起始:</label>
            <input type="date" name="start_date" value="<%=smStart%>" style="width:140px;">
            <label>截止:</label>
            <input type="date" name="end_date" value="<%=smEnd%>" style="width:140px;">
            <label>搜索:</label>
            <input type="text" name="search" value="<%=Server.HTMLEncode(smSearch)%>" placeholder="物品名称/编码/单号" style="width:180px;">
            <button type="submit" class="btn btn-primary"><i class="fas fa-search"></i> 查询</button>
            <a href="stock_movements.asp" class="btn btn-outline"><i class="fas fa-redo"></i> 重置</a>
        </form>
        
        <!-- 数据表格 -->
        <div class="card">
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>类型</th>
                            <th>物品名称</th>
                            <th>物品编码</th>
                            <th>变动数量</th>
                            <th>变动前</th>
                            <th>变动后</th>
                            <th>参考单号</th>
                            <th>操作人</th>
                            <th>操作时间</th>
                            <th>备注</th>
                        </tr>
                    </thead>
                    <tbody>
                    <%
                    Dim smRow : smRow = 0
                    If Not rsSM Is Nothing Then
                        Do While Not rsSM.EOF
                            smRow = smRow + 1
                            Dim smMvType : smMvType = rsSM("MovementType") & ""
                            Dim smQty : smQty = SafeNum(rsSM("Quantity"))
                            Dim smBefore : smBefore = SafeNum(rsSM("BeforeQty"))
                            Dim smAfter : smAfter = SafeNum(rsSM("AfterQty"))
                    %>
                        <tr>
                            <td style="color:#888;"><%=rsSM("MovementID")%></td>
                            <td>
                                <span class="type-badge <%=IIF(smMvType="In","type-in",IIF(smMvType="Out","type-out",IIF(smMvType="Adjust","type-adjust","type-transfer")))%>">
                                    <%=IIF(smMvType="In","入库",IIF(smMvType="Out","出库",IIF(smMvType="Adjust","调整","调拨")))%>
                                </span>
                            </td>
                            <td><strong><%=Server.HTMLEncode(rsSM("ItemName") & "")%></strong></td>
                            <td style="color:#888;"><%=rsSM("ItemCode") & ""%></td>
                            <td>
                                <span class="<%=IIF(smQty>=0,"qty-positive","qty-negative")%>">
                                    <%=IIF(smQty>=0,"+" & smQty,smQty)%> <%=rsSM("Unit") & ""%>
                                </span>
                            </td>
                            <td class="text-muted"><%=FormatNumber(smBefore,1)%></td>
                            <td><%=FormatNumber(smAfter,1)%></td>
                            <td style="color:#888;"><%=rsSM("ReferenceNo") & ""%></td>
                            <td><%=rsSM("CreatedBy") & ""%></td>
                            <td class="text-muted"><%=Left(rsSM("CreatedAt") & "", 19)%></td>
                            <td class="text-muted" style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="<%=Server.HTMLEncode(rsSM("Notes") & "")%>"><%=Server.HTMLEncode(rsSM("Notes") & "")%></td>
                        </tr>
                    <%
                            rsSM.MoveNext
                        Loop
                        rsSM.Close
                    End If
                    Set rsSM = Nothing
                    
                    If smRow = 0 Then
                    %>
                        <tr><td colspan="11" class="text-center text-muted" style="padding:40px;">
                            <i class="fas fa-inbox" style="font-size:32px;display:block;margin-bottom:10px;opacity:0.3;"></i>
                            暂无库存流水记录
                            <% If smSearch <> "" Or smType <> "" Then %>
                                <br><a href="stock_movements.asp" style="color:#00BCD4;text-decoration:none;">清除筛选条件</a>
                            <% End If %>
                        </td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- 分页 -->
        <% If smTotalPages > 1 Then
            Dim pi
        %>
        <div class="pagination">
            <% If smPage > 1 Then %>
                <a href="<%=BuildFilterURL(true)%>page=<%=smPage-1%>"><i class="fas fa-chevron-left"></i></a>
            <% Else %>
                <span class="disabled"><i class="fas fa-chevron-left"></i></span>
            <% End If %>
            
            <%
            Dim pStart, pEnd
            pStart = smPage - 2 : If pStart < 1 Then pStart = 1
            pEnd = smPage + 2 : If pEnd > smTotalPages Then pEnd = smTotalPages
            
            If pStart > 1 Then
            %>
                <a href="<%=BuildFilterURL(true)%>page=1">1</a>
                <% If pStart > 2 Then %><span class="disabled">...</span><% End If %>
            <% End If %>
            
            <% For pi = pStart To pEnd %>
                <% If pi = smPage Then %>
                    <span class="current"><%=pi%></span>
                <% Else %>
                    <a href="<%=BuildFilterURL(true)%>page=<%=pi%>"><%=pi%></a>
                <% End If %>
            <% Next %>
            
            <% If pEnd < smTotalPages Then %>
                <% If pEnd < smTotalPages - 1 Then %><span class="disabled">...</span><% End If %>
                <a href="<%=BuildFilterURL(true)%>page=<%=smTotalPages%>"><%=smTotalPages%></a>
            <% End If %>
            
            <% If smPage < smTotalPages Then %>
                <a href="<%=BuildFilterURL(true)%>page=<%=smPage+1%>"><i class="fas fa-chevron-right"></i></a>
            <% Else %>
                <span class="disabled"><i class="fas fa-chevron-right"></i></span>
            <% End If %>
            
            <span style="padding:8px 14px;font-size:12px;color:#888;">共 <%=smTotal%> 条 / <%=smTotalPages%> 页</span>
        </div>
        <% End If %>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
