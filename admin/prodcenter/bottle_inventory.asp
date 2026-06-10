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

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
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
    Set rs = Nothing
    GetScalar = val
End Function

' 自动添加库存字段（幂等）
On Error Resume Next
conn.Execute "SELECT StockQty FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD StockQty INT DEFAULT 0"
conn.Execute "SELECT SafetyStock FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD SafetyStock INT DEFAULT 0"
conn.Execute "SELECT UnitCost FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD UnitCost DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT BottleType FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD BottleType NVARCHAR(30) DEFAULT 'Standard'"
conn.Execute "SELECT CapacityML FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD CapacityML INT DEFAULT 50"
conn.Execute "SELECT UpdatedAt FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD UpdatedAt DATETIME"
On Error GoTo 0

' POST处理
Dim bAction, bottleId
bAction = Request.Form("action")
If bAction = "restock" Then
    bottleId = Request.Form("bottleId")
    Dim bQty : bQty = SafeNum(Request.Form("qty"))
    If IsNumeric(bottleId) And bQty > 0 Then
        conn.Execute "UPDATE BottleStyles SET StockQty = ISNULL(StockQty,0) + " & bQty & ", UpdatedAt = GETDATE() WHERE BottleID = " & CLng(bottleId)
        Response.Redirect "bottle_inventory.asp?msg=入库成功"
        Response.End
    End If
ElseIf bAction = "update_safety" Then
    bottleId = Request.Form("bottleId")
    Dim bSafety : bSafety = SafeNum(Request.Form("safetyStock"))
    If IsNumeric(bottleId) Then
        conn.Execute "UPDATE BottleStyles SET SafetyStock = " & bSafety & ", UpdatedAt = GETDATE() WHERE BottleID = " & CLng(bottleId)
        Response.Redirect "bottle_inventory.asp?msg=安全库存已更新"
        Response.End
    End If
ElseIf bAction = "update_cost" Then
    bottleId = Request.Form("bottleId")
    Dim bCost : bCost = SafeNum(Request.Form("unitCost"))
    If IsNumeric(bottleId) Then
        conn.Execute "UPDATE BottleStyles SET UnitCost = " & bCost & ", UpdatedAt = GETDATE() WHERE BottleID = " & CLng(bottleId)
        Response.Redirect "bottle_inventory.asp?msg=单价已更新"
        Response.End
    End If
End If

' 统计
Dim btTotal, btActive, btLowStock, btTotalQty, btValue
btTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleStyles"))
btActive = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleStyles WHERE IsActive=1"))
btLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleStyles WHERE ISNULL(StockQty,0) <= ISNULL(SafetyStock,0) AND ISNULL(SafetyStock,0) > 0"))
btTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(ISNULL(StockQty,0)),0) FROM BottleStyles"))
btValue = SafeNum(GetScalar("SELECT ISNULL(SUM(ISNULL(StockQty,0) * ISNULL(UnitCost,0)),0) FROM BottleStyles"))

Dim rsBottles
On Error Resume Next
Set rsBottles = conn.Execute("SELECT BottleID, BottleName, IsActive, ISNULL(StockQty,0) AS StockQty, ISNULL(SafetyStock,0) AS SafetyStock, ISNULL(UnitCost,0) AS UnitCost, ISNULL(PriceAddition,0) AS PriceAddition, ISNULL(BottleType,'Standard') AS BottleType, ISNULL(CapacityML,50) AS CapacityML FROM BottleStyles ORDER BY BottleName")
If Err.Number <> 0 Then
    Err.Clear
    Set rsBottles = Nothing
End If
' 注意：不要在此处使用 On Error GoTo 0，保持 Resume Next 以防字段不存在时崩溃
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>瓶子库存 - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #9C27B0; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: var(--accent); }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 16px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 11px; color: #888; margin-bottom: 6px; }
        .stat-card .stat-value { font-size: 24px; font-weight: 700; color: var(--accent); }
        .stat-card .stat-sub { font-size: 11px; color: #666; margin-top: 4px; }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(156,39,176,0.15); color: #ba68c8; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .stock-bar { height: 6px; border-radius: 3px; background: rgba(255,255,255,0.1); margin-top: 4px; overflow: hidden; }
        .stock-bar-fill { height: 100%; border-radius: 3px; transition: width 0.3s; }
        .stock-normal { background: #4CAF50; }
        .stock-low { background: #FF9800; }
        .stock-zero { background: #f44336; }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-active { background: rgba(76,175,80,0.15); color: #81c784; }
        .badge-inactive { background: rgba(158,158,158,0.15); color: #9e9e9e; }
        .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 999; align-items: center; justify-content: center; }
        .modal.active { display: flex; }
        .modal-content { background: #2d2d44; border-radius: 12px; padding: 24px; width: 420px; max-width: 90vw; border: 1px solid rgba(255,255,255,0.1); }
        .modal-title { font-size: 18px; color: #e0e0e0; margin-bottom: 16px; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input, .form-group select { width: 100%; padding: 9px 12px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; }
        .form-group input:focus { outline: none; border-color: var(--accent); }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
</style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-flask"></i> 瓶子库存管理</h2>
        </div>

        <% Dim msg : msg = Request.QueryString("msg")
        If msg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(msg)%></div>
        <% End If %>

        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-flask"></i> 瓶子款式</div>
                <div class="stat-value"><%=btActive%></div>
                <div class="stat-sub">总计 <%=btTotal%> 种</div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-cubes"></i> 库存总数</div>
                <div class="stat-value" style="color:#4CAF50;"><%=btTotalQty%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-exclamation-triangle"></i> 低库存预警</div>
                <div class="stat-value" style="color:<%=IIf(btLowStock>0,"#f44336","#4CAF50")%>;"><%=btLowStock%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-dollar-sign"></i> 库存价值</div>
                <div class="stat-value" style="color:#4CAF50;">¥<%=FormatNumber(btValue,2)%></div>
            </div>
        </div>

        <!-- 瓶子列表 -->
        <table class="data-table">
            <thead>
                <tr>
                    <th>瓶子名称</th>
                    <th>类型</th>
                    <th>容量(ml)</th>
                    <th>价格加成</th>
                    <th>库存量</th>
                    <th>安全库存</th>
                    <th>单价</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsBottles Is Nothing Then
                    Do While Not rsBottles.EOF
                        Dim bSId, bSName, bSActive, bSQty, bSSafety, bSCost, bSPrice, bSType, bSCap
                        bSId = rsBottles("BottleID")
                        bSName = rsBottles("BottleName") & ""
                        bSActive = rsBottles("IsActive") & ""
                        bSQty = SafeNum(rsBottles("StockQty"))
                        bSSafety = SafeNum(rsBottles("SafetyStock"))
                        bSCost = SafeNum(rsBottles("UnitCost"))
                        bSPrice = SafeNum(rsBottles("PriceAddition"))
                        bSType = rsBottles("BottleType") & ""
                        bSCap = SafeNum(rsBottles("CapacityML"))

                        Dim stockClass, stockPct
                        If bSQty <= 0 Then
                            stockClass = "stock-zero" : stockPct = 0
                        ElseIf bSSafety > 0 And bSQty <= bSSafety Then
                            stockClass = "stock-low" : stockPct = Int((bSQty / (bSSafety * 2)) * 100)
                        Else
                            stockClass = "stock-normal" : stockPct = 100
                        End If
                        If stockPct > 100 Then stockPct = 100
                %>
                <tr>
                    <td><strong><%=Server.HTMLEncode(bSName)%></strong></td>
                    <td style="color:#888;"><%=IIf(bSType="","Standard",Server.HTMLEncode(bSType))%></td>
                    <td><%=IIf(bSCap>0, bSCap & "ml", "-")%></td>
                    <td>¥<%=FormatNumber(bSPrice,2)%></td>
                    <td>
                        <%=bSQty%>
                        <div class="stock-bar"><div class="stock-bar-fill <%=stockClass%>" style="width:<%=stockPct%>%;"></div></div>
                    </td>
                    <td><%=IIf(bSSafety=0,"-",bSSafety)%></td>
                    <td>¥<%=FormatNumber(bSCost,2)%></td>
                    <td><span class="badge <%=IIf(bSActive="True","badge-active","badge-inactive")%>"><%=IIf(bSActive="True","启用","停用")%></span></td>
                    <td>
                        <button class="btn btn-sm btn-outline" onclick="openRestock(<%=bSId%>,'<%=Server.HTMLEncode(Replace(bSName,"'","\'"))%>')"><i class="fas fa-plus"></i> 入库</button>
                        <button class="btn btn-sm btn-outline" onclick="openSafety(<%=bSId%>,<%=bSSafety%>,'<%=Server.HTMLEncode(Replace(bSName,"'","\'"))%>')"><i class="fas fa-shield-alt"></i></button>
                        <button class="btn btn-sm btn-outline" onclick="openCost(<%=bSId%>,<%=bSCost%>,'<%=Server.HTMLEncode(Replace(bSName,"'","\'"))%>')"><i class="fas fa-tag"></i></button>
                    </td>
                </tr>
                <%      rsBottles.MoveNext
                    Loop
                    rsBottles.Close : Set rsBottles = Nothing
                End If
                On Error GoTo 0 %>
            </tbody>
        </table>
    </div>

    <!-- 入库 Modal -->
    <div id="restockModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-plus-circle"></i> 瓶子入库 — <span id="restockName"></span></div>
            <form method="post">
                <input type="hidden" name="action" value="restock">
                <input type="hidden" name="bottleId" id="restockBottleId">
                <div class="form-group">
                    <label>入库数量</label>
                    <input type="number" name="qty" min="1" value="1" required>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="closeModal('restockModal')">取消</button>
                    <button type="submit" class="btn btn-primary">确认入库</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 安全库存 Modal -->
    <div id="safetyModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-shield-alt"></i> 安全库存设置 — <span id="safetyName"></span></div>
            <form method="post">
                <input type="hidden" name="action" value="update_safety">
                <input type="hidden" name="bottleId" id="safetyBottleId">
                <div class="form-group">
                    <label>安全库存阈值</label>
                    <input type="number" name="safetyStock" id="safetyStock" min="0" required>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="closeModal('safetyModal')">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 单价 Modal -->
    <div id="costModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-tag"></i> 单价设置 — <span id="costName"></span></div>
            <form method="post">
                <input type="hidden" name="action" value="update_cost">
                <input type="hidden" name="bottleId" id="costBottleId">
                <div class="form-group">
                    <label>单价 (元)</label>
                    <input type="number" name="unitCost" id="unitCost" step="0.01" min="0" required>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="closeModal('costModal')">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    function openRestock(id, name) {
        document.getElementById('restockBottleId').value = id;
        document.getElementById('restockName').textContent = name;
        document.getElementById('restockModal').classList.add('active');
    }
    function openSafety(id, val, name) {
        document.getElementById('safetyBottleId').value = id;
        document.getElementById('safetyStock').value = val;
        document.getElementById('safetyName').textContent = name;
        document.getElementById('safetyModal').classList.add('active');
    }
    function openCost(id, val, name) {
        document.getElementById('costBottleId').value = id;
        document.getElementById('unitCost').value = val;
        document.getElementById('costName').textContent = name;
        document.getElementById('costModal').classList.add('active');
    }
    function closeModal(id) {
        document.getElementById(id).classList.remove('active');
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
