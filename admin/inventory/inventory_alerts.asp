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
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

' ========== 预警设置 ==========
Dim enableAlert
enableAlert = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableLowStockAlert'")
If IsNull(enableAlert) Or enableAlert = "" Then enableAlert = "1"

' ========== 成品预警 ==========
Dim piAlerts, piCritical, piWarning
piAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
piCritical = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductInventory WHERE StockQty <= 0 AND SafetyStock > 0"))
piWarning = piAlerts - piCritical

Dim rsPIAlerts
Set rsPIAlerts = conn.Execute("SELECT pi.*, p.ProductName FROM ProductInventory pi LEFT JOIN Products p ON pi.ProductID=p.ProductID WHERE pi.StockQty <= pi.SafetyStock AND pi.SafetyStock > 0 ORDER BY CASE WHEN pi.StockQty<=0 THEN 0 ELSE 1 END, pi.StockQty ASC")

' ========== 瓶子预警 ==========
Dim btAlerts, btCritical, btWarning
btAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
btCritical = SafeNum(GetScalar("SELECT COUNT(*) FROM BottleInventory WHERE StockQty <= 0 AND SafetyStock > 0"))
btWarning = btAlerts - btCritical

Dim rsBTAlerts
Set rsBTAlerts = conn.Execute("SELECT * FROM BottleInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0 ORDER BY CASE WHEN StockQty<=0 THEN 0 ELSE 1 END, StockQty ASC")

' ========== 包装物预警 ==========
Dim pkAlerts, pkCritical, pkWarning
pkAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
pkCritical = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory WHERE StockQty <= 0 AND SafetyStock > 0"))
pkWarning = pkAlerts - pkCritical

Dim rsPKAlerts
Set rsPKAlerts = conn.Execute("SELECT * FROM PackagingInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0 ORDER BY CASE WHEN StockQty<=0 THEN 0 ELSE 1 END, StockQty ASC")

' ========== 原料预警 ==========
Dim rmAlerts, rmCritical, rmWarning
rmAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
rmCritical = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= 0 AND SafetyStock > 0"))
rmWarning = rmAlerts - rmCritical

Dim rsRMAlerts
Set rsRMAlerts = conn.Execute("SELECT * FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0 ORDER BY CASE WHEN StockQty<=0 THEN 0 ELSE 1 END, StockQty ASC")

' ========== 香调预警 ==========
Dim ntAlerts, ntCritical, ntWarning
ntAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel AND MinStockLevel > 0"))
ntCritical = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= 0 AND MinStockLevel > 0"))
ntWarning = ntAlerts - ntCritical

Dim rsNTAlerts
Set rsNTAlerts = conn.Execute("SELECT ni.*, fn.NoteName, fn.NoteType FROM NoteInventory ni INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.StockQuantity <= ni.MinStockLevel AND ni.MinStockLevel > 0 ORDER BY CASE WHEN ni.StockQuantity<=0 THEN 0 ELSE 1 END, ni.StockQuantity ASC")

' ========== 总计 ==========
Dim totalAlerts, totalCritical, totalWarning
totalAlerts = piAlerts + btAlerts + pkAlerts + rmAlerts + ntAlerts
totalCritical = piCritical + btCritical + pkCritical + rmCritical + ntCritical
totalWarning = piWarning + btWarning + pkWarning + rmWarning + ntWarning
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>库存预警 - 库存管理中心</title>
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
        .page-title i { color: #f44336; }
        
        .alert-status { display: inline-flex; align-items: center; gap: 8px; padding: 8px 16px; border-radius: 8px; font-size: 14px; margin-bottom: 20px; }
        .alert-status.on { background: rgba(76,175,80,0.15); color: #81c784; }
        .alert-status.off { background: rgba(244,67,54,0.15); color: #e57373; }
        .alert-dot { width: 8px; height: 8px; border-radius: 50%; }
        .alert-dot.on { background: #4CAF50; }
        .alert-dot.off { background: #f44336; }
        
        .stats-summary { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .summary-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 22px; border-radius: 14px; border: 1px solid rgba(255,255,255,0.06); text-align: center; }
        .summary-card.all { border-top: 3px solid #2196F3; }
        .summary-card.critical-card { border-top: 3px solid #f44336; }
        .summary-card.warning-card { border-top: 3px solid #FF9800; }
        .summary-card.ok-card { border-top: 3px solid #4CAF50; }
        .summary-card .num { font-size: 32px; font-weight: 700; display: block; }
        .summary-card .label { font-size: 12px; color: #888; margin-top: 6px; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 18px; border-radius: 10px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 26px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 11px; color: #888; display: block; margin-top: 4px; }
        
        .section-title { font-size: 16px; color: #e0e0e0; margin: 30px 0 15px; display: flex; align-items: center; gap: 8px; padding-bottom: 10px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); overflow: hidden; }
        .card-header { padding: 14px 20px; font-weight: 600; font-size: 15px; color: #e0e0e0; border-bottom: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; gap: 8px; }
        .card-header.red { background: rgba(244,67,54,0.08); }
        .card-header.orange { background: rgba(255,152,0,0.08); }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { padding: 10px 12px; text-align: left; font-weight: 600; font-size: 12px; color: #888; border-bottom: 1px solid rgba(255,255,255,0.04); }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.03); color: #e0e0e0; font-size: 13px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        tr.critical-row td { background: rgba(244,67,54,0.06); }
        
        .urgency-badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .urgency-high { background: rgba(244,67,54,0.2); color: #e57373; }
        .urgency-medium { background: rgba(255,152,0,0.2); color: #ffb74d; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        
        .category-tabs { display: flex; gap: 5px; margin-bottom: 20px; flex-wrap: wrap; }
        .category-tab { padding: 8px 16px; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); border-radius: 8px; color: #b0b0b0; cursor: pointer; font-size: 13px; text-decoration: none; transition: all 0.2s; }
        .category-tab:hover, .category-tab.active { background: rgba(244,67,54,0.12); border-color: #f44336; color: #e57373; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-exclamation-triangle"></i> 库存预警</h2>
            <p style="font-size:13px;color:#888;margin-top:5px;">全品类低库存预警统一监控</p>
        </div>
        
        <div class="alert-status <%=IIF(enableAlert="1","on","off")%>">
            <span class="alert-dot <%=IIF(enableAlert="1","on","off")%>"></span>
            预警系统: <%=IIF(enableAlert="1","已启用","已关闭")%>
        </div>
        
        <!-- 总体汇总 -->
        <div class="stats-summary">
            <div class="summary-card all">
                <span class="num" style="color:#2196F3;"><%=totalAlerts%></span>
                <span class="label">预警总数</span>
            </div>
            <div class="summary-card critical-card">
                <span class="num" style="color:#f44336;"><%=totalCritical%></span>
                <span class="label">严重（零库存）</span>
            </div>
            <div class="summary-card warning-card">
                <span class="num" style="color:#FF9800;"><%=totalWarning%></span>
                <span class="label">警告（低库存）</span>
            </div>
            <div class="summary-card ok-card">
                <span class="num" style="color:#4CAF50;"><%=IIF(totalAlerts=0,"✓","")%></span>
                <span class="label"><%=IIF(totalAlerts=0,"一切正常","需关注")%></span>
            </div>
        </div>
        
        <%
        ' 辅助渲染函数 - 输出预警表格行
        Sub RenderAlertRow(itemName, itemCode, stockQty, safetyStock, unitName, extraInfo)
            Dim urgency, stockClass, qty
            qty = SafeNum(stockQty)
            If qty <= 0 Then urgency = "high" Else urgency = "medium"
            stockClass = "IF qty<=0:f44336 ELSE FF9800"
        %>
                        <tr class="<%=IIF(urgency="high","critical-row","")%>">
                            <td style="color:#888;"><%=itemCode%></td>
                            <td><strong><%=Server.HTMLEncode(itemName)%></strong></td>
                            <td style="color:<%=IIF(qty<=0,"#f44336","#FF9800")%>;"><%=FormatNumber(qty,1)%></td>
                            <td><%=FormatNumber(SafeNum(safetyStock),1)%></td>
                            <td><%=unitName%></td>
                            <% If extraInfo <> "" Then %><td><%=extraInfo%></td><% End If %>
                            <td><span class="urgency-badge <%=IIF(urgency="high","urgency-high","urgency-medium")%>"><%=IIF(urgency="high","紧急","警告")%></span></td>
                        </tr>
        <%
        End Sub
        %>
        
        <!-- ===== 成品预警 ===== -->
        <div class="section-title"><i class="fas fa-box" style="color:#4CAF50;"></i> 成品库存预警</div>
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=piAlerts%></span><span class="label">预警总数</span></div>
            <div class="stat-card" style="border-top:2px solid #f44336;"><span class="num" style="color:#f44336;"><%=piCritical%></span><span class="label">严重（零库存）</span></div>
            <div class="stat-card" style="border-top:2px solid #FF9800;"><span class="num" style="color:#FF9800;"><%=piWarning%></span><span class="label">警告（低库存）</span></div>
        </div>
        <div class="card">
            <div class="card-header red"><i class="fas fa-list"></i> 成品预警清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>产品名称</th><th>库存类型</th><th>当前库存</th><th>安全库存</th><th>成本</th><th>紧急程度</th></tr></thead>
                    <tbody>
                    <%
                    Dim piRowCount : piRowCount = 0
                    If Not rsPIAlerts Is Nothing Then
                        Do While Not rsPIAlerts.EOF
                            piRowCount = piRowCount + 1
                            Dim piQty : piQty = SafeNum(rsPIAlerts("StockQty"))
                            Dim piSafety : piSafety = SafeNum(rsPIAlerts("SafetyStock"))
                            Dim piUrgency : If piQty <= 0 Then piUrgency = "high" Else piUrgency = "medium"
                    %>
                        <tr class="<%=IIF(piUrgency="high","critical-row","")%>">
                            <td><strong><%=Server.HTMLEncode(rsPIAlerts("ProductName") & "")%></strong></td>
                            <td><%=rsPIAlerts("StockType") & ""%></td>
                            <td style="color:<%=IIF(piQty<=0,"#f44336","#FF9800")%>;"><%=piQty%></td>
                            <td><%=piSafety%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsPIAlerts("UnitCost")),2)%></td>
                            <td><span class="urgency-badge <%=IIF(piUrgency="high","urgency-high","urgency-medium")%>"><%=IIF(piUrgency="high","紧急","警告")%></span></td>
                        </tr>
                    <%
                            rsPIAlerts.MoveNext
                        Loop
                        rsPIAlerts.Close
                    End If
                    Set rsPIAlerts = Nothing
                    If piRowCount = 0 Then
                    %>
                        <tr><td colspan="6" class="text-center text-muted" style="padding:20px;color:#81c784;">✓ 所有成品库存正常</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- ===== 瓶子预警 ===== -->
        <div class="section-title"><i class="fas fa-flask" style="color:#2196F3;"></i> 瓶子库存预警</div>
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=btAlerts%></span><span class="label">预警总数</span></div>
            <div class="stat-card" style="border-top:2px solid #f44336;"><span class="num" style="color:#f44336;"><%=btCritical%></span><span class="label">严重</span></div>
            <div class="stat-card" style="border-top:2px solid #FF9800;"><span class="num" style="color:#FF9800;"><%=btWarning%></span><span class="label">警告</span></div>
        </div>
        <div class="card">
            <div class="card-header orange"><i class="fas fa-list"></i> 瓶子预警清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>瓶子名称</th><th>当前库存</th><th>安全库存</th><th>成本</th><th>紧急程度</th></tr></thead>
                    <tbody>
                    <%
                    Dim btRowCount : btRowCount = 0
                    If Not rsBTAlerts Is Nothing Then
                        Do While Not rsBTAlerts.EOF
                            btRowCount = btRowCount + 1
                            Dim btQty : btQty = SafeNum(rsBTAlerts("StockQty"))
                            Dim btSafety : btSafety = SafeNum(rsBTAlerts("SafetyStock"))
                            Dim btUrgency : If btQty <= 0 Then btUrgency = "high" Else btUrgency = "medium"
                    %>
                        <tr class="<%=IIF(btUrgency="high","critical-row","")%>">
                            <td><strong><%=Server.HTMLEncode(rsBTAlerts("BottleName") & "")%></strong></td>
                            <td style="color:<%=IIF(btQty<=0,"#f44336","#FF9800")%>;"><%=btQty%></td>
                            <td><%=btSafety%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsBTAlerts("UnitCost")),2)%></td>
                            <td><span class="urgency-badge <%=IIF(btUrgency="high","urgency-high","urgency-medium")%>"><%=IIF(btUrgency="high","紧急","警告")%></span></td>
                        </tr>
                    <%
                            rsBTAlerts.MoveNext
                        Loop
                        rsBTAlerts.Close
                    End If
                    Set rsBTAlerts = Nothing
                    If btRowCount = 0 Then
                    %>
                        <tr><td colspan="5" class="text-center text-muted" style="padding:20px;color:#81c784;">✓ 所有瓶子库存正常</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- ===== 包装物预警 ===== -->
        <div class="section-title"><i class="fas fa-box-open" style="color:#FF9800;"></i> 包装物库存预警</div>
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=pkAlerts%></span><span class="label">预警总数</span></div>
            <div class="stat-card" style="border-top:2px solid #f44336;"><span class="num" style="color:#f44336;"><%=pkCritical%></span><span class="label">严重</span></div>
            <div class="stat-card" style="border-top:2px solid #FF9800;"><span class="num" style="color:#FF9800;"><%=pkWarning%></span><span class="label">警告</span></div>
        </div>
        <div class="card">
            <div class="card-header red"><i class="fas fa-list"></i> 包装物预警清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>物品名称</th><th>当前库存</th><th>安全库存</th><th>成本</th><th>紧急程度</th></tr></thead>
                    <tbody>
                    <%
                    Dim pkRowCount : pkRowCount = 0
                    If Not rsPKAlerts Is Nothing Then
                        Do While Not rsPKAlerts.EOF
                            pkRowCount = pkRowCount + 1
                            Dim pkQty : pkQty = SafeNum(rsPKAlerts("StockQty"))
                            Dim pkSafety : pkSafety = SafeNum(rsPKAlerts("SafetyStock"))
                            Dim pkUrgency : If pkQty <= 0 Then pkUrgency = "high" Else pkUrgency = "medium"
                    %>
                        <tr class="<%=IIF(pkUrgency="high","critical-row","")%>">
                            <td><strong><%=Server.HTMLEncode(rsPKAlerts("ItemName") & "")%></strong></td>
                            <td style="color:<%=IIF(pkQty<=0,"#f44336","#FF9800")%>;"><%=pkQty%></td>
                            <td><%=pkSafety%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsPKAlerts("UnitCost")),2)%></td>
                            <td><span class="urgency-badge <%=IIF(pkUrgency="high","urgency-high","urgency-medium")%>"><%=IIF(pkUrgency="high","紧急","警告")%></span></td>
                        </tr>
                    <%
                            rsPKAlerts.MoveNext
                        Loop
                        rsPKAlerts.Close
                    End If
                    Set rsPKAlerts = Nothing
                    If pkRowCount = 0 Then
                    %>
                        <tr><td colspan="5" class="text-center text-muted" style="padding:20px;color:#81c784;">✓ 所有包装物库存正常</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- ===== 原料预警 ===== -->
        <div class="section-title"><i class="fas fa-boxes" style="color:#9C27B0;"></i> 原料库存预警</div>
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=rmAlerts%></span><span class="label">预警总数</span></div>
            <div class="stat-card" style="border-top:2px solid #f44336;"><span class="num" style="color:#f44336;"><%=rmCritical%></span><span class="label">严重</span></div>
            <div class="stat-card" style="border-top:2px solid #FF9800;"><span class="num" style="color:#FF9800;"><%=rmWarning%></span><span class="label">警告</span></div>
        </div>
        <div class="card">
            <div class="card-header red"><i class="fas fa-list"></i> 原料预警清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>编码</th><th>名称</th><th>当前库存</th><th>安全库存</th><th>单位</th><th>单价</th><th>紧急程度</th></tr></thead>
                    <tbody>
                    <%
                    Dim rmRowCount : rmRowCount = 0
                    If Not rsRMAlerts Is Nothing Then
                        Do While Not rsRMAlerts.EOF
                            rmRowCount = rmRowCount + 1
                            Dim rmQty : rmQty = SafeNum(rsRMAlerts("StockQty"))
                            Dim rmSafety : rmSafety = SafeNum(rsRMAlerts("SafetyStock"))
                            Dim rmUrgency : If rmQty <= 0 Then rmUrgency = "high" Else rmUrgency = "medium"
                    %>
                        <tr class="<%=IIF(rmUrgency="high","critical-row","")%>">
                            <td style="color:#888;"><%=rsRMAlerts("ItemCode") & ""%></td>
                            <td><strong><%=Server.HTMLEncode(rsRMAlerts("ItemName") & "")%></strong></td>
                            <td style="color:<%=IIF(rmQty<=0,"#f44336","#FF9800")%>;"><%=FormatNumber(rmQty,1)%></td>
                            <td><%=FormatNumber(rmSafety,1)%></td>
                            <td><%=rsRMAlerts("Unit") & ""%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsRMAlerts("UnitPrice")),2)%></td>
                            <td><span class="urgency-badge <%=IIF(rmUrgency="high","urgency-high","urgency-medium")%>"><%=IIF(rmUrgency="high","紧急","警告")%></span></td>
                        </tr>
                    <%
                            rsRMAlerts.MoveNext
                        Loop
                        rsRMAlerts.Close
                    End If
                    Set rsRMAlerts = Nothing
                    If rmRowCount = 0 Then
                    %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:20px;color:#81c784;">✓ 所有原料库存正常</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- ===== 香调预警 ===== -->
        <div class="section-title"><i class="fas fa-layer-group" style="color:#00BCD4;"></i> 香调库存预警</div>
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=ntAlerts%></span><span class="label">预警总数</span></div>
            <div class="stat-card" style="border-top:2px solid #f44336;"><span class="num" style="color:#f44336;"><%=ntCritical%></span><span class="label">严重</span></div>
            <div class="stat-card" style="border-top:2px solid #FF9800;"><span class="num" style="color:#FF9800;"><%=ntWarning%></span><span class="label">警告</span></div>
        </div>
        <div class="card">
            <div class="card-header orange"><i class="fas fa-list"></i> 香调预警清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>香调名称</th><th>类型</th><th>当前库存</th><th>最低库存</th><th>紧急程度</th></tr></thead>
                    <tbody>
                    <%
                    Dim ntRowCount : ntRowCount = 0
                    If Not rsNTAlerts Is Nothing Then
                        Do While Not rsNTAlerts.EOF
                            ntRowCount = ntRowCount + 1
                            Dim ntQty : ntQty = SafeNum(rsNTAlerts("StockQuantity"))
                            Dim ntMin : ntMin = SafeNum(rsNTAlerts("MinStockLevel"))
                            Dim ntUrgency : If ntQty <= 0 Then ntUrgency = "high" Else ntUrgency = "medium"
                    %>
                        <tr class="<%=IIF(ntUrgency="high","critical-row","")%>">
                            <td><strong><%=Server.HTMLEncode(rsNTAlerts("NoteName") & "")%></strong></td>
                            <td><%=rsNTAlerts("NoteType") & ""%></td>
                            <td style="color:<%=IIF(ntQty<=0,"#f44336","#FF9800")%>;"><%=ntQty%></td>
                            <td><%=ntMin%></td>
                            <td><span class="urgency-badge <%=IIF(ntUrgency="high","urgency-high","urgency-medium")%>"><%=IIF(ntUrgency="high","紧急","警告")%></span></td>
                        </tr>
                    <%
                            rsNTAlerts.MoveNext
                        Loop
                        rsNTAlerts.Close
                    End If
                    Set rsNTAlerts = Nothing
                    If ntRowCount = 0 Then
                    %>
                        <tr><td colspan="5" class="text-center text-muted" style="padding:20px;color:#81c784;">✓ 所有香调库存正常</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
