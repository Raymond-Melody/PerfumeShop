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

' ========== 原料预警 ==========
Dim rawAlerts, rawCritical, rawWarning
rawAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
rawCritical = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= 0 AND SafetyStock > 0"))
rawWarning = rawAlerts - rawCritical

Dim rsRawAlerts
Set rsRawAlerts = conn.Execute("SELECT MaterialID, ItemName, ItemCode, StockQty, SafetyStock, Unit, UnitPrice FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0 ORDER BY StockQty ASC")

' ========== 香调预警 ==========
Dim noteAlerts, noteCritical, noteWarning
noteAlerts = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel AND MinStockLevel > 0"))
noteCritical = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= 0 AND MinStockLevel > 0"))
noteWarning = noteAlerts - noteCritical

Dim rsNoteAlerts
Set rsNoteAlerts = conn.Execute("SELECT ni.*, fn.NoteName, fn.NoteType FROM NoteInventory ni INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.StockQuantity <= ni.MinStockLevel AND ni.MinStockLevel > 0 ORDER BY ni.StockQuantity ASC")

' ========== 预警设置 ==========
Dim enableAlert
enableAlert = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableLowStockAlert'")
If IsNull(enableAlert) Or enableAlert = "" Then enableAlert = "1"
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>库存预警 - 半成品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
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
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 25px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card.critical { border-top: 3px solid #f44336; }
        .stat-card.warning { border-top: 3px solid #FF9800; }
        .stat-card.info { border-top: 3px solid #2196F3; }
        .stat-card .num { font-size: 36px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        
        .section-title { font-size: 16px; color: #e0e0e0; margin: 25px 0 15px; display: flex; align-items: center; gap: 8px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(244,67,54,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(244,67,54,0.15); color: #e57373; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        tr.critical-row td { background: rgba(244,67,54,0.08); }
        
        .urgency-badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .urgency-high { background: rgba(244,67,54,0.2); color: #e57373; }
        .urgency-medium { background: rgba(255,152,0,0.2); color: #ffb74d; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-exclamation-triangle"></i> 库存预警</h2>
        </div>
        
        <div class="alert-status <%=IIF(enableAlert="1","on","off")%>">
            <span class="alert-dot <%=IIF(enableAlert="1","on","off")%>"></span>
            预警系统: <%=IIF(enableAlert="1","已启用","已关闭")%>
        </div>
        
        <!-- 原料预警统计 -->
        <div class="section-title"><i class="fas fa-boxes" style="color:#4CAF50;"></i> 原料库存预警</div>
        <div class="stats-grid">
            <div class="stat-card info"><span class="num" style="color:#2196F3;"><%=rawAlerts%></span><span class="label">预警总数</span></div>
            <div class="stat-card critical"><span class="num" style="color:#f44336;"><%=rawCritical%></span><span class="label">严重（零库存）</span></div>
            <div class="stat-card warning"><span class="num" style="color:#FF9800;"><%=rawWarning%></span><span class="label">警告（低库存）</span></div>
        </div>
        
        <div class="card">
            <div class="card-header">原料预警清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>编码</th><th>名称</th><th>当前库存</th><th>安全库存</th><th>单位</th><th>单价</th><th>紧急程度</th></tr></thead>
                    <tbody>
                    <%
                    Dim raRowCount : raRowCount = 0
                    If Not rsRawAlerts Is Nothing Then
                        Do While Not rsRawAlerts.EOF
                            raRowCount = raRowCount + 1
                            Dim rStock : rStock = SafeNum(rsRawAlerts("StockQty"))
                            Dim rSafety : rSafety = SafeNum(rsRawAlerts("SafetyStock"))
                            Dim rUrgency
                            If rStock <= 0 Then rUrgency = "high" Else rUrgency = "medium"
                    %>
                        <tr class="<%=IIF(rUrgency="high","critical-row","")%>">
                            <td style="color:#888;"><%=rsRawAlerts("ItemCode") & ""%></td>
                            <td><strong><%=Server.HTMLEncode(rsRawAlerts("ItemName") & "")%></strong></td>
                            <td style="color:<%=IIF(rStock<=0,"#f44336","#FF9800")%>;"><%=FormatNumber(rStock,1)%></td>
                            <td><%=FormatNumber(rSafety,1)%></td>
                            <td><%=rsRawAlerts("Unit") & ""%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsRawAlerts("UnitPrice")),2)%></td>
                            <td><span class="urgency-badge <%=IIF(rUrgency="high","urgency-high","urgency-medium")%>"><%=IIF(rUrgency="high","紧急","警告")%></span></td>
                        </tr>
                    <%
                            rsRawAlerts.MoveNext
                        Loop
                        rsRawAlerts.Close
                    End If
                    Set rsRawAlerts = Nothing
                    If raRowCount = 0 Then
                    %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:30px;color:#81c784;">✓ 所有原料库存正常</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- 香调预警统计 -->
        <div class="section-title"><i class="fas fa-layer-group" style="color:#FF9800;"></i> 香调库存预警</div>
        <div class="stats-grid">
            <div class="stat-card info"><span class="num" style="color:#2196F3;"><%=noteAlerts%></span><span class="label">预警总数</span></div>
            <div class="stat-card critical"><span class="num" style="color:#f44336;"><%=noteCritical%></span><span class="label">严重（零库存）</span></div>
            <div class="stat-card warning"><span class="num" style="color:#FF9800;"><%=noteWarning%></span><span class="label">警告（低库存）</span></div>
        </div>
        
        <div class="card">
            <div class="card-header">香调预警清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>香调名称</th><th>类型</th><th>当前库存</th><th>最低库存</th><th>紧急程度</th></tr></thead>
                    <tbody>
                    <%
                    Dim naRowCount : naRowCount = 0
                    If Not rsNoteAlerts Is Nothing Then
                        Do While Not rsNoteAlerts.EOF
                            naRowCount = naRowCount + 1
                            Dim nStock : nStock = SafeNum(rsNoteAlerts("StockQuantity"))
                            Dim nMin : nMin = SafeNum(rsNoteAlerts("MinStockLevel"))
                            Dim nUrgency
                            If nStock <= 0 Then nUrgency = "high" Else nUrgency = "medium"
                    %>
                        <tr class="<%=IIF(nUrgency="high","critical-row","")%>">
                            <td><strong><%=Server.HTMLEncode(rsNoteAlerts("NoteName") & "")%></strong></td>
                            <td><%=rsNoteAlerts("NoteType") & ""%></td>
                            <td style="color:<%=IIF(nStock<=0,"#f44336","#FF9800")%>;"><%=nStock%></td>
                            <td><%=nMin%></td>
                            <td><span class="urgency-badge <%=IIF(nUrgency="high","urgency-high","urgency-medium")%>"><%=IIF(nUrgency="high","紧急","警告")%></span></td>
                        </tr>
                    <%
                            rsNoteAlerts.MoveNext
                        Loop
                        rsNoteAlerts.Close
                    End If
                    Set rsNoteAlerts = Nothing
                    If naRowCount = 0 Then
                    %>
                        <tr><td colspan="5" class="text-center text-muted" style="padding:30px;color:#81c784;">✓ 所有香调库存正常</td></tr>
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
