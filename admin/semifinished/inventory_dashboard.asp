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

' ========== 原料统计 ==========
Dim rawTotal, rawLowStock, rawZeroStock, rawTotalValue
rawTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory"))
rawLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
rawZeroStock = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= 0"))
rawTotalValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * UnitPrice),0) FROM RawMaterialInventory"))

' ========== 基香统计 ==========
Dim bnTotal, bnActive, bnInactive
bnTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes"))
bnActive = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes WHERE IsActive=1"))
bnInactive = bnTotal - bnActive

' ========== 香调统计 ==========
Dim noteTotal, noteLowStock, noteTotalQty, noteValue
noteTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory"))
noteLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel AND MinStockLevel > 0"))
noteTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQuantity),0) FROM NoteInventory"))
noteValue = SafeNum(GetScalar("SELECT ISNULL(SUM(ni.StockQuantity * fn.PriceAddition),0) FROM NoteInventory ni INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID"))

' ========== Accord生产统计 ==========
Dim accordPending, accordProgress, accordCompleted, accordThisMonth
accordPending = SafeNum(GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status='Pending'"))
accordProgress = SafeNum(GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status='InProgress'"))
accordCompleted = SafeNum(GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status IN ('Completed','QC')"))
accordThisMonth = SafeNum(GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status IN ('Completed','QC') AND CompletedAt >= DATEADD(month,-1,GETDATE())"))

' ========== 低库存预警列表 ==========
Dim rsLowRaw, rsLowNote
Set rsLowRaw = conn.Execute("SELECT TOP 5 MaterialID, ItemName, ItemCode, StockQty, SafetyStock, Unit FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0 ORDER BY (SafetyStock - StockQty) DESC")
Set rsLowNote = conn.Execute("SELECT TOP 5 ni.NoteID, fn.NoteName, fn.NoteType, ni.StockQuantity, ni.MinStockLevel FROM NoteInventory ni INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.StockQuantity <= ni.MinStockLevel AND ni.MinStockLevel > 0 ORDER BY ni.StockQuantity ASC")

' ========== 出库统计 ==========
Dim moToday, moThisMonth
moToday = SafeNum(GetScalar("SELECT COUNT(*) FROM MaterialOutbound WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)"))
moThisMonth = SafeNum(GetScalar("SELECT ISNULL(COUNT(*),0) FROM MaterialOutbound WHERE CreatedAt >= DATEADD(month,-1,GETDATE())"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>库存仪表盘 - 半成品生产中心</title>
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
        .page-title i { color: #2196F3; }
        
        .section-title { font-size: 16px; color: #e0e0e0; margin-bottom: 15px; display: flex; align-items: center; gap: 8px; padding-bottom: 10px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .section-title i { font-size: 14px; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; gap: 15px; }
        .stat-icon { width: 50px; height: 50px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 22px; flex-shrink: 0; }
        .stat-icon.raw { background: rgba(76,175,80,0.15); color: #4CAF50; }
        .stat-icon.base { background: rgba(33,150,243,0.15); color: #2196F3; }
        .stat-icon.note { background: rgba(255,152,0,0.15); color: #FF9800; }
        .stat-icon.accord { background: rgba(156,39,176,0.15); color: #9C27B0; }
        .stat-info .num { font-size: 24px; font-weight: bold; }
        .stat-info .label { font-size: 12px; color: #888; }
        
        .info-cards { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 25px; }
        .info-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); overflow: hidden; }
        .info-card-header { padding: 14px 20px; font-weight: 600; font-size: 15px; color: #e0e0e0; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .info-card-header.raw { background: rgba(76,175,80,0.08); }
        .info-card-header.note { background: rgba(255,152,0,0.08); }
        .info-card-body { padding: 16px 20px; }
        
        table { width: 100%; border-collapse: collapse; }
        th { padding: 10px 12px; text-align: left; font-weight: 600; font-size: 12px; color: #888; border-bottom: 1px solid rgba(255,255,255,0.04); }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.03); color: #e0e0e0; font-size: 13px; }
        
        .alert-badge { background: rgba(244,67,54,0.15); color: #e57373; padding: 2px 8px; border-radius: 10px; font-size: 11px; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-chart-pie"></i> 库存仪表盘</h2>
        </div>
        
        <!-- 核心指标 -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon raw"><i class="fas fa-boxes"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#4CAF50;"><%=rawTotal%></div>
                    <div class="label">原料种类 | 总值 ¥<%=FormatNumber(rawTotalValue,0)%></div>
                    <% If rawLowStock > 0 Then %><div class="alert-badge" style="margin-top:4px;"><%=rawLowStock%> 低库存</div><% End If %>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon base"><i class="fas fa-database"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#2196F3;"><%=bnActive%></div>
                    <div class="label">激活基香 / 共 <%=bnTotal%> 种</div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon note"><i class="fas fa-layer-group"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#FF9800;"><%=noteTotalQty%></div>
                    <div class="label">香调库存总量 | <%=noteTotal%> 种</div>
                    <% If noteLowStock > 0 Then %><div class="alert-badge" style="margin-top:4px;"><%=noteLowStock%> 低库存</div><% End If %>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon accord"><i class="fas fa-cogs"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#9C27B0;"><%=accordProgress%></div>
                    <div class="label">Accord生产中 | 本月完成 <%=accordThisMonth%></div>
                </div>
            </div>
        </div>
        
        <!-- 详细面板 -->
        <div class="info-cards">
            <!-- 原料低库存预警 -->
            <div class="info-card">
                <div class="info-card-header raw"><i class="fas fa-exclamation-triangle"></i> 原料低库存预警</div>
                <div class="info-card-body">
                    <table>
                        <thead><tr><th>编码</th><th>名称</th><th>库存</th><th>安全库存</th><th>单位</th></tr></thead>
                        <tbody>
                        <%
                        Dim rawRowCount : rawRowCount = 0
                        If Not rsLowRaw Is Nothing Then
                            Do While Not rsLowRaw.EOF
                                rawRowCount = rawRowCount + 1
                        %>
                            <tr>
                                <td style="color:#888;"><%=rsLowRaw("ItemCode") & ""%></td>
                                <td><%=Server.HTMLEncode(rsLowRaw("ItemName") & "")%></td>
                                <td style="color:#f44336;"><%=FormatNumber(SafeNum(rsLowRaw("StockQty")),1)%></td>
                                <td><%=FormatNumber(SafeNum(rsLowRaw("SafetyStock")),1)%></td>
                                <td><%=rsLowRaw("Unit") & ""%></td>
                            </tr>
                        <%
                                rsLowRaw.MoveNext
                            Loop
                            rsLowRaw.Close
                        End If
                        Set rsLowRaw = Nothing
                        If rawRowCount = 0 Then
                        %>
                            <tr><td colspan="5" class="text-center text-muted">暂无低库存原料</td></tr>
                        <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- 香调低库存预警 -->
            <div class="info-card">
                <div class="info-card-header note"><i class="fas fa-bell"></i> 香调低库存预警</div>
                <div class="info-card-body">
                    <table>
                        <thead><tr><th>香调</th><th>类型</th><th>库存</th><th>最低库存</th></tr></thead>
                        <tbody>
                        <%
                        Dim noteRowCount : noteRowCount = 0
                        If Not rsLowNote Is Nothing Then
                            Do While Not rsLowNote.EOF
                                noteRowCount = noteRowCount + 1
                        %>
                            <tr>
                                <td><%=Server.HTMLEncode(rsLowNote("NoteName") & "")%></td>
                                <td><%=rsLowNote("NoteType") & ""%></td>
                                <td style="color:#f44336;"><%=rsLowNote("StockQuantity")%></td>
                                <td><%=rsLowNote("MinStockLevel")%></td>
                            </tr>
                        <%
                                rsLowNote.MoveNext
                            Loop
                            rsLowNote.Close
                        End If
                        Set rsLowNote = Nothing
                        If noteRowCount = 0 Then
                        %>
                            <tr><td colspan="4" class="text-center text-muted">暂无低库存香调</td></tr>
                        <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- 今日概览 -->
        <div class="section-title"><i class="fas fa-calendar-day" style="color:#2196F3;"></i> 今日运营概览</div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon" style="background:rgba(156,39,176,0.15);color:#9C27B0;"><i class="fas fa-truck-loading"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#9C27B0;"><%=moToday%></div>
                    <div class="label">今日出库单</div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:rgba(33,150,243,0.15);color:#2196F3;"><i class="fas fa-calendar-check"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#2196F3;"><%=moThisMonth%></div>
                    <div class="label">本月出库单</div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:rgba(76,175,80,0.15);color:#4CAF50;"><i class="fas fa-check-circle"></i></div>
                <div class="stat-info">
                    <div class="num" style="color:#4CAF50;"><%=accordCompleted%></div>
                    <div class="label">已完成Accord</div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
