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

Dim piTotal, piTotalQty, piLowStock, piValue
piTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductInventory"))
piTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM ProductInventory"))
piLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
piValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * UnitCost),0) FROM ProductInventory"))

Dim rsPI
Set rsPI = conn.Execute("SELECT pi.*, p.ProductName FROM ProductInventory pi LEFT JOIN Products p ON pi.ProductID=p.ProductID ORDER BY p.ProductName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>成品库存 - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #FF9800; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #FF9800; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; }
        .stat-card .label { font-size: 12px; color: #888; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(255,152,0,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(255,152,0,0.15); color: #ffb74d; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        tr.low-stock td { background: rgba(244,67,54,0.06); }
        
        .stock-bar { height: 6px; background: rgba(255,255,255,0.1); border-radius: 3px; margin-top: 4px; }
        .stock-bar-fill { height: 100%; border-radius: 3px; }
        .stock-bar-fill.safe { background: #4CAF50; }
        .stock-bar-fill.warning { background: #FF9800; }
        .stock-bar-fill.danger { background: #f44336; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-box"></i> 成品库存</h2>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=piTotal%></span><span class="label">库存品类</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=piTotalQty%></span><span class="label">总库存量</span></div>
            <div class="stat-card"><span class="num" style="color:#f44336;"><%=piLowStock%></span><span class="label">低库存预警</span></div>
            <div class="stat-card"><span class="num" style="color:#2196F3;">¥<%=FormatNumber(piValue,0)%></span><span class="label">库存总值</span></div>
        </div>
        
        <div class="card">
            <div class="card-header">成品库存清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>产品名称</th><th>库存类型</th><th>库存量</th><th>安全库存</th><th>单位成本</th><th>更新时间</th></tr></thead>
                    <tbody>
                    <%
                    Dim piRow : piRow = 0
                    If Not rsPI Is Nothing Then
                        Do While Not rsPI.EOF
                            piRow = piRow + 1
                            Dim piQty : piQty = SafeNum(rsPI("StockQty"))
                            Dim piSafety : piSafety = SafeNum(rsPI("SafetyStock"))
                            Dim piSC : piSC = "safe" : If piSafety > 0 And piQty <= piSafety Then piSC = "warning" : If piQty <= 0 Then piSC = "danger"
                            Dim piPct : piPct = 100 : If piSafety > 0 Then piPct = (piQty / piSafety) * 100 : If piPct > 100 Then piPct = 100
                    %>
                        <tr class="<%=IIF(piSC<>"safe","low-stock","")%>">
                            <td><strong><%=rsPI("ProductName") & ""%></strong></td>
                            <td><%=rsPI("StockType") & ""%></td>
                            <td><%=piQty%><div class="stock-bar"><div class="stock-bar-fill <%=piSC%>" style="width:<%=piPct%>%;"></div></div></td>
                            <td><%=piSafety%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsPI("UnitCost")),2)%></td>
                            <td class="text-muted"><%=IIF(IsNull(rsPI("UpdatedAt")),"",Left(rsPI("UpdatedAt"),10))%></td>
                        </tr>
                    <%
                            rsPI.MoveNext
                        Loop
                        rsPI.Close
                    End If
                    Set rsPI = Nothing
                    If piRow = 0 Then %>
                        <tr><td colspan="6" class="text-center text-muted" style="padding:40px;">暂无成品库存</td></tr>
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
