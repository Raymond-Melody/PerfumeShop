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

' 确保字段存在
On Error Resume Next
conn.Execute "SELECT WarehouseInAt FROM ProductionOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE ProductionOrders ADD WarehouseInAt DATETIME"
On Error GoTo 0

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
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

Dim msg, msgType
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Then msgType = "error"

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim whAction : whAction = Trim(Request.Form("action"))
    If whAction = "warehouse_in" Then
        Dim whPoID : whPoID = SafeNum(Request.Form("production_id"))
        If whPoID > 0 Then
            conn.Execute "UPDATE ProductionOrders SET Status='WarehouseIn', WarehouseInAt=GETDATE(), UpdatedAt=GETDATE() WHERE ProductionID=" & whPoID
            conn.Execute "INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedBy, CreatedAt) VALUES (" & _
                whPoID & ",'WarehouseIn','成品入库','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
            Response.Redirect "prod_warehouse.asp?msg=入库成功"
            Response.End
        End If
    End If
End If

Dim whQC, whIn, whToday
whQC = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='QC_Passed'"))
whIn = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='WarehouseIn'"))
whToday = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='WarehouseIn' AND WarehouseInAt >= CAST(GETDATE() AS DATE)"))

Dim rsWH
Set rsWH = conn.Execute("SELECT po.*, o.OrderNo, o.ShippingName FROM ProductionOrders po LEFT JOIN Orders o ON po.OrderID=o.OrderID WHERE po.Status IN ('QC_Passed','WarehouseIn','ShippedOut') ORDER BY po.UpdatedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>成品入库 - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #00BCD4; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #00BCD4; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; }
        .stat-card .label { font-size: 12px; color: #888; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(0,188,212,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(0,188,212,0.15); color: #80deea; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-qc-pass { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .status-warehouse { background: rgba(76,175,80,0.15); color: #81c784; }
        .status-shipped { background: rgba(0,188,212,0.15); color: #80deea; }
        

        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-warehouse"></i> 成品入库</h2>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=whQC%></span><span class="label">待入库</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=whIn%></span><span class="label">已入库</span></div>
            <div class="stat-card"><span class="num" style="color:#00BCD4;"><%=whToday%></span><span class="label">今日入库</span></div>
        </div>
        
        <div class="card">
            <div class="card-header">入库管理列表</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>工单号</th><th>订单号</th><th>客户</th><th>配方</th><th>计划量</th><th>状态</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim whRow : whRow = 0
                    If Not rsWH Is Nothing Then
                        Do While Not rsWH.EOF
                            whRow = whRow + 1
                            Dim whStatus : whStatus = CStr(rsWH("Status") & "")
                    %>
                        <tr>
                            <td><strong><%=rsWH("WorkOrderNo") & ""%></strong></td>
                            <td><%=rsWH("OrderNo") & ""%></td>
                            <td><%=rsWH("ShippingName") & ""%></td>
                            <td><%=rsWH("RecipeName") & ""%></td>
                            <td><%=rsWH("PlannedQty") & ""%></td>
                            <td><span class="status-badge <%=IIF(whStatus="QC_Passed","status-qc-pass",IIF(whStatus="WarehouseIn","status-warehouse","status-shipped"))%>"><%=whStatus%></span></td>
                            <td>
                                <% If whStatus = "QC_Passed" Then %>
                                <form method="post" style="display:inline;">
                                    <input type="hidden" name="action" value="warehouse_in">
                                    <input type="hidden" name="production_id" value="<%=rsWH("ProductionID")%>">
                                    <button type="submit" class="btn btn-primary btn-sm">入库</button>
                                </form>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rsWH.MoveNext
                        Loop
                        rsWH.Close
                    End If
                    Set rsWH = Nothing
                    If whRow = 0 Then %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:40px;">暂无待入库工单</td></tr>
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
