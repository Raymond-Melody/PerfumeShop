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

' 进行中的工单
Dim rsWorkshop
Set rsWorkshop = conn.Execute("SELECT po.*, o.OrderNo, o.ShippingName FROM ProductionOrders po LEFT JOIN Orders o ON po.OrderID=o.OrderID WHERE po.Status='InProgress' ORDER BY po.Priority DESC, po.StartedAt ASC")

' 统计
Dim wsInProgress, wsCompletedToday
wsInProgress = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='InProgress'"))
wsCompletedToday = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed' AND CompletedAt >= CAST(GETDATE() AS DATE)"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>车间作业 - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #FF9800; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #FF9800; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 25px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 32px; font-weight: bold; }
        .stat-card .label { font-size: 12px; color: #888; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(255,152,0,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(255,152,0,0.15); color: #ffb74d; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        

        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-cogs"></i> 车间作业</h2>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=wsInProgress%></span><span class="label">进行中工单</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=wsCompletedToday%></span><span class="label">今日完成</span></div>
        </div>
        
        <div class="card">
            <div class="card-header">生产中工单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>工单号</th><th>订单号</th><th>客户</th><th>配方</th><th>计划量</th><th>负责人</th><th>开始时间</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim wsRow : wsRow = 0
                    If Not rsWorkshop Is Nothing Then
                        Do While Not rsWorkshop.EOF
                            wsRow = wsRow + 1
                    %>
                        <tr>
                            <td><strong><%=rsWorkshop("WorkOrderNo") & ""%></strong></td>
                            <td><%=rsWorkshop("OrderNo") & ""%></td>
                            <td><%=rsWorkshop("ShippingName") & ""%></td>
                            <td><%=rsWorkshop("RecipeName") & ""%></td>
                            <td><%=rsWorkshop("PlannedQty") & ""%></td>
                            <td><%=rsWorkshop("AssignedTo") & ""%></td>
                            <td class="text-muted"><%=IIF(IsNull(rsWorkshop("StartedAt")),"",Left(rsWorkshop("StartedAt"),10))%></td>
                            <td>
                                <form method="post" action="production_management.asp" style="display:inline;">
                                    <input type="hidden" name="action" value="update_status">
                                    <input type="hidden" name="production_id" value="<%=rsWorkshop("ProductionID")%>">
                                    <input type="hidden" name="new_status" value="Completed">
                                    <button type="submit" class="btn btn-success btn-sm">完成</button>
                                </form>
                            </td>
                        </tr>
                    <%
                            rsWorkshop.MoveNext
                        Loop
                        rsWorkshop.Close
                    End If
                    Set rsWorkshop = Nothing
                    If wsRow = 0 Then %>
                        <tr><td colspan="8" class="text-center text-muted" style="padding:40px;">当前无进行中工单</td></tr>
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
