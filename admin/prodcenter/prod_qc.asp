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

Dim action, msg, msgType
action = Trim(Request.Form("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Then msgType = "error"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "qc_pass" Then
        Dim qcID, qcResult, qcNotes
        qcID = SafeNum(Request.Form("production_id"))
        qcResult = Trim(Request.Form("qc_result"))
        qcNotes = Trim(Request.Form("qc_notes"))
        If qcID > 0 Then
            conn.Execute "UPDATE ProductionOrders SET QCNotes='" & SafeSQL(qcNotes) & "', QCPassedAt=GETDATE(), UpdatedAt=GETDATE() WHERE ProductionID=" & qcID
            If qcResult = "Pass" Then
                conn.Execute "UPDATE ProductionOrders SET Status='QC_Passed' WHERE ProductionID=" & qcID
                conn.Execute "INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedBy, CreatedAt) VALUES (" & qcID & ",'QC_Passed','质检通过 - " & SafeSQL(qcNotes) & "','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
                Response.Redirect "prod_qc.asp?msg=质检通过"
            Else
                conn.Execute "UPDATE ProductionOrders SET Status='QC_Fail' WHERE ProductionID=" & qcID
                conn.Execute "INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedBy, CreatedAt) VALUES (" & qcID & ",'QC_Fail','质检不通过 - " & SafeSQL(qcNotes) & "','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
                Response.Redirect "prod_qc.asp?msg=质检未通过，需返工"
            End If
            Response.End
        End If
    End If
End If

Dim qcPending, qcPassed, qcFailed
qcPending = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='Completed'"))
qcPassed = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='QC_Passed' AND QCPassedAt >= CAST(GETDATE() AS DATE)"))
qcFailed = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductionOrders WHERE Status='QC_Fail'"))

Dim rsQC
Set rsQC = conn.Execute("SELECT po.*, o.OrderNo FROM ProductionOrders po LEFT JOIN Orders o ON po.OrderID=o.OrderID WHERE po.Status IN ('Completed','QC_Passed','QC_Fail') ORDER BY po.UpdatedAt DESC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>质量检验 - 产品生产中心</title>
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
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #9C27B0; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; }
        .stat-card .label { font-size: 12px; color: #888; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(156,39,176,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(156,39,176,0.15); color: #ce93d8; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-completed { background: rgba(76,175,80,0.15); color: #81c784; }
        .status-qc-pass { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .status-qc-fail { background: rgba(244,67,54,0.15); color: #e57373; }
        

        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #e57373; border: 1px solid rgba(244,67,54,0.3); }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); width: 90%; max-width: 500px; margin: 80px auto; padding: 30px; border-radius: 15px; }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .modal-header h3 { margin: 0; font-size: 18px; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 25px; }
        
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; margin-bottom: 6px; font-weight: 600; color: #e0e0e0; font-size: 13px; }
        .form-group select, .form-group textarea { width: 100%; padding: 10px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-check-circle"></i> 质量检验</h2>
        </div>
        
        <% If msg <> "" Then %><div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div><% End If %>
        
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=qcPending%></span><span class="label">待检验</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=qcPassed%></span><span class="label">今日通过</span></div>
            <div class="stat-card"><span class="num" style="color:#f44336;"><%=qcFailed%></span><span class="label">未通过</span></div>
        </div>
        
        <div class="card">
            <div class="card-header">检验列表</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>工单号</th><th>订单号</th><th>配方</th><th>计划量</th><th>状态</th><th>完成时间</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim qcRow : qcRow = 0
                    If Not rsQC Is Nothing Then
                        Do While Not rsQC.EOF
                            qcRow = qcRow + 1
                            Dim qcStatus : qcStatus = CStr(rsQC("Status") & "")
                    %>
                        <tr>
                            <td><strong><%=rsQC("WorkOrderNo") & ""%></strong></td>
                            <td><%=rsQC("OrderNo") & ""%></td>
                            <td><%=rsQC("RecipeName") & ""%></td>
                            <td><%=rsQC("PlannedQty") & ""%></td>
                            <td><span class="status-badge <%=IIF(qcStatus="Completed","status-completed",IIF(qcStatus="QC_Passed","status-qc-pass","status-qc-fail"))%>"><%=qcStatus%></span></td>
                            <td class="text-muted"><%=IIF(IsNull(rsQC("CompletedAt")),"",Left(rsQC("CompletedAt"),10))%></td>
                            <td>
                                <% If qcStatus = "Completed" Then %>
                                <button class="btn btn-success btn-sm" onclick="openQCModal(<%=rsQC("ProductionID")%>,'<%=rsQC("WorkOrderNo") & ""%>')">检验</button>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rsQC.MoveNext
                        Loop
                        rsQC.Close
                    End If
                    Set rsQC = Nothing
                    If qcRow = 0 Then %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:40px;">暂无待检验工单</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <div id="qcModal" class="modal">
        <div class="modal-content">
            <div class="modal-header"><h3>质量检验 - <span id="qcWONo"></span></h3><button class="modal-close" onclick="closeModal('qcModal')">&times;</button></div>
            <form method="post">
                <input type="hidden" name="action" value="qc_pass">
                <input type="hidden" name="production_id" id="qcProdID">
                <div class="form-group"><label>检验结果</label><select name="qc_result" required><option value="Pass">合格</option><option value="Fail">不合格</option></select></div>
                <div class="form-group"><label>备注</label><textarea name="qc_notes" rows="3"></textarea></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('qcModal')">取消</button>
                    <button type="submit" class="btn btn-success">提交</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    function openQCModal(id, wono) { document.getElementById('qcProdID').value = id; document.getElementById('qcWONo').innerText = wono; document.getElementById('qcModal').style.display = 'block'; }
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }
    window.onclick = function(event) { if (event.target.classList.contains('modal')) event.target.style.display = 'none'; }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
