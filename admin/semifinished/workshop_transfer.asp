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
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

Dim action, msg, msgType
action = Trim(Request.Form("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

Dim currentRole
currentRole = Session("AdminRoleCode")

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    
    If action = "create_transfer" Then
        Dim tNoteID, tQty, tNotes
        tNoteID = SafeNum(Request.Form("note_id"))
        tQty = SafeNum(Request.Form("request_qty"))
        tNotes = Trim(Request.Form("notes"))
        
        If tNoteID > 0 And tQty > 0 Then
            Dim tNoteName, tNoteStock, tTransferNo
            tNoteName = CStr(GetScalar("SELECT NoteName FROM FragranceNotes WHERE NoteID=" & tNoteID) & "")
            tNoteStock = SafeNum(GetScalar("SELECT StockQuantity FROM NoteInventory WHERE NoteID=" & tNoteID))
            
            If tQty > tNoteStock Then
                msg = "库存不足！当前库存: " & tNoteStock & ", 申请量: " & tQty
                msgType = "error"
            Else
                tTransferNo = "TRF" & Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2)
                
                On Error Resume Next
                Err.Clear
                Call BeginTransaction()
                
                conn.Execute "INSERT INTO WorkshopTransfer (TransferNo, NoteID, RequestQty, FromWorkshop, ToWorkshop, Status, RequestedBy, RequestedAt, Notes, CreatedAt) VALUES ('" & _
                    tTransferNo & "', " & tNoteID & ", " & tQty & ", 'SEMI', 'PROD', 'Requested', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), '" & SafeSQL(tNotes) & "', GETDATE())"
                
                If Err.Number <> 0 Then
                    msg = "创建调拨单失败: " & Err.Description
                    msgType = "error"
                    Call RollbackTransaction()
                    Err.Clear
                Else
                    ' 调拨申请自动扣减半成品库存
                    conn.Execute "UPDATE NoteInventory SET StockQuantity=StockQuantity-" & tQty & ", UpdatedAt=GETDATE() WHERE NoteID=" & tNoteID
                    If Err.Number <> 0 Then
                        Call RollbackTransaction()
                        Err.Clear
                        msg = "扣减库存失败"
                        msgType = "error"
                    Else
                        conn.Execute "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES (" & _
                            tNoteID & "," & tQty & ",'车间调拨','OUT','调拨至成品车间[" & tTransferNo & "]','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
                        Call CommitTransaction()
                    End If
                End If
                On Error GoTo 0
                
                If msg = "" Then
                    Response.Redirect "workshop_transfer.asp?msg=调拨申请成功！单号：" & tTransferNo & " 香调：" & tNoteName
                    Response.End
                End If
            End If
        Else
            msg = "请选择香调和需求量"
            msgType = "error"
        End If
    
    ElseIf action = "approve_transfer" Then
        Dim aTransferID, aAction
        aTransferID = SafeNum(Request.Form("transfer_id"))
        aAction = Trim(Request.Form("approve_action"))
        
        If aTransferID > 0 And aAction <> "" Then
            If aAction = "approve" Then
                conn.Execute "UPDATE WorkshopTransfer SET Status='Approved', FulfilledAt=GETDATE() WHERE TransferID=" & aTransferID
                Response.Redirect "workshop_transfer.asp?msg=调拨已批准"
            ElseIf aAction = "reject" Then
                Dim rNoteID, rQty
                rNoteID = SafeNum(GetScalar("SELECT NoteID FROM WorkshopTransfer WHERE TransferID=" & aTransferID))
                rQty = SafeNum(GetScalar("SELECT RequestQty FROM WorkshopTransfer WHERE TransferID=" & aTransferID))
                
                ' 退回库存
                conn.Execute "UPDATE NoteInventory SET StockQuantity=StockQuantity+" & rQty & ", UpdatedAt=GETDATE() WHERE NoteID=" & rNoteID
                conn.Execute "UPDATE WorkshopTransfer SET Status='Rejected' WHERE TransferID=" & aTransferID
                conn.Execute "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES (" & _
                    rNoteID & "," & rQty & ",'调拨退回','IN','调拨拒绝退回','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
                Response.Redirect "workshop_transfer.asp?msg=调拨已拒绝，库存已退回"
            End If
            Response.End
        End If
    
    ElseIf action = "fulfill_transfer" Then
        Dim fTransferID
        fTransferID = SafeNum(Request.Form("transfer_id"))
        
        If fTransferID > 0 Then
            conn.Execute "UPDATE WorkshopTransfer SET Status='Fulfilled' WHERE TransferID=" & fTransferID
            Response.Redirect "workshop_transfer.asp?msg=调拨已确认收货"
            Response.End
        End If
    End If
End If

' ========== 统计 ==========
Dim wtPending, wtApproved, wtFulfilled, wtTotal
wtPending = SafeNum(GetScalar("SELECT COUNT(*) FROM WorkshopTransfer WHERE Status='Requested'"))
wtApproved = SafeNum(GetScalar("SELECT COUNT(*) FROM WorkshopTransfer WHERE Status='Approved'"))
wtFulfilled = SafeNum(GetScalar("SELECT COUNT(*) FROM WorkshopTransfer WHERE Status='Fulfilled'"))
wtTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM WorkshopTransfer"))

' ========== 调拨记录 ==========
Dim rsTransfer
Set rsTransfer = conn.Execute("SELECT wt.*, fn.NoteName, fn.NoteType FROM WorkshopTransfer wt INNER JOIN FragranceNotes fn ON wt.NoteID=fn.NoteID ORDER BY wt.CreatedAt DESC")

' ========== 可用香调列表（有库存的） ==========
Dim rsAvailNotes
Set rsAvailNotes = conn.Execute("SELECT ni.NoteID, fn.NoteName, fn.NoteType, ni.StockQuantity FROM NoteInventory ni INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID WHERE ni.StockQuantity > 0 ORDER BY fn.NoteType, fn.NoteName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>车间调拨 - 半成品生产中心</title>
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
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #00BCD4; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(0,188,212,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(0,188,212,0.15); color: #80deea; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-requested { background: rgba(255,152,0,0.15); color: #ffb74d; }
        .status-approved { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .status-fulfilled { background: rgba(76,175,80,0.15); color: #81c784; }
        .status-rejected { background: rgba(244,67,54,0.15); color: #e57373; }
        

        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #e57373; border: 1px solid rgba(244,67,54,0.3); }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); width: 90%; max-width: 500px; margin: 80px auto; padding: 30px; border-radius: 15px; border: 1px solid rgba(255,255,255,0.06); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .modal-header h3 { margin: 0; font-size: 18px; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 25px; }
        
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; margin-bottom: 6px; font-weight: 600; color: #e0e0e0; font-size: 13px; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 10px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        .form-group input:focus, .form-group select:focus { outline: none; border-color: #00BCD4; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .text-right { text-align: right; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-exchange-alt"></i> 车间调拨管理</h2>
            <button class="btn btn-primary" onclick="openTransferModal()"><i class="fas fa-plus"></i> 新建调拨</button>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#00BCD4;"><%=wtTotal%></span><span class="label">总调拨单</span></div>
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=wtPending%></span><span class="label">待审批</span></div>
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=wtApproved%></span><span class="label">已批准</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=wtFulfilled%></span><span class="label">已收货</span></div>
        </div>
        
        <!-- 调拨记录 -->
        <div class="card">
            <div class="card-header">调拨记录（半成品车间 → 成品车间）</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>调拨单号</th><th>香调</th><th>类型</th><th>申请量</th><th>来源</th><th>目标</th><th>状态</th><th>申请人</th><th>时间</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim wtRowCount : wtRowCount = 0
                    If Not rsTransfer Is Nothing Then
                        Do While Not rsTransfer.EOF
                            wtRowCount = wtRowCount + 1
                            Dim wtStatus : wtStatus = CStr(rsTransfer("Status") & "")
                            Dim wtStatusClass
                            If wtStatus = "Requested" Then
                                wtStatusClass = "status-requested"
                            ElseIf wtStatus = "Approved" Then
                                wtStatusClass = "status-approved"
                            ElseIf wtStatus = "Fulfilled" Then
                                wtStatusClass = "status-fulfilled"
                            Else
                                wtStatusClass = "status-rejected"
                            End If
                    %>
                        <tr>
                            <td><strong><%=rsTransfer("TransferNo") & ""%></strong></td>
                            <td><%=Server.HTMLEncode(rsTransfer("NoteName") & "")%></td>
                            <td><%=rsTransfer("NoteType") & ""%></td>
                            <td><%=rsTransfer("RequestQty")%></td>
                            <td><%=rsTransfer("FromWorkshop") & ""%></td>
                            <td><%=rsTransfer("ToWorkshop") & ""%></td>
                            <td><span class="status-badge <%=wtStatusClass%>"><%=wtStatus%></span></td>
                            <td><%=rsTransfer("RequestedBy") & ""%></td>
                            <td class="text-muted"><%=IIF(IsNull(rsTransfer("CreatedAt")) Or rsTransfer("CreatedAt")="","-",Left(rsTransfer("CreatedAt"),10))%></td>
                            <td>
                                <% If wtStatus = "Requested" Then %>
                                <form method="post" style="display:inline;">
                                    <input type="hidden" name="action" value="approve_transfer">
                                    <input type="hidden" name="transfer_id" value="<%=rsTransfer("TransferID")%>">
                                    <button type="submit" name="approve_action" value="approve" class="btn btn-success btn-sm">批准</button>
                                    <button type="submit" name="approve_action" value="reject" class="btn btn-danger btn-sm">拒绝</button>
                                </form>
                                <% ElseIf wtStatus = "Approved" Then %>
                                <form method="post" style="display:inline;">
                                    <input type="hidden" name="action" value="fulfill_transfer">
                                    <input type="hidden" name="transfer_id" value="<%=rsTransfer("TransferID")%>">
                                    <button type="submit" class="btn btn-primary btn-sm">确认收货</button>
                                </form>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rsTransfer.MoveNext
                        Loop
                        rsTransfer.Close
                    End If
                    Set rsTransfer = Nothing
                    If wtRowCount = 0 Then
                    %>
                        <tr><td colspan="10" class="text-center text-muted" style="padding:40px;">暂无调拨记录</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 新建调拨弹窗 -->
    <div id="transferModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>新建香调调拨（半成品车间 → 成品车间）</h3>
                <button class="modal-close" onclick="closeModal('transferModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="create_transfer">
                <div class="form-group">
                    <label>选择香调</label>
                    <select name="note_id" required>
                        <option value="">请选择有库存的香调</option>
                        <%
                        If Not rsAvailNotes Is Nothing Then
                            Do While Not rsAvailNotes.EOF
                        %>
                        <option value="<%=rsAvailNotes("NoteID")%>"><%=Server.HTMLEncode(rsAvailNotes("NoteName") & "")%> [<%=rsAvailNotes("NoteType") & ""%>] (库存:<%=rsAvailNotes("StockQuantity")%>)</option>
                        <%
                                rsAvailNotes.MoveNext
                            Loop
                            rsAvailNotes.Close
                        End If
                        Set rsAvailNotes = Nothing
                        %>
                    </select>
                </div>
                <div class="form-group">
                    <label>调拨数量</label>
                    <input type="number" name="request_qty" required min="1" placeholder="输入调拨数量">
                </div>
                <div class="form-group">
                    <label>备注</label>
                    <textarea name="notes" rows="2" placeholder="调拨备注（可选）"></textarea>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('transferModal')">取消</button>
                    <button type="submit" class="btn btn-primary">提交调拨</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    function openTransferModal() { document.getElementById('transferModal').style.display = 'block'; }
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }
    window.onclick = function(event) { if (event.target.classList.contains('modal')) event.target.style.display = 'none'; }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
