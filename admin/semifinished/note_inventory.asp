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

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "update" Then
        Dim nID, nQty, nMinLevel, nNotes
        nID = SafeNum(Request.Form("note_id"))
        nQty = SafeNum(Request.Form("stock_qty"))
        nMinLevel = SafeNum(Request.Form("min_stock_level"))
        nNotes = Trim(Request.Form("notes"))
        
        If nID > 0 Then
            conn.Execute "UPDATE NoteInventory SET StockQuantity=" & nQty & ", MinStockLevel=" & nMinLevel & ", UpdatedAt=GETDATE() WHERE NoteID=" & nID
            conn.Execute "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES (" & _
                nID & "," & nQty & ",'手动调整','ADJUST','" & SafeSQL(nNotes) & "','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
            Response.Redirect "note_inventory.asp?msg=库存更新成功"
            Response.End
        End If
    
    ElseIf action = "restock" Then
        Dim rNoteID, rAddQty, rNotes
        rNoteID = SafeNum(Request.Form("note_id"))
        rAddQty = SafeNum(Request.Form("add_qty"))
        rNotes = Trim(Request.Form("notes"))
        
        If rNoteID > 0 And rAddQty > 0 Then
            conn.Execute "UPDATE NoteInventory SET StockQuantity=StockQuantity+" & rAddQty & ", LastRestockDate=GETDATE(), UpdatedAt=GETDATE() WHERE NoteID=" & rNoteID
            conn.Execute "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES (" & _
                rNoteID & "," & rAddQty & ",'入库','IN','" & SafeSQL(rNotes) & "','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
            Response.Redirect "note_inventory.asp?msg=入库成功"
            Response.End
        End If
    End If
End If

' ========== 统计 ==========
Dim niTotal, niLowStock, niZeroStock, niTotalQty
niTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory"))
niLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= MinStockLevel AND MinStockLevel > 0"))
niZeroStock = SafeNum(GetScalar("SELECT COUNT(*) FROM NoteInventory WHERE StockQuantity <= 0"))
niTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQuantity),0) FROM NoteInventory"))

' ========== 按类型筛选 ==========
Dim filterType
filterType = Trim(Request.QueryString("type"))

Dim niSQL
niSQL = "SELECT ni.*, fn.NoteName, fn.NoteType, fn.BaseNoteID, fn.PriceAddition, bn.BaseNoteName " & _
    "FROM NoteInventory ni " & _
    "INNER JOIN FragranceNotes fn ON ni.NoteID=fn.NoteID " & _
    "LEFT JOIN BaseNotes bn ON fn.BaseNoteID=bn.BaseNoteID "
If filterType <> "" Then
    niSQL = niSQL & "WHERE fn.NoteType='" & SafeSQL(filterType) & "' "
End If
niSQL = niSQL & "ORDER BY fn.NoteType, fn.NoteName"
Dim rsNI
Set rsNI = conn.Execute(niSQL)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>香调库存 - 半成品生产中心</title>
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
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #FF9800; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(255,152,0,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; display: flex; justify-content: space-between; align-items: center; }
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
        
        .filter-tabs { display: flex; gap: 8px; margin-bottom: 15px; }
        .filter-tab { padding: 6px 16px; border-radius: 16px; font-size: 13px; text-decoration: none; color: #888; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); transition: all 0.2s; }
        .filter-tab:hover { color: #e0e0e0; background: rgba(255,255,255,0.1); }
        .filter-tab.active { color: #fff; background: #FF9800; border-color: #FF9800; }
        

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
        .form-group input, .form-group select { width: 100%; padding: 10px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        .form-group input:focus { outline: none; border-color: #2196F3; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-layer-group"></i> 香调库存管理</h2>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=niTotal%></span><span class="label">香调种类</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=niTotalQty%></span><span class="label">总库存量</span></div>
            <div class="stat-card"><span class="num" style="color:#f44336;"><%=niLowStock%></span><span class="label">低库存预警</span></div>
            <div class="stat-card"><span class="num" style="color:#888;"><%=niZeroStock%></span><span class="label">零库存</span></div>
        </div>
        
        <!-- 类型筛选 -->
        <div class="filter-tabs">
            <a href="note_inventory.asp" class="filter-tab <%=IIF(filterType="","active","")%>">全部</a>
            <a href="?type=Top" class="filter-tab <%=IIF(filterType="Top","active","")%>">前调</a>
            <a href="?type=Middle" class="filter-tab <%=IIF(filterType="Middle","active","")%>">中调</a>
            <a href="?type=Base" class="filter-tab <%=IIF(filterType="Base","active","")%>">后调</a>
        </div>
        
        <!-- 香调库存列表 -->
        <div class="card">
            <div class="card-header">香调库存清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>香调名称</th><th>类型</th><th>基香</th><th>库存量</th><th>最低库存</th><th>价格加成</th><th>最近入库</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    If Not rsNI Is Nothing Then
                        Dim niRowCount : niRowCount = 0
                        Do While Not rsNI.EOF
                            niRowCount = niRowCount + 1
                            Dim nIDRow, nQtyRow, nMinLevelRow, nType
                            nIDRow = rsNI("NoteID")
                            nQtyRow = SafeNum(rsNI("StockQuantity"))
                            nMinLevelRow = SafeNum(rsNI("MinStockLevel"))
                            nType = CStr(rsNI("NoteType") & "")
                            
                            Dim niStockClass, niStockPct
                            If nMinLevelRow > 0 Then
                                niStockPct = (nQtyRow / nMinLevelRow) * 100
                                If nQtyRow <= 0 Then
                                    niStockClass = "danger"
                                ElseIf nQtyRow <= nMinLevelRow Then
                                    niStockClass = "warning"
                                Else
                                    niStockClass = "safe"
                                End If
                            Else
                                niStockPct = 100 : niStockClass = "safe"
                            End If
                            If niStockPct > 100 Then niStockPct = 100
                    %>
                        <tr class="<%=IIF(nMinLevel>0 And nQty<=nMinLevel,"low-stock","")%>">
                            <td><strong><%=Server.HTMLEncode(rsNI("NoteName") & "")%></strong></td>
                            <td>
                                <span style="color:<%=IIF(nType="Top","#f44336",IIF(nType="Middle","#FF9800","#2196F3"))%>">
                                    <%=IIF(nType="Top","前调",IIF(nType="Middle","中调","后调"))%>
                                </span>
                            </td>
                            <td><%=rsNI("BaseNoteName") & ""%></td>
                            <td>
                                <%=nQtyRow%>
                                <div class="stock-bar"><div class="stock-bar-fill <%=niStockClass%>" style="width:<%=niStockPct%>%;"></div></div>
                            </td>
                            <td><%=nMinLevelRow%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsNI("PriceAddition")),2)%></td>
                            <td class="text-muted"><%=IIF(IsNull(rsNI("LastRestockDate")) Or rsNI("LastRestockDate")="","-",Left(rsNI("LastRestockDate"),10))%></td>
                            <td>
                                <button class="btn btn-primary btn-sm" onclick="openEditModal(<%=nIDRow%>,<%=nQtyRow%>,<%=nMinLevelRow%>,'<%=Server.HTMLEncode(rsNI("NoteName") & "")%>')">编辑</button>
                                <button class="btn btn-success btn-sm" onclick="openRestockModal(<%=nIDRow%>,'<%=Server.HTMLEncode(rsNI("NoteName") & "")%>')">入库</button>
                            </td>
                        </tr>
                    <%
                            rsNI.MoveNext
                        Loop
                        rsNI.Close
                    End If
                    Set rsNI = Nothing
                    If niRowCount = 0 Then
                    %>
                        <tr><td colspan="8" class="text-center text-muted" style="padding:40px;">暂无香调库存数据</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 编辑弹窗 -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>编辑库存 - <span id="editNoteName"></span></h3>
                <button class="modal-close" onclick="closeModal('editModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="update">
                <input type="hidden" name="note_id" id="editNoteID">
                <div class="form-group"><label>当前库存量</label><input type="number" name="stock_qty" id="editQty" required></div>
                <div class="form-group"><label>最低库存水平</label><input type="number" name="min_stock_level" id="editMin"></div>
                <div class="form-group"><label>备注</label><input type="text" name="notes" placeholder="调整原因"></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('editModal')">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 入库弹窗 -->
    <div id="restockModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>香调入库存 - <span id="restockNoteName"></span></h3>
                <button class="modal-close" onclick="closeModal('restockModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="restock">
                <input type="hidden" name="note_id" id="restockNoteID">
                <div class="form-group"><label>入库数量</label><input type="number" name="add_qty" required min="1"></div>
                <div class="form-group"><label>备注</label><input type="text" name="notes" placeholder="入库说明"></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('restockModal')">取消</button>
                    <button type="submit" class="btn btn-success">确认入库</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    function openEditModal(id, qty, min, name) {
        document.getElementById('editNoteID').value = id;
        document.getElementById('editQty').value = qty;
        document.getElementById('editMin').value = min;
        document.getElementById('editNoteName').innerText = name;
        document.getElementById('editModal').style.display = 'block';
    }
    function openRestockModal(id, name) {
        document.getElementById('restockNoteID').value = id;
        document.getElementById('restockNoteName').innerText = name;
        document.getElementById('restockModal').style.display = 'block';
    }
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }
    window.onclick = function(event) { if (event.target.classList.contains('modal')) event.target.style.display = 'none'; }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
