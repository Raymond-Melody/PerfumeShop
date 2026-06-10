<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<% Call OpenConnection()

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then
        SafeNum = 0
    Else
        On Error Resume Next
        SafeNum = val * 1.0
        If Err.Number <> 0 Then SafeNum = 0
        On Error GoTo 0
    End If
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = ""
    Else
        SafeSQL = Replace(str, "'", "''")
    End If
End Function

Dim action, msg, msgType
action = Trim(Request.QueryString("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

' ========== POST 处理：基香收货入库 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim postAction
    postAction = Trim(Request.Form("action"))
    
    If postAction = "receive_base_note" Then
        Dim recPurchaseID, recSupplierID, recDetailCount, recNotes
        recPurchaseID = SafeNum(Request.Form("purchase_id"))
        recSupplierID = SafeNum(Request.Form("supplier_id"))
        recDetailCount = SafeNum(Request.Form("detail_count"))
        recNotes = Trim(Request.Form("notes"))
        
        If recPurchaseID > 0 And recDetailCount > 0 Then
            Dim recNo, recReceiptID
            recNo = "BN" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
            
            Dim recTotalQty
            recTotalQty = 0
            Dim ri
            For ri = 1 To recDetailCount
                recTotalQty = recTotalQty + SafeNum(Request.Form("accepted_qty_" & ri))
            Next
            
            On Error Resume Next
            Err.Clear
            Call BeginTransaction()
            If Err.Number <> 0 Then Err.Clear
            
            ' 创建收货单
            Dim sqlReceipt
            sqlReceipt = "INSERT INTO PurchaseReceipts (PurchaseID, ReceiptNo, SupplierID, ReceivedBy, ReceiptDate, Status, TotalReceivedQty, Notes, CreatedAt) VALUES (" & _
                recPurchaseID & ", '" & recNo & "', " & recSupplierID & ", '" & SafeSQL(Session("AdminRealName")) & "', GETDATE(), 'Complete', " & recTotalQty & ", "
            If recNotes <> "" Then
                sqlReceipt = sqlReceipt & "'" & SafeSQL(recNotes) & "'"
            Else
                sqlReceipt = sqlReceipt & "Null"
            End If
            sqlReceipt = sqlReceipt & ", GETDATE())"
            
            conn.Execute sqlReceipt
            
            If Err.Number <> 0 Then
                msg = "创建收货单失败: " & Err.Description
                msgType = "error"
                Call RollbackTransaction()
                Err.Clear
            Else
                ' 获取收货单ID
                Dim rsRecID
                Set rsRecID = conn.Execute("SELECT SCOPE_IDENTITY()")
                recReceiptID = 0
                If Not rsRecID Is Nothing Then
                    If Not rsRecID.EOF Then
                        recReceiptID = rsRecID(0)
                        If IsNull(recReceiptID) Then recReceiptID = 0
                    End If
                    rsRecID.Close
                End If
                Set rsRecID = Nothing
                
                If recReceiptID > 0 Then
                    Dim rj, anyError
                    anyError = False
                    
                    For rj = 1 To recDetailCount
                        Dim rDetailID, rAccepted, rRejected, rReason, rMatName, rMatCode, rMatUnit, rMatPrice
                        Dim rNoteID, rNoteName
                        rDetailID = SafeNum(Request.Form("detail_id_" & rj))
                        rAccepted = SafeNum(Request.Form("accepted_qty_" & rj))
                        rRejected = SafeNum(Request.Form("rejected_qty_" & rj))
                        rReason = Trim(Request.Form("reject_reason_" & rj))
                        rMatName = Trim(Request.Form("item_name_" & rj))
                        rMatCode = Trim(Request.Form("item_code_" & rj))
                        rMatUnit = Trim(Request.Form("unit_" & rj))
                        rMatPrice = SafeNum(Request.Form("unit_price_" & rj))
                        rNoteID = SafeNum(Request.Form("note_id_" & rj))
                        rNoteName = Trim(Request.Form("note_name_" & rj))
                        
                        ' 写入收货明细
                        Dim sqlRecDetail
                        sqlRecDetail = "INSERT INTO PurchaseReceiptDetails (ReceiptID, PurchaseDetailID, ReceivedQty, AcceptedQty, RejectedQty, RejectReason, UnitPrice) VALUES (" & _
                            recReceiptID & ", " & rDetailID & ", " & (rAccepted + rRejected) & ", " & rAccepted & ", " & rRejected & ", "
                        If rReason <> "" Then
                            sqlRecDetail = sqlRecDetail & "'" & SafeSQL(rReason) & "'"
                        Else
                            sqlRecDetail = sqlRecDetail & "Null"
                        End If
                        sqlRecDetail = sqlRecDetail & ", " & rMatPrice & ")"
                        
                        conn.Execute sqlRecDetail
                        
                        If Err.Number <> 0 Then
                            anyError = True
                            Err.Clear
                        Else
                            ' 1) 更新原材料库存
                            If rAccepted > 0 And rMatName <> "" Then
                                Dim rsMat, matID
                                matID = 0
                                Set rsMat = conn.Execute("SELECT MaterialID FROM RawMaterialInventory WHERE ItemName='" & SafeSQL(rMatName) & "'")
                                If Not rsMat Is Nothing Then
                                    If Not rsMat.EOF Then matID = rsMat("MaterialID")
                                    rsMat.Close
                                End If
                                Set rsMat = Nothing
                                
                                If matID > 0 Then
                                    conn.Execute "UPDATE RawMaterialInventory SET StockQty = StockQty + " & rAccepted & ", UnitPrice = " & rMatPrice & ", LastPurchaseDate = GETDATE(), UpdatedAt = GETDATE() WHERE MaterialID=" & matID
                                Else
                                    conn.Execute "INSERT INTO RawMaterialInventory (ItemName, ItemCode, SupplierID, CategoryCode, Unit, StockQty, UnitPrice, LastPurchaseDate, UpdatedAt) VALUES ('" & _
                                        SafeSQL(rMatName) & "', " & IIf(rMatCode <> "", "'" & SafeSQL(rMatCode) & "'", "Null") & ", " & recSupplierID & ", 'BASE_FRAGRANCE', '" & SafeSQL(rMatUnit) & "', " & rAccepted & ", " & rMatPrice & ", GETDATE(), GETDATE())"
                                End If
                                If Err.Number <> 0 Then
                                    anyError = True
                                    Err.Clear
                                End If
                                
                                ' 记录原材料流水
                                If matID > 0 Then
                                    conn.Execute "INSERT INTO InventoryTransactions (NoteID, MaterialID, Quantity, TransactionType, TransactionDirection, UnitCost, Notes, CreatedBy, CreatedAt) VALUES (" & _
                                        "0, " & matID & ", " & rAccepted & ", '基香采购入库', 'IN', " & rMatPrice & ", '收货单" & recNo & "', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
                                End If
                            End If
                            
                            ' 2) 映射到香调 - 更新NoteInventory
                            If rNoteID > 0 And rAccepted > 0 Then
                                Dim rsNI
                                Set rsNI = conn.Execute("SELECT InventoryID FROM NoteInventory WHERE NoteID=" & rNoteID)
                                If Not rsNI Is Nothing Then
                                    If Not rsNI.EOF Then
                                        ' 更新现有库存
                                        conn.Execute "UPDATE NoteInventory SET StockQuantity = StockQuantity + " & rAccepted & ", LastRestockDate = GETDATE(), UpdatedAt = GETDATE() WHERE NoteID=" & rNoteID
                                    Else
                                        ' 新增库存记录
                                        conn.Execute "INSERT INTO NoteInventory (NoteID, StockQuantity, MinStockLevel, LastRestockDate, UpdatedAt) VALUES (" & rNoteID & ", " & rAccepted & ", 10, GETDATE(), GETDATE())"
                                    End If
                                    rsNI.Close
                                End If
                                Set rsNI = Nothing
                                
                                If Err.Number <> 0 Then
                                    anyError = True
                                    Err.Clear
                                Else
                                    ' 记录香调库存流水
                                    conn.Execute "INSERT INTO InventoryTransactions (NoteID, MaterialID, Quantity, TransactionType, TransactionDirection, UnitCost, Notes, CreatedBy, CreatedAt) VALUES (" & _
                                        rNoteID & ", 0, " & rAccepted & ", '基香入库映射', 'IN', " & rMatPrice & ", '从收货单" & recNo & "映射至香调[" & SafeSQL(rNoteName) & "]', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
                                End If
                            End If
                        End If
                    Next
                    
                    If Not anyError Then
                        ' 更新采购订单明细收货数量
                        For rj = 1 To recDetailCount
                            rDetailID = SafeNum(Request.Form("detail_id_" & rj))
                            rAccepted = SafeNum(Request.Form("accepted_qty_" & rj))
                            If rDetailID > 0 Then
                                conn.Execute "UPDATE PurchaseOrderDetails SET ReceivedQty = ReceivedQty + " & rAccepted & " WHERE DetailID=" & rDetailID
                                If Err.Number <> 0 Then
                                    anyError = True
                                    Err.Clear
                                End If
                            End If
                        Next
                        
                        If Not anyError Then
                            conn.Execute "UPDATE PurchaseOrders SET Status='Received', UpdatedAt= GETDATE() WHERE PurchaseID=" & recPurchaseID
                            If Err.Number <> 0 Then
                                anyError = True
                                Err.Clear
                            End If
                        End If
                    End If
                    
                    If Not anyError Then
                        Call CommitTransaction()
                        Response.Redirect "base_note_receiving.asp?msg=基香收货成功！收货单号：" & recNo
                        Response.End
                    Else
                        Call RollbackTransaction()
                        msg = "基香收货失败，数据已回滚，请重试"
                        msgType = "error"
                    End If
                Else
                    Call RollbackTransaction()
                    msg = "获取收货单ID失败"
                    msgType = "error"
                End If
            End If
            On Error GoTo 0
        Else
            msg = "参数错误"
            msgType = "error"
        End If
    End If
End If

' ========== 统计 ==========
Dim statsToday, statsTotalBN
statsToday = 0 : statsTotalBN = 0
On Error Resume Next
Dim rsStats
Set rsStats = conn.Execute("SELECT COUNT(*) FROM PurchaseReceipts WHERE ReceiptNo LIKE 'BN%'")
If Not rsStats Is Nothing And Not rsStats.EOF Then statsTotalBN = CLng(rsStats(0))
rsStats.Close : Set rsStats = Nothing
Set rsStats = conn.Execute("SELECT COUNT(*) FROM PurchaseReceipts WHERE ReceiptNo LIKE 'BN%' AND CAST(ReceiptDate AS DATE)=CAST(GETDATE() AS DATE)")
If Not rsStats Is Nothing And Not rsStats.EOF Then statsToday = CLng(rsStats(0))
rsStats.Close : Set rsStats = Nothing
On Error GoTo 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>基香入库管理 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --card: #16213e; --border: #2a2a4a; --text: #e0e0e0; --accent: #9c27b0; --success: #27ae60; --warning: #f39c12; --danger: #e74c3c; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
        .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
        h1 { font-size: 24px; margin-bottom: 20px; color: var(--text); border-bottom: 2px solid var(--accent); padding-bottom: 10px; }
        h2 { font-size: 18px; margin: 20px 0 10px; color: var(--accent); }
        .message { padding: 12px 20px; border-radius: 6px; margin-bottom: 16px; font-weight: 500; }
        .message.success { background: rgba(39,174,96,0.15); color: #27ae60; border: 1px solid rgba(39,174,96,0.3); }
        .message.error { background: rgba(231,76,60,0.15); color: #e74c3c; border: 1px solid rgba(231,76,60,0.3); }
        .stats-bar { display: flex; gap: 16px; margin-bottom: 24px; }
        .stat-card { flex: 1; background: var(--card); border-radius: 8px; padding: 16px; border: 1px solid var(--border); text-align: center; }
        .stat-card .num { font-size: 28px; font-weight: 700; color: var(--accent); display: block; }
        .stat-card .label { font-size: 13px; color: #999; margin-top: 4px; }
        .card { background: var(--card); border-radius: 8px; border: 1px solid var(--border); margin-bottom: 20px; overflow: hidden; }
        .card-header { padding: 14px 20px; background: rgba(156,39,176,0.08); border-bottom: 1px solid var(--border); font-weight: 600; font-size: 15px; }
        .card-body { padding: 16px 20px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px 12px; background: rgba(156,39,176,0.06); border-bottom: 1px solid var(--border); font-size: 13px; color: #999; font-weight: 600; }
        td { padding: 10px 12px; border-bottom: 1px solid var(--border); font-size: 14px; }
        tr:hover { background: rgba(255,255,255,0.02); }
        .status { padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .status-submitted { background: rgba(255,152,0,0.2); color: #FF9800; }
        .status-ordered { background: rgba(33,150,243,0.2); color: #2196F3; }
        .status-approved { background: rgba(0,188,212,0.2); color: #00BCD4; }
        .status-partial { background: rgba(255,193,7,0.2); color: #FFC107; }
        input[type="text"], input[type="number"], textarea, select { width: 100%; padding: 9px 12px; background: var(--input-bg); border: 1px solid var(--border); border-radius: 5px; color: var(--text); font-size: 14px; }
        input:focus, select:focus, textarea:focus { border-color: var(--accent); outline: none; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }
        .detail-header { display: grid; grid-template-columns: 2fr 1fr 1.5fr 1fr 1fr; gap: 8px; padding: 8px 12px; background: rgba(255,255,255,0.03); font-size: 13px; color: #999; font-weight: 600; margin-bottom: 4px; }
        .detail-row { display: grid; grid-template-columns: 2fr 1fr 1.5fr 1fr 1fr; gap: 8px; padding: 8px 12px; border-bottom: 1px solid var(--border); align-items: center; }
        .detail-row input, .detail-row select { width: 100%; padding: 6px 8px; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .text-muted { color: #999; font-size: 13px; }
        .mb-2 { margin-bottom: 16px; }
        .mt-2 { margin-top: 16px; }
        .empty-state { text-align: center; padding: 40px; color: #999; }
        .note-type-badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
        .type-top { background: rgba(76,175,80,0.2); color: #81c784; }
        .type-middle { background: rgba(255,152,0,0.2); color: #ffb74d; }
        .type-base { background: rgba(156,39,176,0.2); color: #ce93d8; }
        /* V9: 隐藏不合格列 */
        .reject-cell.hidden { display: none; }
        .reject-visible .reject-cell { display: block; }
    </style>
</head>
<body data-theme="purchase-dark">
<div class="container">
    <h1>🌸 基香入库管理（香调映射）</h1>
    
    <% If msg <> "" Then %>
    <div class="message <%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
    <% End If %>
    
    <div class="stats-bar">
        <div class="stat-card"><span class="num"><%=statsTotalBN%></span><span class="label">基香收货单总数</span></div>
        <div class="stat-card"><span class="num"><%=statsToday%></span><span class="label">今日收货</span></div>
    </div>
    
    <%
    Dim purchaseId
    purchaseId = Trim(Request.QueryString("purchase_id"))
    
    ' ========== 收货录入页面 ==========
    If action = "receive" And purchaseId <> "" And IsNumeric(purchaseId) Then
        Dim rsBN, bnStatus
        Set rsBN = conn.Execute("SELECT po.PurchaseID, po.PurchaseNo, po.SupplierID, po.OrderType, po.CategoryCode, CAST(ISNULL(po.TotalAmount,0) AS FLOAT) as TotalAmount, po.Status, po.CreatedBy, po.ExpectedDate, po.Remarks, s.SupplierName FROM PurchaseOrders po LEFT JOIN Suppliers s ON po.SupplierID=s.SupplierID WHERE po.PurchaseID=" & CLng(purchaseId) & " AND po.CategoryCode='BASE'")
        If Not rsBN Is Nothing And Not rsBN.EOF Then
            bnStatus = CStr(rsBN("Status"))
    %>
    <div class="card">
        <div class="card-header">基香收货录入 - <%=rsBN("PurchaseNo")%> | 供应商：<%=rsBN("SupplierName")%></div>
        <div class="card-body">
            <form method="post" id="receiptForm">
                <input type="hidden" name="action" value="receive_base_note">
                <input type="hidden" name="purchase_id" value="<%=purchaseId%>">
                <input type="hidden" name="supplier_id" value="<%=rsBN("SupplierID")%>">
                
                <h2>收货明细 & 香调映射</h2>
                <p class="text-muted mb-2">请为每条物料选择对应的香调，收货后将自动更新香调库存</p>
                <!-- V9: 快捷操作按钮 -->
                <div style="display:flex;gap:10px;margin-bottom:12px;flex-wrap:wrap;">
                    <button type="button" class="btn btn-outline btn-sm" onclick="suggestAllMappings()" title="根据物料名称智能匹配香调">
                        <span style="font-size:13px;">🔍</span> 建议匹配香调
                    </button>
                    <button type="button" class="btn btn-outline btn-sm" onclick="setAllAccepted()" title="将所有合格数量设为剩余待收数量">
                        <span style="font-size:13px;">📦</span> 全部合格入库
                    </button>
                    <button type="button" class="btn btn-outline btn-sm" id="toggleRejectBtn" onclick="toggleRejectColumns()" title="显示/隐藏不合格数量">
                        <span style="font-size:13px;">⚠️</span> 显示不合格
                    </button>
                </div>
                <div class="detail-header">
                    <span>物料名称</span><span>采购数量</span><span>映射香调</span><span>合格数量</span><span>不合格数量</span>
                </div>
                
                <% 
                Dim rsBND, detailIdx
                detailIdx = 0
                Set rsBND = conn.Execute("SELECT DetailID, PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, CAST(ISNULL(UnitPrice,0) AS FLOAT) as UnitPrice, CAST(ISNULL(TotalPrice,0) AS FLOAT) as TotalPrice, ReceivedQty FROM PurchaseOrderDetails WHERE PurchaseID=" & CLng(purchaseId) & " ORDER BY DetailID")
                ' 预加载香调列表
                Dim rsNotesSelect
                Set rsNotesSelect = conn.Execute("SELECT NoteID, NoteName, NoteType FROM FragranceNotes WHERE IsBaseNote=1 ORDER BY NoteType, NoteName")
                Dim noteOptions
                noteOptions = "<option value=''>-- 选择香调 --</option>"
                If Not rsNotesSelect Is Nothing Then
                    Do While Not rsNotesSelect.EOF
                        noteOptions = noteOptions & "<option value='" & rsNotesSelect("NoteID") & "'>" & rsNotesSelect("NoteName") & " [" & rsNotesSelect("NoteType") & "]</option>"
                        rsNotesSelect.MoveNext
                    Loop
                    rsNotesSelect.Close
                End If
                Set rsNotesSelect = Nothing
                
                If Not rsBND Is Nothing Then
                    Do While Not rsBND.EOF
                        detailIdx = detailIdx + 1
                        Dim bPurchasedQty, bReceivedQty, bRemaining
                        bPurchasedQty = SafeNum(rsBND("Quantity"))
                        bReceivedQty = SafeNum(rsBND("ReceivedQty"))
                        bRemaining = bPurchasedQty - bReceivedQty
                        If bRemaining < 0 Then bRemaining = 0
                %>
                <div class="detail-row">
                    <div>
                        <input type="hidden" name="detail_id_<%=detailIdx%>" value="<%=rsBND("DetailID")%>">
                        <input type="hidden" name="item_name_<%=detailIdx%>" value="<%=Server.HTMLEncode(rsBND("ItemName"))%>">
                        <input type="hidden" name="item_code_<%=detailIdx%>" value="<%=Server.HTMLEncode(rsBND("ItemCode") & "")%>">
                        <input type="hidden" name="unit_<%=detailIdx%>" value="<%=Server.HTMLEncode(rsBND("Unit") & "")%>">
                        <input type="hidden" name="unit_price_<%=detailIdx%>" value="<%=rsBND("UnitPrice")%>">
                        <input type="hidden" name="note_name_<%=detailIdx%>" id="note_name_<%=detailIdx%>" value="">
                        <input type="hidden" id="raw_item_name_<%=detailIdx%>" value="<%=Server.HTMLEncode(rsBND("ItemName"))%>">
                        <strong><%=Server.HTMLEncode(rsBND("ItemName"))%></strong>
                        <div class="text-muted" style="font-size:12px">单价: ¥<%=FormatNumber(CDbl("0" & rsBND("UnitPrice")),2)%> | 剩余: <%=bRemaining%></div>
                    </div>
                    <div class="text-center"><%=bPurchasedQty%></div>
                    <div style="display:flex;gap:4px;align-items:center;">
                        <select name="note_id_<%=detailIdx%>" id="note_select_<%=detailIdx%>" onchange="document.getElementById('note_name_<%=detailIdx%>').value=this.options[this.selectedIndex].text" style="flex:1;">
                            <%=noteOptions%>
                        </select>
                        <button type="button" class="btn btn-outline btn-sm" onclick="suggestMapping(<%=detailIdx%>)" title="智能匹配香调" style="padding:4px 8px;">🔍</button>
                    </div>
                    <div><input type="number" name="accepted_qty_<%=detailIdx%>" id="accepted_qty_<%=detailIdx%>" value="<%=bRemaining%>" min="0" max="<%=bRemaining%>" step="0.01"></div>
                    <div class="reject-cell"><input type="number" name="rejected_qty_<%=detailIdx%>" value="0" min="0" max="<%=bRemaining%>" step="0.01">
                        <input type="text" name="reject_reason_<%=detailIdx%>" placeholder="原因" style="font-size:12px; margin-top:2px"></div>
                </div>
                <%
                        rsBND.MoveNext
                    Loop
                    rsBND.Close
                End If
                Set rsBND = Nothing
                %>
                <input type="hidden" name="detail_count" value="<%=detailIdx%>">
                
                <div class="mt-2 grid-2">
                    <div><label class="text-muted">备注</label><textarea name="notes" rows="3" placeholder="收货备注..."></textarea></div>
                    <div style="text-align:right; padding-top:20px;"><span class="text-muted">基香收货将自动映射至香调库存</span></div>
                </div>
                
                <div class="mt-2" style="display:flex; gap:12px; justify-content:flex-end">
                    <a href="base_note_receiving.asp" class="btn btn-outline">返回列表</a>
                    <button type="submit" class="btn btn-success" onclick="return confirm('确认提交基香收货？收货后将自动更新香调库存。')">确认收货 & 映射香调</button>
                </div>
            </form>
        </div>
    </div>
    <%
        Else
            Response.Write "<div class='message error'>基香采购订单不存在或非BASE分类</div>"
        End If
        If Not rsBN Is Nothing Then rsBN.Close
        Set rsBN = Nothing
    %>
    
    <%
    ' ========== 收货记录 + 待收货列表 ==========
    Else
    %>
    <!-- 待收货基香采购订单 -->
    <div class="card">
        <div class="card-header">待收货基香采购订单</div>
        <div class="card-body" style="overflow-x:auto;">
            <table>
                <thead>
                    <tr><th>采购单号</th><th>供应商</th><th>订单日期</th><th>金额</th><th>状态</th><th>操作</th></tr>
                </thead>
                <tbody>
                <%
                Dim sqlBNPending, rsBNPending
                sqlBNPending = "SELECT po.*, s.SupplierName FROM PurchaseOrders po LEFT JOIN Suppliers s ON po.SupplierID=s.SupplierID WHERE po.CategoryCode='BASE' AND po.Status IN ('Submitted','FinanceApproved','Ordered','PartialReceived') ORDER BY po.PurchaseID DESC"
                Set rsBNPending = conn.Execute(sqlBNPending)
                If Not rsBNPending Is Nothing Then
                    Dim bnPendingCount
                    bnPendingCount = 0
                    Do While Not rsBNPending.EOF And bnPendingCount < 50
                        bnPendingCount = bnPendingCount + 1
                %>
                    <tr>
                        <td><strong><%=rsBNPending("PurchaseNo")%></strong></td>
                        <td><%=rsBNPending("SupplierName") & ""%></td>
                        <td><%=rsBNPending("OrderDate") & ""%></td>
                        <td class="text-right">¥<%=FormatNumber(CDbl("0" & rsBNPending("TotalAmount")), 2)%></td>
                        <td>
                            <% If CStr(rsBNPending("Status")) = "PartialReceived" Then %>
                                <span class="status status-partial">部分收货</span>
                            <% ElseIf CStr(rsBNPending("Status")) = "Ordered" Then %>
                                <span class="status status-ordered">已下单</span>
                            <% ElseIf CStr(rsBNPending("Status")) = "FinanceApproved" Then %>
                                <span class="status status-approved">已审批</span>
                            <% Else %>
                                <span class="status status-submitted"><%=rsBNPending("Status")%></span>
                            <% End If %>
                        </td>
                        <td>
                            <a href="base_note_receiving.asp?action=receive&purchase_id=<%=rsBNPending("PurchaseID")%>" class="btn btn-primary btn-sm">基香收货</a>
                        </td>
                    </tr>
                <%
                        rsBNPending.MoveNext
                    Loop
                    rsBNPending.Close
                    If bnPendingCount = 0 Then
                %>
                    <tr><td colspan="6" class="text-center text-muted" style="padding:40px">暂无待收货基香采购订单</td></tr>
                <% End If %>
                <% End If
                Set rsBNPending = Nothing %>
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- 基香收货历史 -->
    <div class="card">
        <div class="card-header">基香收货记录</div>
        <div class="card-body" style="overflow-x:auto;">
            <table>
                <thead>
                    <tr><th>收货单号</th><th>采购单号</th><th>供应商</th><th>收货日期</th><th>收货数量</th><th>操作</th></tr>
                </thead>
                <tbody>
                <%
                Dim sqlBNRecs, rsBNRecs
                sqlBNRecs = "SELECT pr.*, po.PurchaseNo, s.SupplierName FROM (PurchaseReceipts pr LEFT JOIN PurchaseOrders po ON pr.PurchaseID=po.PurchaseID) LEFT JOIN Suppliers s ON pr.SupplierID=s.SupplierID WHERE pr.ReceiptNo LIKE 'BN%' ORDER BY pr.ReceiptID DESC"
                Set rsBNRecs = conn.Execute(sqlBNRecs)
                If Not rsBNRecs Is Nothing Then
                    Dim bnRecCount
                    bnRecCount = 0
                    Do While Not rsBNRecs.EOF And bnRecCount < 50
                        bnRecCount = bnRecCount + 1
                %>
                    <tr>
                        <td><strong><%=rsBNRecs("ReceiptNo")%></strong></td>
                        <td><%=rsBNRecs("PurchaseNo") & ""%></td>
                        <td><%=rsBNRecs("SupplierName") & ""%></td>
                        <td><%=rsBNRecs("ReceiptDate") & ""%></td>
                        <td><%=rsBNRecs("TotalReceivedQty")%></td>
                        <td><a href="receiving.asp" class="btn btn-outline btn-sm">查看详情</a></td>
                    </tr>
                <%
                        rsBNRecs.MoveNext
                    Loop
                    rsBNRecs.Close
                    If bnRecCount = 0 Then
                %>
                    <tr><td colspan="6" class="text-center text-muted" style="padding:40px">暂无基香收货记录</td></tr>
                <% End If %>
                <% End If
                Set rsBNRecs = Nothing %>
                </tbody>
            </table>
        </div>
    </div>
    <% End If %>
</div>
<!-- V9: 交互优化脚本 -->
<script>
var rejectVisible = false;

// 页面加载时隐藏不合格列
document.addEventListener('DOMContentLoaded', function() {
    hideRejectCells();
});

function hideRejectCells() {
    var cells = document.querySelectorAll('.reject-cell');
    for (var i = 0; i < cells.length; i++) {
        cells[i].classList.add('hidden');
    }
}

function showRejectCells() {
    var cells = document.querySelectorAll('.reject-cell');
    for (var i = 0; i < cells.length; i++) {
        cells[i].classList.remove('hidden');
    }
}

// 切换不合格列显隐
function toggleRejectColumns() {
    rejectVisible = !rejectVisible;
    var btn = document.getElementById('toggleRejectBtn');
    if (rejectVisible) {
        showRejectCells();
        btn.innerHTML = '<span style="font-size:13px;">⚠️</span> 隐藏不合格';
    } else {
        hideRejectCells();
        btn.innerHTML = '<span style="font-size:13px;">⚠️</span> 显示不合格';
    }
}

// 智能匹配单个行的香调
function suggestMapping(idx) {
    var itemName = document.getElementById('raw_item_name_' + idx);
    if (!itemName) return;
    var name = itemName.value.toLowerCase().trim();
    if (!name) return;
    
    var select = document.getElementById('note_select_' + idx);
    if (!select) return;
    
    var bestMatch = null;
    var bestScore = 0;
    
    for (var i = 0; i < select.options.length; i++) {
        var opt = select.options[i];
        if (!opt.value) continue; // skip placeholder
        var optText = opt.text.toLowerCase();
        
        // 计算匹配分数
        var score = 0;
        if (optText === name) {
            score = 100;
        } else if (optText.indexOf(name) !== -1 || name.indexOf(optText.replace(/\[.*\]/g,'').trim()) !== -1) {
            score = 80;
        } else {
            // 分词匹配
            var nameWords = name.split(/[\s\-_\(\)]+/);
            var optWords = optText.split(/[\s\-_\(\)\[\]]+/);
            for (var w = 0; w < nameWords.length; w++) {
                var nw = nameWords[w];
                if (nw.length < 2) continue;
                for (var ow = 0; ow < optWords.length; ow++) {
                    if (optWords[ow] === nw) { score += 30; break; }
                    if (optWords[ow].indexOf(nw) !== -1 || nw.indexOf(optWords[ow]) !== -1) { score += 15; break; }
                }
            }
        }
        
        if (score > bestScore) {
            bestScore = score;
            bestMatch = opt.value;
        }
    }
    
    if (bestMatch && bestScore >= 30) {
        select.value = bestMatch;
        // 更新隐藏的note_name字段
        var noteNameHidden = document.getElementById('note_name_' + idx);
        if (noteNameHidden) {
            noteNameHidden.value = select.options[select.selectedIndex].text;
        }
    }
}

// 全部建议匹配
function suggestAllMappings() {
    var allSelects = document.querySelectorAll('[id^="note_select_"]');
    var matchCount = 0;
    for (var i = 0; i < allSelects.length; i++) {
        var idx = allSelects[i].id.replace('note_select_', '');
        suggestMapping(parseInt(idx));
        if (allSelects[i].value) matchCount++;
    }
    alert('智能匹配完成：' + matchCount + '/' + allSelects.length + ' 条已匹配香调');
}

// 全部设为合格入库（取剩余最大数量）
function setAllAccepted() {
    var allAccepted = document.querySelectorAll('[id^="accepted_qty_"]');
    for (var i = 0; i < allAccepted.length; i++) {
        var max = parseFloat(allAccepted[i].max) || 0;
        allAccepted[i].value = max;
    }
}
</script>
</body>
</html>
<%
Call CloseConnection()
%>
