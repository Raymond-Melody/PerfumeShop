<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection()
%>
<!--#include file="includes/db_setup.asp"-->
<%

Function GetStatusName(statusCode)
    Select Case statusCode
        Case "Ordered"        : GetStatusName = "已下单"
        Case "PartialReceived": GetStatusName = "部分收货"
        Case "Received"       : GetStatusName = "已收货"
        Case "Completed"      : GetStatusName = "已完成"
        Case Else             : GetStatusName = statusCode
    End Select
End Function

' ========== 消息 ==========
Dim msg, msgType
msg = ""
msgType = "success"

' ========== 指定采购单ID ==========
Dim targetPID : targetPID = SafeNum(Request.QueryString("purchase_id"))

' ========== POST处理：执行收货 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        msg = "安全令牌验证失败"
        msgType = "error"
    Else
        Dim postAction : postAction = Trim(Request.Form("action"))
        
        If postAction = "create_receipt" Then
            Dim recPID : recPID = SafeNum(Request.Form("purchase_id"))
            Dim recSID : recSID = SafeNum(Request.Form("supplier_id"))
            Dim recDetailCnt : recDetailCnt = SafeNum(Request.Form("detail_count"))
            Dim recNotes : recNotes = SafeSQL(Trim(Request.Form("notes")))
            
            If recPID > 0 And recDetailCnt > 0 Then
                Dim recNo : recNo = "FBRE" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & Right("0" & Hour(Now),2) & Right("0" & Minute(Now),2) & Right("0" & Second(Now),2)
                Dim totalRecQty : totalRecQty = 0
                
                Call BeginTransaction()
                
                Dim insRecSQL : insRecSQL = "INSERT INTO FixedBrandReceipts (PurchaseID, ReceiptNo, SupplierID, ReceivedBy, Notes) VALUES (" & _
                    recPID & ", '" & recNo & "', " & recSID & ", '" & SafeSQL(Session("AdminName")) & "', '" & recNotes & "')"
                
                If ExecuteNonQuery(insRecSQL) Then
                    Dim newRecID : newRecID = SafeNum(GetScalar("SELECT MAX(ReceiptID) FROM FixedBrandReceipts"))
                    
                    If newRecID > 0 Then
Dim ri, allOK : allOK = True
For ri = 1 To recDetailCnt
    Dim rDetailID : rDetailID = SafeNum(Request.Form("rec_detail_id_" & ri))
    Dim rFPID : rFPID = SafeNum(Request.Form("rec_fpid_" & ri))
    Dim rAccept : rAccept = SafeNum(Request.Form("rec_accept_" & ri))
    Dim rReject : rReject = SafeNum(Request.Form("rec_reject_" & ri))
    Dim rReason : rReason = SafeSQL(Trim(Request.Form("rec_reason_" & ri)))
    Dim rPrice : rPrice = SafeNum(Request.Form("rec_price_" & ri))
    
    ' 服务器端验证：收货数量不能超过待收数量
    If rDetailID > 0 Then
        Dim rOrderedQty : rOrderedQty = SafeNum(GetScalar("SELECT ISNULL(Quantity,0) FROM FixedBrandPurchaseDetails WHERE DetailID=" & rDetailID))
        Dim rAlreadyRec : rAlreadyRec = SafeNum(GetScalar("SELECT ISNULL(ReceivedQty,0) FROM FixedBrandPurchaseDetails WHERE DetailID=" & rDetailID))
        Dim rRemaining : rRemaining = rOrderedQty - rAlreadyRec
        If rRemaining < 0 Then rRemaining = 0
        If rAccept > rRemaining Then
            msg = "合格收货数量(" & rAccept & ")超过待收数量(" & rRemaining & ")，收货失败"
            msgType = "error"
            allOK = False
            Exit For
        End If
    End If
    
    If rAccept > 0 Or rReject > 0 Then
                                Dim insRecDetSQL : insRecDetSQL = "INSERT INTO FixedBrandReceiptDetails (ReceiptID, DetailID, FixedProductID, AcceptedQty, RejectedQty, RejectReason, UnitPrice) VALUES (" & _
                                    newRecID & ", " & rDetailID & ", " & rFPID & ", " & rAccept & ", " & rReject & ", '" & rReason & "', " & rPrice & ")"
                                
                                If ExecuteNonQuery(insRecDetSQL) Then
                                    totalRecQty = totalRecQty + rAccept
                                    
                                    ' 更新订单明细的已收货数量
                                    Call ExecuteNonQuery("UPDATE FixedBrandPurchaseDetails SET ReceivedQty = ReceivedQty + " & rAccept & " WHERE DetailID=" & rDetailID)
                                    
                                    ' 更新库存
                                    Dim invExists : invExists = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandInventory WHERE FixedProductID=" & rFPID))
                                    If invExists > 0 Then
                                        ' 加权平均成本计算
                                        Dim curStock : curStock = SafeNum(GetScalar("SELECT ISNULL(StockQty,0) FROM FixedBrandInventory WHERE FixedProductID=" & rFPID))
                                        Dim curAvgCost : curAvgCost = SafeNum(GetScalar("SELECT ISNULL(AvgUnitCost,0) FROM FixedBrandInventory WHERE FixedProductID=" & rFPID))
                                        Dim newTotalValue : newTotalValue = curStock * curAvgCost + rAccept * rPrice
                                        Dim newAvgCost : newAvgCost = 0
                                        If (curStock + rAccept) > 0 Then newAvgCost = newTotalValue / (curStock + rAccept)
                                        
                                        Call ExecuteNonQuery("UPDATE FixedBrandInventory SET StockQty=StockQty+" & rAccept & ", AvgUnitCost=" & newAvgCost & ", LastPurchasePrice=" & rPrice & ", LastPurchaseDate=GETDATE(), LastPurchaseID=" & recPID & ", TotalPurchased=TotalPurchased+" & rAccept & ", UpdatedAt=GETDATE() WHERE FixedProductID=" & rFPID)
                                        
                                        ' ========== 同步到 Products 表和 ProductInventory 表 ==========
                                        Dim prodID : prodID = SafeNum(GetScalar("SELECT ISNULL(ProductID,0) FROM FixedBrandProducts WHERE FixedProductID=" & rFPID))
                                        If prodID > 0 Then
                                            ' 更新 Products.UnitCost 为加权平均采购成本
                                            Call ExecuteNonQuery("UPDATE Products SET UnitCost=" & newAvgCost & " WHERE ProductID=" & prodID)
                                            ' 更新 ProductInventory 库存
                                            Dim piExists : piExists = SafeNum(GetScalar("SELECT COUNT(*) FROM ProductInventory WHERE ProductID=" & prodID))
                                            If piExists > 0 Then
                                                Call ExecuteNonQuery("UPDATE ProductInventory SET StockQty = StockQty + " & rAccept & ", UpdatedAt = GETDATE() WHERE ProductID=" & prodID)
                                            Else
                                                Call ExecuteNonQuery("INSERT INTO ProductInventory (ProductID, StockQty) VALUES (" & prodID & ", " & rAccept & ")")
                                            End If
                                        End If
                                    Else
                                        Call ExecuteNonQuery("INSERT INTO FixedBrandInventory (FixedProductID, ProductCode, ProductName, Specification, StockQty, AvgUnitCost, LastPurchasePrice, LastPurchaseDate, LastPurchaseID) VALUES (" & rFPID & ", '', '', '', " & rAccept & ", " & rPrice & ", " & rPrice & ", GETDATE(), " & recPID & ")")
                                    End If
                                Else
                                    allOK = False
                                    Exit For
                                End If
                            End If
                        Next
                        
                        If allOK Then
                            Call ExecuteNonQuery("UPDATE FixedBrandReceipts SET TotalReceivedQty=" & totalRecQty & " WHERE ReceiptID=" & newRecID)
                            
                            ' 更新订单状态：检查是否全部收货完成
                            Dim allReceived : allReceived = True
                            Dim rsCheck : Set rsCheck = conn.Execute("SELECT DetailID, Quantity, ISNULL(ReceivedQty,0) AS ReceivedQty FROM FixedBrandPurchaseDetails WHERE PurchaseID=" & recPID)
                            If Not rsCheck Is Nothing Then
                                Do While Not rsCheck.EOF
                                    If SafeNum(rsCheck("ReceivedQty")) < SafeNum(rsCheck("Quantity")) Then
                                        allReceived = False
                                    End If
                                    rsCheck.MoveNext
                                Loop
                                rsCheck.Close
                            End If
                            Set rsCheck = Nothing
                            
                            If allReceived Then
                                Call ExecuteNonQuery("UPDATE FixedBrandPurchaseOrders SET Status='Received', UpdatedAt=GETDATE() WHERE PurchaseID=" & recPID)
                            Else
                                Call ExecuteNonQuery("UPDATE FixedBrandPurchaseOrders SET Status='PartialReceived', UpdatedAt=GETDATE() WHERE PurchaseID=" & recPID)
                            End If
                            
                            Call CommitTransaction()
                            msg = "收货单 " & recNo & " 创建成功，共入库 " & totalRecQty & " 件"
                            msgType = "success"
                        Else
                            Call RollbackTransaction()
                            msg = "收货明细添加失败，已回滚"
                            msgType = "error"
                        End If
                    Else
                        Call RollbackTransaction()
                        msg = "收货单创建失败"
                        msgType = "error"
                    End If
                Else
                    Call RollbackTransaction()
                    msg = "收货单创建失败"
                    msgType = "error"
                End If
            End If
        End If
    End If
End If

' ========== 查询待收货订单 ==========
Dim sqlReceiving : sqlReceiving = "SELECT * FROM FixedBrandPurchaseOrders WHERE Status IN ('Ordered','PartialReceived')"
If targetPID > 0 Then
    sqlReceiving = "SELECT * FROM FixedBrandPurchaseOrders WHERE PurchaseID=" & targetPID
End If
sqlReceiving = sqlReceiving & " ORDER BY PurchaseID DESC"

Dim rsReceiving : Set rsReceiving = conn.Execute(sqlReceiving)

' ========== 统计 ==========
Dim pendingRecCount : pendingRecCount = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandPurchaseOrders WHERE Status IN ('Ordered','PartialReceived')"))
Dim recentRecCount : recentRecCount = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandReceipts"))
Dim recentRecQty : recentRecQty = SafeNum(GetScalar("SELECT ISNULL(SUM(TotalReceivedQty),0) FROM FixedBrandReceipts WHERE ReceiptDate >= DATEADD(DAY,-30,GETDATE())"))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>品牌定香收货入库 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 270px; padding: 25px; min-height: 100vh; }
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .page-title { font-size: 20px; font-weight: 600; color: #fff; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #FF9800; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 18px; border: 1px solid rgba(255,255,255,0.05); }
        .stat-icon { width: 40px; height: 40px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 16px; margin-bottom: 10px; }
        .stat-value { font-size: 22px; font-weight: 700; color: #fff; }
        .stat-label { font-size: 12px; color: #888; margin-top: 4px; }
        
        .order-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 20px; margin-bottom: 15px; border: 1px solid rgba(255,255,255,0.05); }
        .order-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .order-header h3 { color: #fff; font-size: 16px; margin: 0; }
        
        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th, .data-table td { padding: 10px 12px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .data-table th { color: #888; font-size: 11px; font-weight: 600; }
        .data-table td { color: #ccc; }
        
        .rec-input { padding: 6px 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1a1a2e; color: #e0e0e0; width: 70px; font-size: 13px; text-align: center; }
        .rec-input-wide { width: 120px; }
        .rec-total { font-weight: 600; color: #4CAF50; font-size: 15px; }
        
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 500; }
        .status-ordered { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .status-partial { background: rgba(0,188,212,0.2); color: #00BCD4; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="../includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-clipboard-check"></i> 品牌定香收货入库</h2>
            <div class="breadcrumb" style="font-size:13px;color:#888;">
                <a href="index.asp" style="color:#FF9800;text-decoration:none;">品牌定香采购</a> / 收货入库
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div style="padding:12px 20px; border-radius:8px; margin-bottom:20px; font-size:14px; background:<%=IIf(msgType="success","rgba(76,175,80,0.15)","rgba(244,67,54,0.15)")%>; color:<%=IIf(msgType="success","#4CAF50","#F44336")%>; border:1px solid <%=IIf(msgType="success","rgba(76,175,80,0.3)","rgba(244,67,54,0.3)")%>;">
            <i class="fas fa-<%=IIf(msgType="success","check-circle","exclamation-circle")%>"></i> <%= msg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-truck-loading"></i></div>
                <div class="stat-value"><%= pendingRecCount %></div>
                <div class="stat-label">待收货订单</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-clipboard-list"></i></div>
                <div class="stat-value"><%= recentRecCount %></div>
                <div class="stat-label">历史收货单</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#2196F3,#1565C0);"><i class="fas fa-boxes"></i></div>
                <div class="stat-value"><%= recentRecQty %></div>
                <div class="stat-label">近30天入库量</div>
            </div>
        </div>
        
        <% If Not rsReceiving Is Nothing Then
            If Not rsReceiving.EOF Then
                Do While Not rsReceiving.EOF
                    Dim pid : pid = SafeNum(rsReceiving("PurchaseID"))
                    Dim pStatus : pStatus = CStr(rsReceiving("Status"))
                    Dim rsDetails : Set rsDetails = conn.Execute("SELECT d.*, p.Status FROM FixedBrandPurchaseDetails d JOIN FixedBrandPurchaseOrders p ON d.PurchaseID=p.PurchaseID WHERE d.PurchaseID=" & pid & " ORDER BY d.DetailID")
        %>
        <div class="order-card">
            <div class="order-header">
                <h3>
                    <i class="fas fa-file-invoice" style="color:#FF9800;"></i>
                    <%= Server.HTMLEncode(CStr(rsReceiving("PurchaseNo"))) %>
                </h3>
                <div style="display:flex;gap:10px;align-items:center;">
                    <span class="status-badge <%= IIf(pStatus="Ordered","status-ordered","status-partial") %>"><%= GetStatusName(pStatus) %></span>
                    <span style="color:#888;font-size:13px;"><i class="fas fa-truck"></i> <%= Server.HTMLEncode(CStr(rsReceiving("SupplierName") & "")) %></span>
                </div>
            </div>
            
            <% If pStatus = "Ordered" Or pStatus = "PartialReceived" Then %>
            <form method="post">
                <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                <input type="hidden" name="action" value="create_receipt">
                <input type="hidden" name="purchase_id" value="<%= pid %>">
                <input type="hidden" name="supplier_id" value="<%= SafeNum(rsReceiving("SupplierID")) %>">
                
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>产品名称</th>
                            <th>规格</th>
                            <th>订购数量</th>
                            <th>已收数量</th>
                            <th>单价</th>
                            <th>本次合格</th>
                            <th>本次拒收</th>
                            <th>拒收原因</th>
                            <th>小计</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If Not rsDetails Is Nothing Then
                            Dim detailIdx : detailIdx = 0
                            Do While Not rsDetails.EOF
                                detailIdx = detailIdx + 1
                                Dim orderQty : orderQty = SafeNum(rsDetails("Quantity"))
                                Dim rcvQty : rcvQty = SafeNum(rsDetails("ReceivedQty"))
                                Dim remainingQty : remainingQty = orderQty - rcvQty
                                If remainingQty < 0 Then remainingQty = 0
                        %>
                        <tr>
                            <td><%= Server.HTMLEncode(CStr(rsDetails("ProductName"))) %></td>
                            <td><%= Server.HTMLEncode(CStr(rsDetails("Specification") & "")) %></td>
                            <td><%= orderQty %></td>
                            <td style="color:<%=IIf(rcvQty>=orderQty,"#4CAF50","#FF9800")%>;"><%= rcvQty %></td>
                            <td>¥<%= FormatNumber(SafeNum(rsDetails("UnitPrice")), 2) %></td>
                            <td>
                                <input type="hidden" name="rec_detail_id_<%= detailIdx %>" value="<%= SafeNum(rsDetails("DetailID")) %>">
                                <input type="hidden" name="rec_fpid_<%= detailIdx %>" value="<%= SafeNum(rsDetails("FixedProductID")) %>">
                                <input type="hidden" name="rec_price_<%= detailIdx %>" value="<%= SafeNum(rsDetails("UnitPrice")) %>">
                                <input type="number" name="rec_accept_<%= detailIdx %>" class="rec-input" value="<%= remainingQty %>" min="0" max="<%= remainingQty %>" onchange="updateRowTotal(this, <%= detailIdx %>, <%= SafeNum(rsDetails("UnitPrice")) %>, <%= pid %>)">
                            </td>
                            <td><input type="number" name="rec_reject_<%= detailIdx %>" class="rec-input" value="0" min="0"></td>
                            <td><input type="text" name="rec_reason_<%= detailIdx %>" class="rec-input rec-input-wide" placeholder="拒收原因"></td>
                            <td><span class="rec-total" id="row-total-<%= pid %>-<%= detailIdx %>">¥<%= FormatNumber(remainingQty * SafeNum(rsDetails("UnitPrice")), 2) %></span></td>
                        </tr>
                        <% 
                                rsDetails.MoveNext
                            Loop
                        %>
                        <input type="hidden" name="detail_count" value="<%= detailIdx %>">
                        <% End If %>
                    </tbody>
                </table>
                
                <div style="margin-top:15px;display:flex;justify-content:space-between;align-items:center;">
                    <div class="form-group" style="flex:1;margin-right:20px;">
                        <label style="display:block;font-size:12px;color:#888;margin-bottom:5px;">收货备注</label>
                        <input type="text" name="notes" placeholder="收货备注..." style="width:100%;padding:8px 12px;border-radius:6px;border:1px solid rgba(255,255,255,0.1);background:#1a1a2e;color:#e0e0e0;font-size:13px;">
                    </div>
                    <div id="grand-total-<%= pid %>" style="font-size:18px;color:#fff;">
                        本次收货总额：<span style="color:#4CAF50;font-weight:700;" id="grand-value-<%= pid %>">¥0.00</span>
                    </div>
                </div>
                <button type="submit" class="btn btn--success" style="margin-top:10px;" onclick="return confirm('确认提交收货？')"><i class="fas fa-check-circle"></i> 提交收货单</button>
            </form>
            <% Else %>
            <p style="color:#666;text-align:center;padding:15px;">该订单已完成收货</p>
            <% End If %>
        </div>
        <%
                    rsDetails.Close
                    Set rsDetails = Nothing
                    rsReceiving.MoveNext
                Loop
            Else
        %>
        <div style="text-align:center;padding:60px;color:#666;">
            <i class="fas fa-inbox" style="font-size:48px;display:block;margin-bottom:15px;"></i>
            <p>暂无待收货的品牌定香采购订单</p>
            <a href="purchase_orders.asp" class="btn btn--primary"><i class="fas fa-file-invoice"></i> 查看采购订单</a>
        </div>
        <%      End If
            rsReceiving.Close
            Set rsReceiving = Nothing
        End If %>
    </div>
    
    <script>
        function updateRowTotal(input, idx, price, pid) {
            var val = parseFloat(input.value) || 0;
            var total = val * price;
            var rowEl = document.getElementById('row-total-' + pid + '-' + idx);
            if (rowEl) rowEl.textContent = '¥' + total.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            
            // 仅更新当前订单的总金额
            updateGrandTotal(pid);
        }
        
        function updateGrandTotal(pid) {
            var grandTotal = 0;
            // 仅选取当前订单的行（id以 row-total-{pid}- 开头）
            var totals = document.querySelectorAll('[id^="row-total-' + pid + '-"]');
            totals.forEach(function(el) {
                var val = parseFloat(el.textContent.replace('¥', '').replace(/,/g, '')) || 0;
                grandTotal += val;
            });
            var gv = document.getElementById('grand-value-' + pid);
            if (gv) gv.textContent = '¥' + grandTotal.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        }
        
        // 页面加载时为每个订单独立计算初始总额
        document.addEventListener('DOMContentLoaded', function() {
            var grandValues = document.querySelectorAll('[id^="grand-value-"]');
            grandValues.forEach(function(gv) {
                var pid = gv.id.replace('grand-value-', '');
                updateGrandTotal(pid);
            });
        });
    </script>
</body>
</html>
<%
If Not rsDetails Is Nothing Then
    rsDetails.Close
    Set rsDetails = Nothing
End If
Call CloseConnection()
%>
