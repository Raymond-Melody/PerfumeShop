<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Redirect "purchase_orders.asp?new=1&order_type=Printing"
%>

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
    Set rs = Nothing : GetScalar = val
End Function

' 确保 PrintingInventory 表存在
On Error Resume Next
conn.Execute "SELECT TOP 1 1 FROM PrintingInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PrintingInventory (PrintingID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(200) NOT NULL, ItemCode NVARCHAR(50), PrintingType NVARCHAR(30) DEFAULT 'Manual', Specification NVARCHAR(200), Unit NVARCHAR(20) DEFAULT 'pcs', StockQty FLOAT DEFAULT 0, SafetyStock FLOAT DEFAULT 0, UnitPrice DECIMAL(19,4) DEFAULT 0, WeightedUnitCost DECIMAL(19,4) DEFAULT 0, SupplierID INT, LastPurchaseDate DATETIME, Notes NVARCHAR(500), IsActive BIT DEFAULT 1, CreatedAt DATETIME DEFAULT GETDATE(), UpdatedAt DATETIME)"
End If

' 确保字段存在（兼容旧版不同结构）
conn.Execute "SELECT SupplierID FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD SupplierID INT"
conn.Execute "SELECT IsActive FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD IsActive BIT DEFAULT 1"
conn.Execute "SELECT ItemCode FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD ItemCode NVARCHAR(50)"
conn.Execute "SELECT PrintingType FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD PrintingType NVARCHAR(30) DEFAULT 'Manual'"
conn.Execute "SELECT Specification FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD Specification NVARCHAR(200)"
conn.Execute "SELECT Unit FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD Unit NVARCHAR(20) DEFAULT 'pcs'"
conn.Execute "SELECT UnitPrice FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD UnitPrice DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT WeightedUnitCost FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD WeightedUnitCost DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT Notes FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD Notes NVARCHAR(500)"
conn.Execute "SELECT LastPurchaseDate FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD LastPurchaseDate DATETIME"
conn.Execute "SELECT CreatedAt FROM PrintingInventory WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PrintingInventory ADD CreatedAt DATETIME DEFAULT GETDATE()"
On Error GoTo 0

' ========== POST 处理 ==========
Dim msg, msgType
msg = Request.QueryString("msg")
msgType = "success"
If InStr(msg, "失败") > 0 Then msgType = "error"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action : action = Request.Form("action")
    
    If action = "add" Then
        Dim pName, pCode, pType, pSpec, pUnit, pStock, pSafety, pPrice, pSupplier, pNotes
        pName = SafeSQL(Trim(Request.Form("item_name")))
        pCode = SafeSQL(Trim(Request.Form("item_code")))
        pType = SafeSQL(Trim(Request.Form("printing_type")))
        pSpec = SafeSQL(Trim(Request.Form("specification")))
        pUnit = SafeSQL(Trim(Request.Form("unit")))
        pStock = SafeNum(Request.Form("stock_qty"))
        pSafety = SafeNum(Request.Form("safety_stock"))
        pPrice = SafeNum(Request.Form("unit_price"))
        pSupplier = SafeNum(Request.Form("supplier_id"))
        pNotes = SafeSQL(Trim(Request.Form("notes")))
        
        If pName <> "" Then
            conn.Execute "INSERT INTO PrintingInventory (ItemName, ItemCode, PrintingType, Specification, Unit, StockQty, SafetyStock, UnitPrice, WeightedUnitCost, SupplierID, Notes) VALUES ('" & _
                pName & "','" & pCode & "','" & pType & "','" & pSpec & "','" & pUnit & "'," & pStock & "," & pSafety & "," & pPrice & "," & pPrice & "," & IIf(pSupplier>0, pSupplier, "NULL") & ",'" & pNotes & "')"
            Response.Redirect "printing_purchase.asp?msg=添加成功"
            Response.End
        End If
    ElseIf action = "edit" Then
        Dim eID : eID = SafeNum(Request.Form("printing_id"))
        pName = SafeSQL(Trim(Request.Form("item_name")))
        pCode = SafeSQL(Trim(Request.Form("item_code")))
        pType = SafeSQL(Trim(Request.Form("printing_type")))
        pSpec = SafeSQL(Trim(Request.Form("specification")))
        pUnit = SafeSQL(Trim(Request.Form("unit")))
        pSafety = SafeNum(Request.Form("safety_stock"))
        pPrice = SafeNum(Request.Form("unit_price"))
        pSupplier = SafeNum(Request.Form("supplier_id"))
        pNotes = SafeSQL(Trim(Request.Form("notes")))
        
        If eID > 0 And pName <> "" Then
            conn.Execute "UPDATE PrintingInventory SET ItemName='" & pName & "', ItemCode='" & pCode & "', PrintingType='" & pType & "', Specification='" & pSpec & "', Unit='" & pUnit & "', SafetyStock=" & pSafety & ", UnitPrice=" & pPrice & ", SupplierID=" & IIf(pSupplier>0, pSupplier, "NULL") & ", Notes='" & pNotes & "', UpdatedAt=GETDATE() WHERE PrintingID=" & eID
            Response.Redirect "printing_purchase.asp?msg=更新成功"
            Response.End
        End If
    ElseIf action = "restock" Then
        Dim rID : rID = SafeNum(Request.Form("printing_id"))
        Dim rQty : rQty = SafeNum(Request.Form("restock_qty"))
        If rID > 0 And rQty > 0 Then
            conn.Execute "UPDATE PrintingInventory SET StockQty = StockQty + " & rQty & ", UpdatedAt=GETDATE() WHERE PrintingID=" & rID
            Response.Redirect "printing_purchase.asp?msg=入库成功"
            Response.End
        End If
    End If
End If

' ========== 库存统计 ==========
Dim totalItems, totalStock, lowStockCount, totalValue
totalItems = GetScalar("SELECT COUNT(*) FROM PrintingInventory WHERE IsActive=1")
totalStock = GetScalar("SELECT SUM(StockQty) FROM PrintingInventory WHERE IsActive=1")
lowStockCount = GetScalar("SELECT COUNT(*) FROM PrintingInventory WHERE IsActive=1 AND StockQty <= SafetyStock AND SafetyStock > 0")
totalValue = SafeNum(GetScalar("SELECT SUM(StockQty * UnitPrice) FROM PrintingInventory WHERE IsActive=1"))

' ========== 关联采购订单 ==========
Dim relatedPOCount, relatedPOAmount
relatedPOCount = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE OrderType='Printing' AND Status IN ('Ordered','PartialReceived','Received')")
relatedPOAmount = SafeNum(GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE OrderType='Printing' AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())"))

' ========== 获取印刷品列表 ==========
Dim rsInv
Set rsInv = conn.Execute("SELECT p.*, s.SupplierName FROM PrintingInventory p LEFT JOIN Suppliers s ON p.SupplierID = s.SupplierID WHERE p.IsActive=1 ORDER BY p.ItemName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>印刷品采购 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 260px; padding: 30px; }
        .page-header { margin-bottom: 25px; }
        .page-title { font-size: 24px; color: #fff; display: flex; align-items: center; gap: 10px; }
        .breadcrumb { color: #888; font-size: 13px; margin-top: 5px; }
        .breadcrumb a { color: #FF9800; text-decoration: none; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 20px; border: 1px solid rgba(255,255,255,0.05); }
        .stat-card.warn { border: 1px solid rgba(255,152,0,0.5); }
        .stat-label { font-size: 12px; color: #888; margin-bottom: 6px; }
        .stat-value { font-size: 22px; font-weight: 700; color: #fff; }
        
        .toolbar { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; }

        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; }
        .data-table th { background: linear-gradient(135deg, #FF9800, #F57C00); color: white; padding: 12px 15px; text-align: left; font-size: 13px; }
        .data-table td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.05); color: #e0e0e0; font-size: 13px; }
        .data-table tr:hover td { background: rgba(255,255,255,0.02); }
        
        .stock-bar { display: inline-block; height: 6px; border-radius: 3px; vertical-align: middle; margin-right: 8px; }
        .stock-normal { background: #4CAF50; }
        .stock-warning { background: #FF9800; }
        .stock-critical { background: #F44336; }
        
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .status-ok { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-warn { background: rgba(255,152,0,0.2); color: #FF9800; }
        .status-crit { background: rgba(244,67,54,0.2); color: #F44336; }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; backdrop-filter: blur(5px); }
        .modal-content { background: linear-gradient(135deg, #2d2d44, #1e1e32); width: 90%; max-width: 600px; margin: 40px auto; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.1); max-height: 80vh; overflow-y: auto; }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 15px; }
        .modal-header h3 { color: #fff; margin: 0; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; margin-bottom: 5px; color: #bbb; font-size: 13px; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1e1e32; color: #e0e0e0; font-size: 14px; box-sizing: border-box; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        
        .empty-row { text-align: center; padding: 40px; color: #666; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-print" style="color:#00BCD4;"></i> 印刷品采购管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>印刷品采购</span>
            </div>
        </div>
        
        <% If msg <> "" Then %><div class="message" style="padding:12px 20px; border-radius:8px; margin-bottom:20px; background:<%=IIf(msgType="error","rgba(244,67,54,0.15)","rgba(76,175,80,0.15)")%>; color:<%=IIf(msgType="error","#F44336","#4CAF50")%>;"><%= msg %></div><% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-print" style="color:#00BCD4;"></i> 印刷物品类</div>
                <div class="stat-value"><%= totalItems %></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-layer-group" style="color:#00BCD4;"></i> 总库存量</div>
                <div class="stat-value"><%= totalStock %></div>
            </div>
            <div class="stat-card<%= IIf(lowStockCount > 0, " warn", "") %>">
                <div class="stat-label"><i class="fas fa-exclamation-triangle" style="color:#FF9800;"></i> 低库存预警</div>
                <div class="stat-value" <% If lowStockCount > 0 Then %>style="color:#FF9800;"<% End If %>><%= lowStockCount %></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-yen-sign" style="color:#00BCD4;"></i> 库存总值</div>
                <div class="stat-value">¥<%= FormatNumber(totalValue, 0) %></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-file-invoice" style="color:#00BCD4;"></i> 进行中采购</div>
                <div class="stat-value"><%= relatedPOCount %> 单 / ¥<%= FormatNumber(relatedPOAmount, 0) %></div>
            </div>
        </div>
        
        <!-- 工具栏 -->
        <div class="toolbar">
            <button class="btn btn-primary" onclick="openAddModal()"><i class="fas fa-plus"></i> 新增印刷品</button>
            <a href="purchase_orders.asp?order_type=Printing" class="btn btn-outline"><i class="fas fa-file-invoice"></i> 印刷品采购订单</a>
            <a href="restock_order.asp?tab=Printing" class="btn btn-outline"><i class="fas fa-bolt"></i> 智能补货</a>
            <a href="receiving.asp" class="btn btn-outline"><i class="fas fa-clipboard-check"></i> 收货入库</a>
        </div>
        
        <!-- 库存表格 -->
        <table class="data-table">
            <thead>
                <tr>
                    <th>印刷品名称</th>
                    <th>印刷类型</th>
                    <th>规格</th>
                    <th>库存</th>
                    <th>单价</th>
                    <th>供应商</th>
                    <th style="text-align:center;">操作</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsInv Is Nothing Or rsInv.EOF Then
                %>
                <tr><td colspan="7" class="empty-row"><i class="fas fa-inbox"></i> 暂无印刷品库存数据</td></tr>
                <%
                Else
                    Do While Not rsInv.EOF
                        Dim pID : pID = rsInv("PrintingID")
                        Dim pStockQty : pStockQty = SafeNum(rsInv("StockQty"))
                        pSafety = SafeNum(rsInv("SafetyStock"))
                        Dim pBarClass, pStatusClass, pStatusText
                        If pSafety > 0 Then
                            If pStockQty <= 0 Then
                                pBarClass = "stock-critical" : pStatusClass = "status-crit" : pStatusText = "缺货"
                            ElseIf pStockQty <= pSafety Then
                                pBarClass = "stock-warning" : pStatusClass = "status-warn" : pStatusText = "低库存"
                            Else
                                pBarClass = "stock-normal" : pStatusClass = "status-ok" : pStatusText = "正常"
                            End If
                        Else
                            pBarClass = "stock-normal" : pStatusClass = "status-ok" : pStatusText = "正常"
                        End If
                        Dim pBarWidth : pBarWidth = 100
                        If pSafety > 0 Then pBarWidth = Int((pStockQty / (pSafety * 2)) * 100) : If pBarWidth > 100 Then pBarWidth = 100
                %>
                <tr>
                    <td><strong><%= Server.HTMLEncode(rsInv("ItemName") & "") %></strong></td>
                    <td><%= Server.HTMLEncode(rsInv("PrintingType") & "") %></td>
                    <td><%= Server.HTMLEncode(rsInv("Specification") & "") %></td>
                    <td>
                        <span class="stock-bar <%= pBarClass %>" style="width:<%= pBarWidth %>px;"></span>
                        <%= pStockQty %> <%= Server.HTMLEncode(rsInv("Unit") & "") %>
                        <span class="status-badge <%= pStatusClass %>"><%= pStatusText %></span>
                    </td>
                    <td>¥<%= FormatNumber(SafeNum(rsInv("UnitPrice")), 2) %></td>
                    <td><%= Server.HTMLEncode(rsInv("SupplierName") & "") %></td>
                    <td style="text-align:center;">
                        <button class="btn btn-outline btn--sm" onclick="openEditModal(<%= pID %>, '<%= Server.HTMLEncode(rsInv("ItemName") & "") %>', '<%= Server.HTMLEncode(rsInv("ItemCode") & "") %>', '<%= Server.HTMLEncode(rsInv("PrintingType") & "") %>', '<%= Server.HTMLEncode(rsInv("Specification") & "") %>', '<%= Server.HTMLEncode(rsInv("Unit") & "") %>', <%= pSafety %>, <%= FormatNumber(SafeNum(rsInv("UnitPrice")), 2) %>, '<%= Server.HTMLEncode(rsInv("SupplierID") & "") %>', '<%= Server.HTMLEncode(rsInv("Notes") & "") %>')"><i class="fas fa-edit"></i></button>
                        <button class="btn btn-success btn--sm" onclick="openRestockModal(<%= pID %>, '<%= Server.HTMLEncode(rsInv("ItemName") & "") %>')"><i class="fas fa-plus-circle"></i></button>
                    </td>
                </tr>
                <%
                        rsInv.MoveNext
                    Loop
                    rsInv.Close : Set rsInv = Nothing
                End If
                %>
            </tbody>
        </table>
    </div>
    
    <!-- 新增/编辑弹窗 -->
    <div class="modal" id="editModal">
        <div class="modal-content">
            <div class="modal-header">
                <h3 id="modalTitle"><i class="fas fa-plus-circle"></i> 新增印刷品</h3>
                <button class="modal-close" onclick="closeModal('editModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" id="formAction" value="add">
                <input type="hidden" name="printing_id" id="editPrintingID">
                <div class="form-row">
                    <div class="form-group">
                        <label>印刷品名称 <span style="color:#F44336;">*</span></label>
                        <input type="text" name="item_name" id="editName" required>
                    </div>
                    <div class="form-group">
                        <label>编码</label>
                        <input type="text" name="item_code" id="editCode">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>印刷类型</label>
                        <select name="printing_type" id="editType">
                            <option value="Manual">产品说明书</option>
                            <option value="Label">标签</option>
                            <option value="Card">卡片</option>
                            <option value="Booklet">手册</option>
                            <option value="Insert">插页</option>
                            <option value="Other">其他</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>规格尺寸</label>
                        <input type="text" name="specification" id="editSpec">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>单位</label>
                        <input type="text" name="unit" id="editUnit" value="pcs">
                    </div>
                    <div class="form-group">
                        <label>供应商</label>
                        <select name="supplier_id" id="editSupplier">
                            <option value="0">无</option>
                            <%
                            Dim rsSup : Set rsSup = conn.Execute("SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive=1 ORDER BY SupplierName")
                            If Not rsSup Is Nothing Then
                                Do While Not rsSup.EOF
                            %>
                            <option value="<%= rsSup("SupplierID") %>"><%= Server.HTMLEncode(rsSup("SupplierName") & "") %></option>
                            <%
                                    rsSup.MoveNext
                                Loop
                                rsSup.Close : Set rsSup = Nothing
                            End If
                            %>
                        </select>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>单价 (¥)</label>
                        <input type="number" name="unit_price" id="editPrice" step="0.01" value="0">
                    </div>
                    <div class="form-group">
                        <label>安全库存</label>
                        <input type="number" name="safety_stock" id="editSafety" value="0">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>初始库存</label>
                        <input type="number" name="stock_qty" id="editStock" value="0">
                    </div>
                </div>
                <div class="form-group">
                    <label>备注</label>
                    <textarea name="notes" id="editNotes" rows="2"></textarea>
                </div>
                <div style="text-align:right; margin-top:15px;">
                    <button type="button" class="btn btn-outline" onclick="closeModal('editModal')">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 入库弹窗 -->
    <div class="modal" id="restockModal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-plus-circle"></i> 库存入库</h3>
                <button class="modal-close" onclick="closeModal('restockModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="restock">
                <input type="hidden" name="printing_id" id="restockID">
                <div class="form-group">
                    <label>印刷品名称</label>
                    <input type="text" id="restockName" readonly style="background:#1a1a2e;">
                </div>
                <div class="form-group">
                    <label>入库数量 <span style="color:#F44336;">*</span></label>
                    <input type="number" name="restock_qty" required min="1" value="1">
                </div>
                <div style="text-align:right; margin-top:15px;">
                    <button type="button" class="btn btn-outline" onclick="closeModal('restockModal')">取消</button>
                    <button type="submit" class="btn btn-success">确认入库</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function openAddModal() {
            document.getElementById('formAction').value = 'add';
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-plus-circle"></i> 新增印刷品';
            document.getElementById('editPrintingID').value = '';
            document.getElementById('editName').value = '';
            document.getElementById('editCode').value = '';
            document.getElementById('editType').value = 'Manual';
            document.getElementById('editSpec').value = '';
            document.getElementById('editUnit').value = 'pcs';
            document.getElementById('editPrice').value = '0';
            document.getElementById('editSafety').value = '0';
            document.getElementById('editStock').value = '0';
            document.getElementById('editSupplier').value = '0';
            document.getElementById('editNotes').value = '';
            document.getElementById('editStock').parentNode.parentNode.style.display = 'grid';
            document.getElementById('editModal').style.display = 'block';
        }
        
        function openEditModal(id, name, code, type, spec, unit, safety, price, supplier, notes) {
            document.getElementById('formAction').value = 'edit';
            document.getElementById('modalTitle').innerHTML = '<i class="fas fa-edit"></i> 编辑印刷品';
            document.getElementById('editPrintingID').value = id;
            document.getElementById('editName').value = name;
            document.getElementById('editCode').value = code;
            document.getElementById('editType').value = type;
            document.getElementById('editSpec').value = spec;
            document.getElementById('editUnit').value = unit;
            document.getElementById('editPrice').value = price;
            document.getElementById('editSafety').value = safety;
            document.getElementById('editSupplier').value = supplier;
            document.getElementById('editNotes').value = notes;
            document.getElementById('editStock').parentNode.parentNode.style.display = 'none';
            document.getElementById('editModal').style.display = 'block';
        }
        
        function openRestockModal(id, name) {
            document.getElementById('restockID').value = id;
            document.getElementById('restockName').value = name;
            document.getElementById('restockModal').style.display = 'block';
        }
        
        function closeModal(id) { document.getElementById(id).style.display = 'none'; }
        window.onclick = function(e) { if (e.target.classList.contains('modal')) e.target.style.display = 'none'; }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
