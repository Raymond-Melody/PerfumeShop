<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Redirect "purchase_orders.asp?new=1&order_type=Bottle"
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

' 确保瓶子库存字段存在
On Error Resume Next
conn.Execute "SELECT StockQty FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD StockQty INT DEFAULT 0"
conn.Execute "SELECT SafetyStock FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD SafetyStock INT DEFAULT 0"
conn.Execute "SELECT UnitPrice FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD UnitPrice DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT SupplierID FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD SupplierID INT"
conn.Execute "SELECT Capacity FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD Capacity NVARCHAR(50)"
conn.Execute "SELECT Material FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD Material NVARCHAR(100)"
conn.Execute "SELECT CreatedAt FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD CreatedAt DATETIME DEFAULT GETDATE()"
' 检查BottleStyles表是否存在
conn.Execute "SELECT TOP 1 1 FROM BottleStyles"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE BottleStyles (BottleID INT IDENTITY(1,1) PRIMARY KEY, BottleName NVARCHAR(200), Capacity NVARCHAR(50), Material NVARCHAR(100), StockQty INT DEFAULT 0, SafetyStock INT DEFAULT 0, UnitPrice DECIMAL(19,4) DEFAULT 0, SupplierID INT, IsActive BIT DEFAULT 1, CreatedAt DATETIME DEFAULT GETDATE())"
End If
On Error GoTo 0

' ========== POST 处理 ==========
Dim msg, msgType
msg = Request.QueryString("msg")
msgType = "success"
If InStr(msg, "失败") > 0 Then msgType = "error"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action : action = Request.Form("action")
    
    If action = "restock" Then
        Dim rID : rID = SafeNum(Request.Form("bottle_id"))
        Dim rQty : rQty = SafeNum(Request.Form("restock_qty"))
        If rID > 0 And rQty > 0 Then
            conn.Execute "UPDATE BottleStyles SET StockQty = StockQty + " & rQty & " WHERE BottleID=" & rID
            Response.Redirect "bottle_purchase.asp?msg=入库成功"
            Response.End
        End If
    ElseIf action = "update" Then
        Dim eID : eID = SafeNum(Request.Form("bottle_id"))
        Dim eSafety : eSafety = SafeNum(Request.Form("safety_stock"))
        Dim ePrice : ePrice = SafeNum(Request.Form("unit_price"))
        Dim eSupplier : eSupplier = SafeNum(Request.Form("supplier_id"))
        If eID > 0 Then
            conn.Execute "UPDATE BottleStyles SET SafetyStock=" & eSafety & ", UnitPrice=" & ePrice & ", SupplierID=" & IIf(eSupplier>0, eSupplier, "NULL") & " WHERE BottleID=" & eID
            Response.Redirect "bottle_purchase.asp?msg=更新成功"
            Response.End
        End If
    ElseIf action = "add" Then
        Dim aName, aCapacity, aMaterial
        aName = SafeSQL(Trim(Request.Form("bottle_name")))
        aCapacity = SafeSQL(Trim(Request.Form("capacity")))
        aMaterial = SafeSQL(Trim(Request.Form("material")))
        aPrice = SafeNum(Request.Form("unit_price"))
        aSupplier = SafeNum(Request.Form("supplier_id"))
        If aName <> "" Then
            conn.Execute "INSERT INTO BottleStyles (BottleName, Capacity, Material, UnitPrice, SupplierID, IsActive, CreatedAt) VALUES ('" & _
                aName & "','" & aCapacity & "','" & aMaterial & "'," & aPrice & "," & IIf(aSupplier>0, aSupplier, "NULL") & ",1,GETDATE())"
            Response.Redirect "bottle_purchase.asp?msg=瓶子添加成功"
            Response.End
        End If
    End If
End If

' ========== 库存统计 ==========
Dim totalBottles, totalStock, lowStockCount, totalValue
totalBottles = GetScalar("SELECT COUNT(*) FROM BottleStyles")
totalStock = GetScalar("SELECT SUM(ISNULL(StockQty,0)) FROM BottleStyles")
lowStockCount = GetScalar("SELECT COUNT(*) FROM BottleStyles WHERE ISNULL(StockQty,0) <= ISNULL(SafetyStock,0) AND ISNULL(SafetyStock,0) > 0")
totalValue = SafeNum(GetScalar("SELECT SUM(ISNULL(StockQty,0) * ISNULL(UnitPrice,0)) FROM BottleStyles"))

' 关联采购订单
Dim relatedPOCount, relatedPOAmount
relatedPOCount = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE OrderType='Bottle' AND Status IN ('Ordered','PartialReceived','Received')")
relatedPOAmount = SafeNum(GetScalar("SELECT SUM(CAST(ISNULL(TotalAmount,0) AS FLOAT)) FROM PurchaseOrders WHERE OrderType='Bottle' AND Month(OrderDate)=Month(GETDATE()) AND Year(OrderDate)=Year(GETDATE())"))

' 获取瓶子列表
Dim rsBottles
Set rsBottles = conn.Execute("SELECT bs.*, s.SupplierName FROM BottleStyles bs LEFT JOIN Suppliers s ON bs.SupplierID = s.SupplierID ORDER BY bs.BottleName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>瓶子采购 - 采购管理中心</title>
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
        .breadcrumb a { color: #9C27B0; text-decoration: none; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 20px; border: 1px solid rgba(255,255,255,0.05); }
        .stat-card.warn { border: 1px solid rgba(255,152,0,0.5); }
        .stat-label { font-size: 12px; color: #888; margin-bottom: 6px; }
        .stat-value { font-size: 22px; font-weight: 700; color: #fff; }
        
        .toolbar { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; }

        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; }
        .data-table th { background: linear-gradient(135deg, #9C27B0, #7B1FA2); color: white; padding: 12px 15px; text-align: left; font-size: 13px; }
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
        .modal-content { background: linear-gradient(135deg, #2d2d44, #1e1e32); width: 90%; max-width: 500px; margin: 40px auto; padding: 30px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.1); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 15px; }
        .modal-header h3 { color: #fff; margin: 0; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; margin-bottom: 5px; color: #bbb; font-size: 13px; }
        .form-group input, .form-group select { width: 100%; padding: 10px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1e1e32; color: #e0e0e0; font-size: 14px; box-sizing: border-box; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        
        .bottle-preview { width: 40px; height: 40px; border-radius: 6px; display: inline-flex; align-items: center; justify-content: center; font-size: 18px; margin-right: 10px; vertical-align: middle; }
        .empty-row { text-align: center; padding: 40px; color: #666; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-wine-bottle" style="color:#9C27B0;"></i> 瓶子采购管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>瓶子采购</span>
            </div>
        </div>
        
        <% If msg <> "" Then %><div class="message" style="padding:12px 20px; border-radius:8px; margin-bottom:20px;"><%= msg %></div><% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-wine-bottle" style="color:#9C27B0;"></i> 瓶子款式</div>
                <div class="stat-value"><%= totalBottles %></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-layer-group" style="color:#9C27B0;"></i> 总库存量</div>
                <div class="stat-value"><%= totalStock %></div>
            </div>
            <div class="stat-card <% If lowStockCount > 0 Then Response.Write "warn" %>">
                <div class="stat-label"><i class="fas fa-exclamation-triangle" style="color:#FF9800;"></i> 低库存预警</div>
                <div class="stat-value" <% If lowStockCount > 0 Then %>style="color:#FF9800;"<% End If %>><%= lowStockCount %></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-yen-sign" style="color:#9C27B0;"></i> 库存总值</div>
                <div class="stat-value">¥<%= FormatNumber(totalValue, 0) %></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-file-invoice" style="color:#9C27B0;"></i> 进行中采购</div>
                <div class="stat-value"><%= relatedPOCount %> 单 / ¥<%= FormatNumber(relatedPOAmount, 0) %></div>
            </div>
        </div>
        
        <!-- 工具栏 -->
        <div class="toolbar">
            <button class="btn btn-success" onclick="openAddModal()"><i class="fas fa-plus"></i> 新增瓶子</button>
            <a href="purchase_orders.asp?order_type=Bottle" class="btn btn-outline"><i class="fas fa-file-invoice"></i> 瓶子采购订单</a>
            <a href="restock_order.asp?tab=Bottle" class="btn btn-outline"><i class="fas fa-bolt"></i> 智能补货</a>
            <a href="receiving.asp" class="btn btn-outline"><i class="fas fa-clipboard-check"></i> 收货入库</a>
        </div>
        
        <!-- 瓶子表格 -->
        <table class="data-table">
            <thead>
                <tr>
                    <th>瓶子款式</th>
                    <th>容量</th>
                    <th>材质</th>
                    <th>库存</th>
                    <th>单价</th>
                    <th>供应商</th>
                    <th style="text-align:center;">操作</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsBottles Is Nothing Or rsBottles.EOF Then
                %>
                <tr><td colspan="7" class="empty-row"><i class="fas fa-inbox"></i> 暂无瓶子数据</td></tr>
                <%
                Else
                    Do While Not rsBottles.EOF
                        Dim bID : bID = rsBottles(0)  ' 第一个字段即主键ID
                        Dim bStock : bStock = SafeNum(rsBottles("StockQty"))
                        Dim bSafety : bSafety = SafeNum(rsBottles("SafetyStock"))
                        Dim bBarClass, bStatusClass, bStatusText
                        If bSafety > 0 Then
                            If bStock <= 0 Then
                                bBarClass = "stock-critical" : bStatusClass = "status-crit" : bStatusText = "缺货"
                            ElseIf bStock <= bSafety Then
                                bBarClass = "stock-warning" : bStatusClass = "status-warn" : bStatusText = "低库存"
                            Else
                                bBarClass = "stock-normal" : bStatusClass = "status-ok" : bStatusText = "正常"
                            End If
                        Else
                            bBarClass = "stock-normal" : bStatusClass = "status-ok" : bStatusText = "正常"
                        End If
                        Dim bBarWidth : bBarWidth = 100
                        If bSafety > 0 Then bBarWidth = Int((bStock / (bSafety * 2)) * 100) : If bBarWidth > 100 Then bBarWidth = 100
                %>
                <tr>
                    <td>
                        <span class="bottle-preview" style="background:rgba(156,39,176,0.15);"><i class="fas fa-wine-bottle" style="color:#9C27B0;"></i></span>
                        <strong><%= Server.HTMLEncode(rsBottles("BottleName") & "") %></strong>
                    </td>
                    <td><%= Server.HTMLEncode(rsBottles("Capacity") & "") %>ml</td>
                    <td><%= Server.HTMLEncode(rsBottles("Material") & "") %></td>
                    <td>
                        <span class="stock-bar <%= bBarClass %>" style="width:<%= bBarWidth %>px;"></span>
                        <%= bStock %>
                        <span class="status-badge <%= bStatusClass %>"><%= bStatusText %></span>
                    </td>
                    <td>¥<%= FormatNumber(SafeNum(rsBottles("UnitPrice")), 2) %></td>
                    <td><%= Server.HTMLEncode(rsBottles("SupplierName") & "") %></td>
                    <td style="text-align:center;">
                        <button class="btn btn-outline btn--sm" onclick="openEditModal(<%= bID %>, '<%= Server.HTMLEncode(rsBottles("BottleName") & "") %>', <%= bSafety %>, <%= FormatNumber(SafeNum(rsBottles("UnitPrice")), 2) %>, '<%= Server.HTMLEncode(rsBottles("SupplierID") & "") %>')"><i class="fas fa-cog"></i></button>
                        <button class="btn btn-success btn--sm" onclick="openRestockModal(<%= bID %>, '<%= Server.HTMLEncode(rsBottles("BottleName") & "") %>')"><i class="fas fa-plus-circle"></i></button>
                    </td>
                </tr>
                <%
                        rsBottles.MoveNext
                    Loop
                    rsBottles.Close : Set rsBottles = Nothing
                End If
                %>
            </tbody>
        </table>
    </div>
    
    <!-- 编辑弹窗 -->
    <div class="modal" id="editModal">
        <div class="modal-content">
            <form method="post">
                <input type="hidden" name="action" value="update">
                <input type="hidden" name="bottle_id" id="editBottleID">
                <div class="form-group">
                    <label>瓶子名称</label>
                    <input type="text" id="editName" readonly style="background:#1a1a2e;">
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>安全库存</label>
                        <input type="number" name="safety_stock" id="editSafety" value="0">
                    </div>
                    <div class="form-group">
                        <label>单价 (¥)</label>
                        <input type="number" name="unit_price" id="editPrice" step="0.01" value="0">
                    </div>
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
            <form method="post">
                <input type="hidden" name="action" value="restock">
                <input type="hidden" name="bottle_id" id="restockID">
                <div class="form-group">
                    <label>瓶子名称</label>
                    <input type="text" id="restockName" readonly style="background:#1a1a2e;">
                </div>
                <div class="form-group">
                    <label>入库数量</label>
                    <input type="number" name="restock_qty" required min="1" value="1">
                </div>
                <div style="text-align:right; margin-top:15px;">
                    <button type="button" class="btn btn-outline" onclick="closeModal('restockModal')">取消</button>
                    <button type="submit" class="btn btn-success">确认入库</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 新增瓶子弹窗 -->
    <div class="modal" id="addModal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-plus-circle"></i> 新增瓶子</h3>
                <button class="modal-close" onclick="closeModal('addModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="add">
                <div class="form-group">
                    <label>瓶子名称 <span class="required" style="color:#F44336;">*</span></label>
                    <input type="text" name="bottle_name" required placeholder="请输入瓶子名称">
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>容量 (ml)</label>
                        <input type="text" name="capacity" placeholder="如：50">
                    </div>
                    <div class="form-group">
                        <label>材质</label>
                        <input type="text" name="material" placeholder="如：玻璃">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>单价 (¥)</label>
                        <input type="number" name="unit_price" step="0.01" value="0">
                    </div>
                    <div class="form-group">
                        <label>供应商</label>
                        <select name="supplier_id">
                            <option value="0">无</option>
                            <%
                            Dim rsAddSup : Set rsAddSup = conn.Execute("SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive=1 ORDER BY SupplierName")
                            If Not rsAddSup Is Nothing Then
                                Do While Not rsAddSup.EOF
                            %>
                            <option value="<%= rsAddSup("SupplierID") %>"><%= Server.HTMLEncode(rsAddSup("SupplierName") & "") %></option>
                            <%
                                    rsAddSup.MoveNext
                                Loop
                                rsAddSup.Close : Set rsAddSup = Nothing
                            End If
                            %>
                        </select>
                    </div>
                </div>
                <div style="text-align:right; margin-top:20px;">
                    <button type="button" class="btn btn-outline" onclick="closeModal('addModal')">取消</button>
                    <button type="submit" class="btn btn-success"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function openAddModal() {
            document.getElementById('addModal').style.display = 'block';
        }
        
        function openEditModal(id, name, safety, price, supplier) {
            document.getElementById('editBottleID').value = id;
            document.getElementById('editName').value = name;
            document.getElementById('editSafety').value = safety;
            document.getElementById('editPrice').value = price;
            document.getElementById('editSupplier').value = supplier;
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
