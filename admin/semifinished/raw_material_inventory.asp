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

' 确保RawMaterialInventory表有WeightedUnitCost和SupplierID字段
On Error Resume Next
Dim rsRMICheck
Set rsRMICheck = conn.Execute("SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='RawMaterialInventory' AND COLUMN_NAME='WeightedUnitCost'")
If Not rsRMICheck Is Nothing Then
    If Not rsRMICheck.EOF Then
        If CLng(rsRMICheck(0)) = 0 Then
            conn.Execute "ALTER TABLE RawMaterialInventory ADD WeightedUnitCost DECIMAL(18,6) DEFAULT 0"
            Err.Clear
        End If
    End If
    rsRMICheck.Close
End If
Set rsRMICheck = Nothing

Set rsRMICheck = conn.Execute("SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='RawMaterialInventory' AND COLUMN_NAME='SupplierID'")
If Not rsRMICheck Is Nothing Then
    If Not rsRMICheck.EOF Then
        If CLng(rsRMICheck(0)) = 0 Then
            conn.Execute "ALTER TABLE RawMaterialInventory ADD SupplierID INT NULL"
            Err.Clear
        End If
    End If
    rsRMICheck.Close
End If
Set rsRMICheck = Nothing
Err.Clear
On Error GoTo 0

' 检查WeightedUnitCost列是否实际存在（不依赖ALTER TABLE成功）
Dim hasWCCol : hasWCCol = False
Dim hasSupCol : hasSupCol = False
On Error Resume Next
Dim rsTmpChk
Set rsTmpChk = conn.Execute("SELECT COL_LENGTH('RawMaterialInventory','WeightedUnitCost')")
If Err.Number = 0 And Not rsTmpChk Is Nothing Then
    If Not rsTmpChk.EOF Then
        If Not IsNull(rsTmpChk(0)) Then hasWCCol = True
    End If
    rsTmpChk.Close
End If
Set rsTmpChk = Nothing
Err.Clear
Set rsTmpChk = conn.Execute("SELECT COL_LENGTH('RawMaterialInventory','SupplierID')")
If Err.Number = 0 And Not rsTmpChk Is Nothing Then
    If Not rsTmpChk.EOF Then
        If Not IsNull(rsTmpChk(0)) Then hasSupCol = True
    End If
    rsTmpChk.Close
End If
Set rsTmpChk = Nothing
Err.Clear
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
    If action = "add" Then
        Dim mName, mCode, mCat, mQty, mSafety, mUnit, mPrice, mSupplier
        mName = Trim(Request.Form("item_name"))
        mCode = Trim(Request.Form("item_code"))
        mCat = Trim(Request.Form("category_code"))
        mQty = SafeNum(Request.Form("stock_qty"))
        mSafety = SafeNum(Request.Form("safety_stock"))
        mUnit = Trim(Request.Form("unit"))
        mPrice = SafeNum(Request.Form("unit_price"))
        mSupplier = SafeNum(Request.Form("supplier_id"))
        
        If mName <> "" Then
            conn.Execute "INSERT INTO RawMaterialInventory (ItemName, ItemCode, CategoryCode, StockQty, SafetyStock, Unit, UnitPrice, SupplierID, UpdatedAt) VALUES ('" & _
                SafeSQL(mName) & "','" & SafeSQL(mCode) & "','" & SafeSQL(mCat) & "'," & mQty & "," & mSafety & ",'" & SafeSQL(mUnit) & "'," & mPrice & "," & mSupplier & ",GETDATE())"
            Response.Redirect "raw_material_inventory.asp?msg=原料添加成功"
            Response.End
        Else
            msg = "原料名称不能为空"
            msgType = "error"
        End If
    
    ElseIf action = "update" Then
        Dim uID, uQty, uSafety, uPrice
        uID = SafeNum(Request.Form("material_id"))
        uQty = SafeNum(Request.Form("stock_qty"))
        uSafety = SafeNum(Request.Form("safety_stock"))
        uPrice = SafeNum(Request.Form("unit_price"))
        
        If uID > 0 Then
            conn.Execute "UPDATE RawMaterialInventory SET StockQty=" & uQty & ", SafetyStock=" & uSafety & ", UnitPrice=" & uPrice & ", UpdatedAt=GETDATE() WHERE MaterialID=" & uID
            Response.Redirect "raw_material_inventory.asp?msg=原料更新成功"
            Response.End
        End If
    
    ElseIf action = "restock" Then
        Dim rID, rAddQty
        rID = SafeNum(Request.Form("material_id"))
        rAddQty = SafeNum(Request.Form("add_qty"))
        
        If rID > 0 And rAddQty > 0 Then
            conn.Execute "UPDATE RawMaterialInventory SET StockQty=StockQty+" & rAddQty & ", UpdatedAt=GETDATE(), LastPurchaseDate=GETDATE() WHERE MaterialID=" & rID
            conn.Execute "INSERT INTO InventoryTransactions (MaterialID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES (" & _
                rID & "," & rAddQty & ",'入库','IN','手动入库','" & SafeSQL(Session("AdminUsername")) & "',GETDATE())"
            Response.Redirect "raw_material_inventory.asp?msg=入库成功"
            Response.End
        End If
    End If
End If

' ========== 统计 ==========
Dim rmTotal, rmLowStock, rmZeroStock, rmTotalValue
rmTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory"))
rmLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
rmZeroStock = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= 0"))
rmTotalValue = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty * " & IIf(hasWCCol, "ISNULL(WeightedUnitCost, UnitPrice)", "UnitPrice") & "),0) FROM RawMaterialInventory"))

' ========== 原料列表 ==========
Dim rsMaterials, searchKey
searchKey = Trim(Request.QueryString("search"))
Dim matSQL
Dim wcSelectExpr : wcSelectExpr = IIf(hasWCCol, "rmi.WeightedUnitCost", "rmi.UnitPrice AS WeightedUnitCost")
Dim supJoinExpr : supJoinExpr = IIf(hasSupCol, " LEFT JOIN Suppliers s ON rmi.SupplierID=s.SupplierID", "")
Dim supSelectExpr : supSelectExpr = IIf(hasSupCol, ", s.SupplierName", ", '' AS SupplierName")
matSQL = "SELECT rmi.MaterialID, rmi.ItemName, rmi.ItemCode, rmi.CategoryCode, rmi.StockQty, rmi.SafetyStock, rmi.Unit, rmi.UnitPrice, " & wcSelectExpr & supSelectExpr & " FROM RawMaterialInventory rmi" & supJoinExpr
If searchKey <> "" Then
    matSQL = matSQL & " WHERE rmi.ItemName LIKE '%" & SafeSQL(searchKey) & "%' OR rmi.ItemCode LIKE '%" & SafeSQL(searchKey) & "%'"
End If
matSQL = matSQL & " ORDER BY rmi.ItemName ASC"
Set rsMaterials = conn.Execute(matSQL)
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>原料库存 - 半成品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #2196F3; --danger: #f44336; --warning: #FF9800; --success: #4CAF50; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #4CAF50; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(76,175,80,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(76,175,80,0.15); color: #81c784; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        tr.low-stock td { background: rgba(244,67,54,0.06); }
        
        .stock-bar { height: 6px; background: rgba(255,255,255,0.1); border-radius: 3px; margin-top: 4px; }
        .stock-bar-fill { height: 100%; border-radius: 3px; }
        .stock-bar-fill.safe { background: #4CAF50; }
        .stock-bar-fill.warning { background: #FF9800; }
        .stock-bar-fill.danger { background: #f44336; }
        
        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #e57373; border: 1px solid rgba(244,67,54,0.3); }
        
        .search-box { display: flex; gap: 10px; align-items: center; }
        .search-box input { padding: 8px 14px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; width: 250px; }
        .search-box input:focus { outline: none; border-color: #2196F3; }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); width: 90%; max-width: 500px; margin: 80px auto; padding: 30px; border-radius: 15px; border: 1px solid rgba(255,255,255,0.06); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .modal-header h3 { margin: 0; font-size: 18px; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 25px; }
        
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; margin-bottom: 6px; font-weight: 600; color: #e0e0e0; font-size: 13px; }
        .form-group input, .form-group select { width: 100%; padding: 10px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        .form-group input:focus, .form-group select:focus { outline: none; border-color: #2196F3; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .text-right { text-align: right; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-boxes"></i> 原料库存管理</h2>
            <div class="search-box">
                <form method="get" style="display:flex;gap:8px;">
                    <input type="text" name="search" placeholder="搜索原料名称/编码..." value="<%=Server.HTMLEncode(searchKey)%>">
                    <button type="submit" class="btn btn-primary btn-sm"><i class="fas fa-search"></i></button>
                    <% If searchKey <> "" Then %>
                    <a href="raw_material_inventory.asp" class="btn btn-sm" style="background:#555;color:#fff;">清除</a>
                    <% End If %>
                </form>
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=rmTotal%></span><span class="label">原料种类</span></div>
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=rmLowStock%></span><span class="label">低库存预警</span></div>
            <div class="stat-card"><span class="num" style="color:#f44336;"><%=rmZeroStock%></span><span class="label">零库存</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;">¥<%=FormatNumber(rmTotalValue,2)%></span><span class="label">库存总值</span></div>
        </div>
        
        <!-- 原料列表 -->
        <div class="card">
            <div class="card-header">
                原料清单
                <button class="btn btn-success btn-sm" onclick="openAddModal()"><i class="fas fa-plus"></i> 添加原料</button>
            </div>
            <div class="card-body">
                <table>
                    <thead><tr><th>编码</th><th>名称</th><th>类别</th><th>库存量</th><th>安全库存</th><th>单位</th><th>参考单价</th><th>加权成本</th><th>供应商</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    If Not rsMaterials Is Nothing Then
                        Dim rowCount : rowCount = 0
                        Do While Not rsMaterials.EOF
                            rowCount = rowCount + 1
                            Dim mIDRow, mQtyRow, mSafetyRow
                            mIDRow = rsMaterials("MaterialID")
                            mQtyRow = SafeNum(rsMaterials("StockQty"))
                            mSafetyRow = SafeNum(rsMaterials("SafetyStock"))
                            Dim stockClass, stockPct
                            If mSafetyRow > 0 Then
                                stockPct = (mQtyRow / mSafetyRow) * 100
                                If mQtyRow <= 0 Then
                                    stockClass = "danger"
                                ElseIf mQtyRow <= mSafetyRow Then
                                    stockClass = "warning"
                                Else
                                    stockClass = "safe"
                                End If
                            Else
                                stockPct = 100 : stockClass = "safe"
                            End If
                            If stockPct > 100 Then stockPct = 100
                    %>
                        <tr class="<%=IIF(mSafetyRow>0 And mQtyRow<=mSafetyRow,"low-stock","")%>">
                            <td style="color:#888;font-size:13px;"><%=rsMaterials("ItemCode") & ""%></td>
                            <td><strong><%=Server.HTMLEncode(rsMaterials("ItemName") & "")%></strong></td>
                            <td><%=rsMaterials("CategoryCode") & ""%></td>
                            <td>
                                <%=FormatNumber(mQtyRow,1)%>
                                <div class="stock-bar"><div class="stock-bar-fill <%=stockClass%>" style="width:<%=stockPct%>%;"></div></div>
                            </td>
                            <td><%=FormatNumber(mSafetyRow,1)%></td>
                            <td><%=rsMaterials("Unit") & ""%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsMaterials("UnitPrice")),2)%></td>
                            <td>¥<%=FormatNumber(SafeNum(rsMaterials("WeightedUnitCost")),2)%></td>
                            <td><%=rsMaterials("SupplierName") & ""%></td>
                            <td>
                                <button class="btn btn-primary btn-sm" onclick="openEditModal(<%=mIDRow%>,<%=mQtyRow%>,<%=mSafetyRow%>,<%=SafeNum(rsMaterials("UnitPrice"))%>)">编辑</button>
                                <button class="btn btn-success btn-sm" onclick="openRestockModal(<%=mIDRow%>,'<%=Replace(Replace(rsMaterials("ItemName") & "", "'", "\'"), Chr(10), "")%>')">入库</button>
                            </td>
                        </tr>
                    <%
                            rsMaterials.MoveNext
                        Loop
                        rsMaterials.Close
                    End If
                    Set rsMaterials = Nothing
                    If rowCount = 0 Then
                    %>
                        <tr><td colspan="10" class="text-center text-muted" style="padding:40px;">暂无原料数据</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 添加原料弹窗 -->
    <div id="addModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>添加原料</h3>
                <button class="modal-close" onclick="closeModal('addModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="add">
                <div class="form-group"><label>原料名称 *</label><input type="text" name="item_name" required></div>
                <div class="form-group"><label>编码</label><input type="text" name="item_code"></div>
                <div class="form-group"><label>类别</label><input type="text" name="category_code" placeholder="如：香料、溶剂、辅料"></div>
                <div class="form-group"><label>初始库存</label><input type="number" name="stock_qty" value="0" step="0.1"></div>
                <div class="form-group"><label>安全库存</label><input type="number" name="safety_stock" value="0" step="0.1"></div>
                <div class="form-group"><label>单位</label><input type="text" name="unit" value="g"></div>
                <div class="form-group"><label>单价</label><input type="number" name="unit_price" value="0" step="0.01"></div>
                <div class="form-group"><label>供应商ID</label><input type="number" name="supplier_id" value="0"></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('addModal')">取消</button>
                    <button type="submit" class="btn btn-success">确认添加</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 编辑原料弹窗 -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>编辑原料库存</h3>
                <button class="modal-close" onclick="closeModal('editModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="update">
                <input type="hidden" name="material_id" id="editID">
                <div class="form-group"><label>当前库存</label><input type="number" name="stock_qty" id="editQty" step="0.1"></div>
                <div class="form-group"><label>安全库存</label><input type="number" name="safety_stock" id="editSafety" step="0.1"></div>
                <div class="form-group"><label>单价</label><input type="number" name="unit_price" id="editPrice" step="0.01"></div>
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
                <h3>原料入库 - <span id="restockName"></span></h3>
                <button class="modal-close" onclick="closeModal('restockModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="restock">
                <input type="hidden" name="material_id" id="restockID">
                <div class="form-group"><label>入库数量</label><input type="number" name="add_qty" required min="0.1" step="0.1"></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('restockModal')">取消</button>
                    <button type="submit" class="btn btn-success">确认入库</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    function openAddModal() { document.getElementById('addModal').style.display = 'block'; }
    function openEditModal(id, qty, safety, price) {
        document.getElementById('editID').value = id;
        document.getElementById('editQty').value = qty;
        document.getElementById('editSafety').value = safety;
        document.getElementById('editPrice').value = price;
        document.getElementById('editModal').style.display = 'block';
    }
    function openRestockModal(id, name) {
        document.getElementById('restockID').value = id;
        document.getElementById('restockName').innerText = name;
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
