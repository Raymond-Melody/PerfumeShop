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
' ========== 状态消息 ==========
Dim msg, msgType
msg = ""
msgType = "success"

' ========== POST处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        msg = "安全令牌验证失败，请刷新页面后重试"
        msgType = "error"
    Else
        Dim postAction : postAction = Trim(Request.Form("action"))
        
        If postAction = "add" Or postAction = "edit" Then
            Dim fpID : fpID = SafeNum(Request.Form("fixed_product_id"))
            Dim pName : pName = SafeSQL(Trim(Request.Form("product_name")))
            Dim pSpec : pSpec = SafeSQL(Trim(Request.Form("specification")))
            Dim pCode : pCode = SafeSQL(Trim(Request.Form("product_code")))
            Dim uPrice : uPrice = SafeNum(Request.Form("unit_price"))
            Dim sPrice : sPrice = SafeNum(Request.Form("sale_price"))
            Dim sID : sID = SafeNum(Request.Form("supplier_id"))
            Dim sName : sName = SafeSQL(Trim(Request.Form("supplier_name")))
            Dim minQty : minQty = SafeNum(Request.Form("min_order_qty"))
            Dim leadDays : leadDays = SafeNum(Request.Form("lead_time_days"))
            Dim safetyStock : safetyStock = SafeNum(Request.Form("safety_stock"))
            Dim paramMode : paramMode = Trim(Request.Form("param_mode"))
            If paramMode = "" Then paramMode = "Manual"
            Dim pImg : pImg = SafeSQL(Trim(Request.Form("image_url")))
            Dim refProdID : refProdID = SafeNum(Request.Form("ref_product_id"))
            
            If pName = "" Then
                msg = "产品名称不能为空"
                msgType = "error"
            ElseIf uPrice <= 0 Then
                msg = "采购单价必须大于0"
                msgType = "error"
            Else
                If postAction = "add" Then
                    If pCode = "" Then pCode = "FB-" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & "-" & Right("000" & (SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandProducts WHERE ProductCode LIKE 'FB-" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & "-%'")) + 1), 3)
                    
                    Dim insSQL : insSQL = "INSERT INTO FixedBrandProducts (ProductID, ProductCode, ProductName, Specification, UnitPrice, SalePrice, SupplierID, SupplierName, MinOrderQty, LeadTimeDays, SafetyStockManual, LeadTimeDaysManual, ImageURL) VALUES (" & _
                        IIf(refProdID > 0, refProdID, "NULL") & ", '" & pCode & "', '" & pName & "', '" & pSpec & "', " & uPrice & ", " & sPrice & ", " & sID & ", '" & sName & "', " & minQty & ", " & leadDays & ", " & safetyStock & ", " & leadDays & ", '" & pImg & "')"
                    
                    If ExecuteNonQuery(insSQL) Then
                        Dim newFPID : newFPID = SafeNum(GetScalar("SELECT MAX(FixedProductID) FROM FixedBrandProducts"))
                        If newFPID > 0 Then
                            Call ExecuteNonQuery("INSERT INTO FixedBrandInventory (FixedProductID, ProductCode, ProductName, Specification, SafetyStock, MinOrderQty, TotalSold, ParamMode) VALUES (" & newFPID & ", '" & pCode & "', '" & pName & "', '" & pSpec & "', " & safetyStock & ", " & minQty & ", 0, '" & paramMode & "')")
                            
                            ' ========== 同步到 Products 表 ==========
                            If refProdID > 0 Then
                                ' 更新已存在的 Products 记录
                                Call ExecuteNonQuery("UPDATE Products SET ProductName='" & pName & "', BasePrice=" & sPrice & ", Description='" & pSpec & "', UpdatedAt=GETDATE() WHERE ProductID=" & refProdID)
                            Else
                                ' 创建新的 Products 记录
                                Dim insProdSQL : insProdSQL = "INSERT INTO Products (ProductType, ProductName, Description, BasePrice, IsActive) VALUES ('Fixed', '" & pName & "', '" & pSpec & "', " & sPrice & ", 1)"
                                If ExecuteNonQuery(insProdSQL) Then
                                    Dim newProdID : newProdID = SafeNum(GetScalar("SELECT MAX(ProductID) FROM Products"))
                                    If newProdID > 0 Then
                                        Call ExecuteNonQuery("UPDATE FixedBrandProducts SET ProductID=" & newProdID & " WHERE FixedProductID=" & newFPID)
                                    End If
                                End If
                            End If
                        End If
                        msg = "产品添加成功"
                        msgType = "success"
                    Else
                        msg = "添加失败：" & Session("LastDBError")
                        msgType = "error"
                    End If
                Else
                    If fpID <= 0 Then
                        msg = "无效的产品ID"
                        msgType = "error"
                    Else
                        Dim updSQL : updSQL = "UPDATE FixedBrandProducts SET ProductName='" & pName & "', Specification='" & pSpec & "', UnitPrice=" & uPrice & ", SalePrice=" & sPrice & ", SupplierID=" & sID & ", SupplierName='" & sName & "', MinOrderQty=" & minQty & ", LeadTimeDays=" & leadDays & ", SafetyStockManual=" & safetyStock & ", LeadTimeDaysManual=" & leadDays & ", ImageURL='" & pImg & "', UpdatedAt=GETDATE()"
                        If refProdID > 0 Then updSQL = updSQL & ", ProductID=" & refProdID
                        updSQL = updSQL & " WHERE FixedProductID=" & fpID
                        
                        If ExecuteNonQuery(updSQL) Then
                            Call ExecuteNonQuery("UPDATE FixedBrandInventory SET ProductName='" & pName & "', Specification='" & pSpec & "', SafetyStock=" & safetyStock & ", MinOrderQty=" & minQty & ", ParamMode='" & paramMode & "', UpdatedAt=GETDATE() WHERE FixedProductID=" & fpID)
                            
                            ' ========== 同步到 Products 表 ==========
                            If refProdID > 0 Then
                                Call ExecuteNonQuery("UPDATE Products SET ProductName='" & pName & "', BasePrice=" & sPrice & ", Description='" & pSpec & "', UpdatedAt=GETDATE() WHERE ProductID=" & refProdID)
                            Else
                                ' 检查 FixedBrandProducts 是否已有 ProductID
                                Dim existPID : existPID = SafeNum(GetScalar("SELECT ISNULL(ProductID,0) FROM FixedBrandProducts WHERE FixedProductID=" & fpID))
                                If existPID > 0 Then
                                    Call ExecuteNonQuery("UPDATE Products SET ProductName='" & pName & "', BasePrice=" & sPrice & ", Description='" & pSpec & "', UpdatedAt=GETDATE() WHERE ProductID=" & existPID)
                                Else
                                    Dim insProdEditSQL : insProdEditSQL = "INSERT INTO Products (ProductType, ProductName, Description, BasePrice, IsActive) VALUES ('Fixed', '" & pName & "', '" & pSpec & "', " & sPrice & ", 1)"
                                    If ExecuteNonQuery(insProdEditSQL) Then
                                        Dim newProdEditID : newProdEditID = SafeNum(GetScalar("SELECT MAX(ProductID) FROM Products"))
                                        If newProdEditID > 0 Then
                                            Call ExecuteNonQuery("UPDATE FixedBrandProducts SET ProductID=" & newProdEditID & " WHERE FixedProductID=" & fpID)
                                        End If
                                    End If
                                End If
                            End If
                            
                            msg = "产品更新成功"
                            msgType = "success"
                        Else
                            msg = "更新失败：" & Session("LastDBError")
                            msgType = "error"
                        End If
                    End If
                End If
            End If
            
        ElseIf postAction = "delete" Then
            If Not isManager Then
                msg = "权限不足：仅管理员和采购经理可删除产品"
                msgType = "error"
            Else
                Dim delID : delID = SafeNum(Request.Form("fixed_product_id"))
                If delID > 0 Then
                    ' 级联删除：先清理关联数据
                    ' 删除该产品的收货明细
                    Call ExecuteNonQuery("DELETE FROM FixedBrandReceiptDetails WHERE FixedProductID=" & delID)
                    ' 删除该产品的采购明细
                    Call ExecuteNonQuery("DELETE FROM FixedBrandPurchaseDetails WHERE FixedProductID=" & delID)
                    ' 删除该产品的成本分摊记录
                    Call ExecuteNonQuery("DELETE FROM FixedBrandCostAllocation WHERE FixedProductID=" & delID)
                    ' 删除库存记录
                    Call ExecuteNonQuery("DELETE FROM FixedBrandInventory WHERE FixedProductID=" & delID)
                    ' 删除产品记录
                    Call ExecuteNonQuery("DELETE FROM FixedBrandProducts WHERE FixedProductID=" & delID)
                    msg = "产品已删除（含关联采购明细、收货记录、成本分摊）"
                    msgType = "success"
                End If
            End If
            
        ElseIf postAction = "toggle_status" Then
            If Not isManager Then
                msg = "权限不足：仅管理员和采购经理可切换状态"
                msgType = "error"
            Else
                Dim togID : togID = SafeNum(Request.Form("fixed_product_id"))
                If togID > 0 Then
                    Call ExecuteNonQuery("UPDATE FixedBrandProducts SET Status = IIF(Status='Active','Inactive','Active'), UpdatedAt=GETDATE() WHERE FixedProductID=" & togID)
                    msg = "状态已切换"
                    msgType = "success"
                End If
            End If
            
        ElseIf postAction = "import_products" Then
            If Not isManager Then
                msg = "权限不足：仅管理员和采购经理可导入产品"
                msgType = "error"
            Else
            Dim impCount : impCount = 0
            Dim rsImp : Set rsImp = conn.Execute("SELECT ProductID, ProductName, BasePrice FROM Products WHERE ProductType='Fixed' AND IsActive=1")
            If Not rsImp Is Nothing Then
                Do While Not rsImp.EOF
                    Dim impPID : impPID = SafeNum(rsImp("ProductID"))
                    Dim impName : impName = CStr(rsImp("ProductName"))
                    Dim impPrice : impPrice = SafeNum(rsImp("BasePrice"))
                    
                    ' 检查是否已存在
                    Dim existCnt : existCnt = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandProducts WHERE ProductID=" & impPID))
                    If existCnt = 0 And impPID > 0 Then
                        Dim impCode : impCode = "FB-" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & "-IMP" & Right("00" & (impCount + 1), 2)
                        Call ExecuteNonQuery("INSERT INTO FixedBrandProducts (ProductID, ProductCode, ProductName, UnitPrice, SalePrice, SafetyStockManual, LeadTimeDaysManual) VALUES (" & impPID & ", '" & impCode & "', '" & SafeSQL(impName) & "', " & impPrice & ", " & impPrice & ", 10, 7)")
                        Dim newFP : newFP = SafeNum(GetScalar("SELECT MAX(FixedProductID) FROM FixedBrandProducts"))
                        If newFP > 0 Then
                            Call ExecuteNonQuery("INSERT INTO FixedBrandInventory (FixedProductID, ProductCode, ProductName, SafetyStock, ParamMode) VALUES (" & newFP & ", '" & impCode & "', '" & SafeSQL(impName) & "', 10, 'Manual')")
                        End If
                        impCount = impCount + 1
                    End If
                    rsImp.MoveNext
                Loop
                rsImp.Close
            End If
            Set rsImp = Nothing
            msg = "已导入 " & impCount & " 个品牌定香产品"
            msgType = "success"
            End If
        End If
    End If
End If

' ========== 查询产品列表 ==========
Dim searchKey, statusFilter
searchKey = Trim(Request.QueryString("search"))
statusFilter = Trim(Request.QueryString("status"))

Dim whereSQL : whereSQL = " WHERE 1=1"
If searchKey <> "" Then
    whereSQL = whereSQL & " AND (ProductName LIKE '%" & SafeSQL(searchKey) & "%' OR ProductCode LIKE '%" & SafeSQL(searchKey) & "%')"
End If
If statusFilter = "Active" Then
    whereSQL = whereSQL & " AND Status='Active'"
ElseIf statusFilter = "Inactive" Then
    whereSQL = whereSQL & " AND Status='Inactive'"
End If

Dim sqlProducts : sqlProducts = "SELECT fp.*, ISNULL(fi.StockQty,0) AS StockQty, ISNULL(fi.SafetyStock,10) AS SafetyStock, ISNULL(fi.ParamMode,'Manual') AS ParamMode FROM FixedBrandProducts fp LEFT JOIN FixedBrandInventory fi ON fp.FixedProductID=fi.FixedProductID" & whereSQL & " ORDER BY fp.FixedProductID DESC"
Dim rsProducts : Set rsProducts = conn.Execute(sqlProducts)

' ========== 统计 ==========
Dim totalProducts : totalProducts = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandProducts"))
Dim activeProducts : activeProducts = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandProducts WHERE Status='Active'"))
Dim lowStockCount : lowStockCount = SafeNum(GetScalar("SELECT COUNT(*) FROM FixedBrandInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
Dim totalInventory : totalInventory = SafeNum(GetScalar("SELECT ISNULL(SUM(StockQty),0) FROM FixedBrandInventory"))

' ========== 供应商列表(供表单使用) ==========
Dim rsSuppliers : Set rsSuppliers = conn.Execute("SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>品牌定香产品管理 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 270px; padding: 25px; min-height: 100vh; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; padding: 18px; border: 1px solid rgba(255,255,255,0.05); }
        .stat-card .stat-icon { width: 40px; height: 40px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 16px; margin-bottom: 10px; }
        .stat-card .stat-value { font-size: 22px; font-weight: 700; color: #fff; }
        .stat-card .stat-label { font-size: 12px; color: #888; margin-top: 4px; }
        
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .page-title { font-size: 20px; font-weight: 600; color: #fff; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #FF9800; }
        
        .toolbar { display: flex; gap: 10px; align-items: center; margin-bottom: 20px; flex-wrap: wrap; }
        .search-input { padding: 8px 14px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #2d2d44; color: #e0e0e0; width: 250px; font-size: 13px; }
        .search-input::placeholder { color: #666; }
        .status-filter { padding: 8px 14px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #2d2d44; color: #e0e0e0; font-size: 13px; }
        
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 10px; overflow: hidden; }
        .data-table th, .data-table td { padding: 12px 15px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 13px; }
        .data-table th { color: #888; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; background: rgba(0,0,0,0.2); }
        .data-table td { color: #ccc; }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-active { color: #4CAF50; }
        .status-inactive { color: #F44336; }
        .stock-low { color: #FF9800; font-weight: 600; }
        .stock-ok { color: #4CAF50; }
        .stock-zero { color: #F44336; }
        
        .action-btns { display: flex; gap: 6px; }
        
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); z-index: 1000; justify-content: center; align-items: center; }
        .modal-overlay.show { display: flex; }
        .modal-box { background: #2d2d44; border-radius: 12px; padding: 25px; width: 520px; max-height: 85vh; overflow-y: auto; border: 1px solid rgba(255,255,255,0.1); box-shadow: 0 10px 40px rgba(0,0,0,0.5); }
        .modal-box h3 { color: #fff; font-size: 18px; margin: 0 0 20px; display: flex; align-items: center; gap: 8px; }
        .modal-box h3 i { color: #FF9800; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input, .form-group select { width: 100%; padding: 9px 12px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1); background: #1a1a2e; color: #e0e0e0; font-size: 13px; box-sizing: border-box; }
        .form-row { display: flex; gap: 12px; }
        .form-row .form-group { flex: 1; }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        
        .confirm-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); z-index: 1100; justify-content: center; align-items: center; }
        .confirm-overlay.show { display: flex; }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="../includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-boxes"></i> 品牌定香产品管理</h2>
            <div class="breadcrumb" style="font-size:13px;color:#888;">
                <a href="index.asp" style="color:#FF9800;text-decoration:none;">品牌定香采购</a> / 产品管理
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div style="padding:12px 20px; border-radius:8px; margin-bottom:20px; font-size:14px; background:<%=IIf(msgType="success","rgba(76,175,80,0.15)","rgba(244,67,54,0.15)")%>; color:<%=IIf(msgType="success","#4CAF50","#F44336")%>; border:1px solid <%=IIf(msgType="success","rgba(76,175,80,0.3)","rgba(244,67,54,0.3)")%>;">
            <i class="fas fa-<%=IIf(msgType="success","check-circle","exclamation-circle")%>"></i> <%= msg %>
        </div>
        <% End If %>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#4CAF50,#388E3C);"><i class="fas fa-cubes"></i></div>
                <div class="stat-value"><%= totalProducts %></div>
                <div class="stat-label">产品总数</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#2196F3,#1565C0);"><i class="fas fa-check-circle"></i></div>
                <div class="stat-value"><%= activeProducts %></div>
                <div class="stat-label">活跃产品</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#FF9800,#F57C00);"><i class="fas fa-exclamation-triangle"></i></div>
                <div class="stat-value" style="color:<%=IIf(lowStockCount>0,"#FF9800","#fff")%>;"><%= lowStockCount %></div>
                <div class="stat-label">低库存预警</div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background:linear-gradient(135deg,#9C27B0,#6A1B9A);"><i class="fas fa-warehouse"></i></div>
                <div class="stat-value"><%= totalInventory %></div>
                <div class="stat-label">总库存量</div>
            </div>
        </div>
        
        <div class="toolbar">
            <button class="btn btn--primary" onclick="openModal('add')"><i class="fas fa-plus"></i> 添加产品</button>
            <button class="btn btn--success" onclick="importProducts()"><i class="fas fa-download"></i> 从产品库导入</button>
            <form method="get" style="display:flex;gap:8px;margin-left:auto;">
                <select name="status" class="status-filter" onchange="this.form.submit()">
                    <option value="">全部状态</option>
                    <option value="Active" <%= IIf(statusFilter="Active","selected","") %>>活跃</option>
                    <option value="Inactive" <%= IIf(statusFilter="Inactive","selected","") %>>已停用</option>
                </select>
                <input type="text" name="search" class="search-input" placeholder="搜索产品名称或编码..." value="<%=Server.HTMLEncode(searchKey)%>">
                <button type="submit" class="btn btn--primary btn--sm"><i class="fas fa-search"></i></button>
                <% If searchKey <> "" Or statusFilter <> "" Then %>
                <a href="product_management.asp" class="btn btn--neutral btn--sm"><i class="fas fa-times"></i> 清除</a>
                <% End If %>
            </form>
        </div>
        
        <table class="data-table">
            <thead>
                <tr>
                    <th>产品编码</th>
                    <th>产品名称</th>
                    <th>规格</th>
                    <th>采购单价</th>
                    <th>零售价</th>
                    <th>供应商</th>
                    <th>库存</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsProducts Is Nothing Then
                    If Not rsProducts.EOF Then
                        Do While Not rsProducts.EOF
                            Dim stockQty : stockQty = SafeNum(rsProducts("StockQty"))
                            safetyStock = SafeNum(GetScalar("SELECT ISNULL(SafetyStock,10) FROM FixedBrandInventory WHERE FixedProductID=" & SafeNum(rsProducts("FixedProductID"))))
                            Dim stockClass : stockClass = "stock-ok"
                            If stockQty <= 0 Then stockClass = "stock-zero" Else If stockQty <= safetyStock Then stockClass = "stock-low"
                %>
                <tr>
                    <td><span style="font-family:Consolas,monospace;color:#FF9800;"><%=Server.HTMLEncode(CStr(rsProducts("ProductCode")))%></span></td>
                    <td><%=Server.HTMLEncode(CStr(rsProducts("ProductName")))%></td>
                    <td><%=Server.HTMLEncode(CStr(rsProducts("Specification") & ""))%></td>
                    <td>¥<%= FormatNumber(SafeNum(rsProducts("UnitPrice")), 2) %></td>
                    <td>¥<%= FormatNumber(SafeNum(rsProducts("SalePrice")), 2) %></td>
                    <td><%=Server.HTMLEncode(CStr(rsProducts("SupplierName") & ""))%></td>
                    <td><span class="<%= stockClass %>"><%= stockQty %></span></td>
                    <td><span class="<%= IIf(CStr(rsProducts("Status"))="Active","status-active","status-inactive") %>"><%= IIf(CStr(rsProducts("Status"))="Active","活跃","已停用") %></span></td>
                    <td>
                        <div class="action-btns">
                            <button class="btn btn--primary btn--xs" onclick="editProduct(<%= rsProducts("FixedProductID") %>, '<%= Server.HTMLEncode(CStr(rsProducts("ProductName"))) %>', '<%= Server.HTMLEncode(CStr(rsProducts("Specification") & "")) %>', '<%= Server.HTMLEncode(CStr(rsProducts("ProductCode"))) %>', <%= SafeNum(rsProducts("UnitPrice")) %>, <%= SafeNum(rsProducts("SalePrice")) %>, <%= SafeNum(rsProducts("SupplierID")) %>, '<%= Server.HTMLEncode(CStr(rsProducts("SupplierName") & "")) %>', <%= SafeNum(rsProducts("MinOrderQty")) %>, <%= SafeNum(rsProducts("LeadTimeDays")) %>, '<%= Server.HTMLEncode(CStr(rsProducts("ImageURL") & "")) %>', <%= SafeNum(rsProducts("ProductID")) %>, <%= SafeNum(rsProducts("SafetyStock")) %>, '<%= Server.HTMLEncode(IIf(CStr(rsProducts("ParamMode") & "") = "", "Manual", CStr(rsProducts("ParamMode")))) %>')" title="编辑"><i class="fas fa-edit"></i></button>
                            <button class="btn btn--warning btn--xs" onclick="toggleStatus(<%= rsProducts("FixedProductID") %>)" title="<%= IIf(CStr(rsProducts("Status"))="Active","停用","启用") %>"><i class="fas fa-<%= IIf(CStr(rsProducts("Status"))="Active","ban","check") %>"></i></button>
                            <button class="btn btn--danger btn--xs" onclick="confirmDelete(<%= rsProducts("FixedProductID") %>, '<%= Server.HTMLEncode(CStr(rsProducts("ProductName"))) %>')" title="删除"><i class="fas fa-trash"></i></button>
                        </div>
                    </td>
                </tr>
                <%
                            rsProducts.MoveNext
                        Loop
                    Else
                %>
                <tr><td colspan="9" style="text-align:center;padding:40px;color:#666;">
                    <i class="fas fa-inbox" style="font-size:28px;display:block;margin-bottom:10px;"></i>暂无产品数据
                </td></tr>
                <%      End If
                    rsProducts.Close
                    Set rsProducts = Nothing
                Else
                %>
                <tr><td colspan="9" style="text-align:center;padding:40px;color:#666;">数据加载失败</td></tr>
                <% End If %>
            </tbody>
        </table>
    </div>
    
    <!-- 添加/编辑弹窗 -->
    <div class="modal-overlay" id="productModal">
        <div class="modal-box">
            <h3><i class="fas fa-box"></i> <span id="modalTitle">添加产品</span></h3>
            <form method="post" id="productForm">
                <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                <input type="hidden" name="action" id="formAction" value="add">
                <input type="hidden" name="fixed_product_id" id="formFPID" value="0">
                
                <div class="form-row">
                    <div class="form-group">
                        <label>产品编码</label>
                        <input type="text" name="product_code" id="formCode" placeholder="自动生成">
                    </div>
                    <div class="form-group">
                        <label>关联产品ID</label>
                        <input type="number" name="ref_product_id" id="formRefPID" value="0" placeholder="Products表ID">
                    </div>
                </div>
                <div class="form-group">
                    <label><span style="color:#F44336;">*</span> 产品名称</label>
                    <input type="text" name="product_name" id="formName" required placeholder="产品名称">
                </div>
                <div class="form-group">
                    <label>规格</label>
                    <input type="text" name="specification" id="formSpec" placeholder="如: 50ml">
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label><span style="color:#F44336;">*</span> 采购单价</label>
                        <input type="number" name="unit_price" id="formUnitPrice" step="0.01" required placeholder="0.00">
                    </div>
                    <div class="form-group">
                        <label>建议零售价</label>
                        <input type="number" name="sale_price" id="formSalePrice" step="0.01" placeholder="0.00">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>供应商</label>
                        <select name="supplier_id" id="formSupplierID" onchange="updateSupplierName()">
                            <option value="0">-- 选择供应商 --</option>
                            <% If Not rsSuppliers Is Nothing Then
                                Do While Not rsSuppliers.EOF %>
                            <option value="<%= rsSuppliers("SupplierID") %>" data-name="<%= Server.HTMLEncode(CStr(rsSuppliers("SupplierName"))) %>"><%= Server.HTMLEncode(CStr(rsSuppliers("SupplierName"))) %></option>
                            <%      rsSuppliers.MoveNext
                                Loop
                                rsSuppliers.Close
                                Set rsSuppliers = Nothing
                            End If %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>供应商名称</label>
                        <input type="text" name="supplier_name" id="formSupplierName" readonly style="background:#222;">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>最小起订量</label>
                        <input type="number" name="min_order_qty" id="formMinQty" value="1" min="1">
                    </div>
                    <div class="form-group">
                        <label>交货周期(天)</label>
                        <input type="number" name="lead_time_days" id="formLeadDays" value="7" min="1">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>安全库存</label>
                        <input type="number" name="safety_stock" id="formSafetyStock" value="10" min="0">
                        <span style="font-size:10px;color:#888;">低于此值时触发补货预警</span>
                    </div>
                    <div class="form-group">
                        <label>参数模式</label>
                        <select name="param_mode" id="formParamMode" style="width:100%;padding:9px 12px;border-radius:6px;border:1px solid rgba(255,255,255,0.1);background:#1a1a2e;color:#e0e0e0;font-size:13px;">
                            <option value="Manual">人工设定 - 使用手动输入参数</option>
                            <option value="Auto">统计推定 - 根据历史数据自动计算</option>
                        </select>
                        <span style="font-size:10px;color:#888;">Auto模式需积累足够历史数据后生效</span>
                    </div>
                </div>
                <div class="form-group">
                    <label>产品图片URL</label>
                    <input type="text" name="image_url" id="formImage" placeholder="/images/...">
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn--neutral" onclick="closeModal()">取消</button>
                    <button type="submit" class="btn btn--primary">保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 删除确认弹窗 -->
    <div class="confirm-overlay" id="deleteConfirm">
        <div class="modal-box" style="width:400px;text-align:center;">
            <h3 style="justify-content:center;color:#F44336;"><i class="fas fa-exclamation-triangle" style="color:#F44336;"></i> 确认删除</h3>
            <p style="color:#ccc;margin:15px 0;" id="deleteMsg">确定要删除该产品吗？</p>
            <form method="post" id="deleteForm">
                <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="fixed_product_id" id="deleteFPID" value="0">
                <div class="modal-actions" style="justify-content:center;">
                    <button type="button" class="btn btn--neutral" onclick="closeDeleteConfirm()">取消</button>
                    <button type="submit" class="btn btn--danger">确认删除</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function openModal(mode) {
            document.getElementById('productModal').classList.add('show');
            if (mode === 'add') {
                document.getElementById('modalTitle').textContent = '添加产品';
                document.getElementById('formAction').value = 'add';
                document.getElementById('formFPID').value = '0';
                document.getElementById('formCode').value = '';
                document.getElementById('formName').value = '';
                document.getElementById('formSpec').value = '';
                document.getElementById('formUnitPrice').value = '';
                document.getElementById('formSalePrice').value = '';
                document.getElementById('formSupplierID').value = '0';
                document.getElementById('formSupplierName').value = '';
                document.getElementById('formMinQty').value = '1';
                document.getElementById('formLeadDays').value = '7';
                document.getElementById('formSafetyStock').value = '10';
                document.getElementById('formParamMode').value = 'Manual';
                document.getElementById('formImage').value = '';
                document.getElementById('formRefPID').value = '0';
            }
        }
        
        function closeModal() {
            document.getElementById('productModal').classList.remove('show');
        }
        
        function editProduct(id, name, spec, code, unitPrice, salePrice, sid, sname, minQty, leadDays, img, refPID, safetyStock, paramMode) {
            openModal('edit');
            document.getElementById('modalTitle').textContent = '编辑产品';
            document.getElementById('formAction').value = 'edit';
            document.getElementById('formFPID').value = id;
            document.getElementById('formCode').value = code;
            document.getElementById('formName').value = name;
            document.getElementById('formSpec').value = spec;
            document.getElementById('formUnitPrice').value = unitPrice;
            document.getElementById('formSalePrice').value = salePrice;
            document.getElementById('formSupplierID').value = sid;
            document.getElementById('formSupplierName').value = sname;
            document.getElementById('formMinQty').value = minQty;
            document.getElementById('formLeadDays').value = leadDays;
            document.getElementById('formSafetyStock').value = safetyStock || 10;
            document.getElementById('formParamMode').value = paramMode || 'Manual';
            document.getElementById('formImage').value = img;
            document.getElementById('formRefPID').value = refPID;
        }
        
        function updateSupplierName() {
            var sel = document.getElementById('formSupplierID');
            var name = sel.options[sel.selectedIndex].getAttribute('data-name');
            document.getElementById('formSupplierName').value = name || '';
        }
        
        function confirmDelete(id, name) {
            document.getElementById('deleteFPID').value = id;
            document.getElementById('deleteMsg').textContent = '确定要删除产品 "' + name + '" 吗？此操作不可恢复。';
            document.getElementById('deleteConfirm').classList.add('show');
        }
        
        function closeDeleteConfirm() {
            document.getElementById('deleteConfirm').classList.remove('show');
        }
        
        function toggleStatus(id) {
            var f = document.createElement('form');
            f.method = 'post';
            f.innerHTML = '<input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>"><input type="hidden" name="action" value="toggle_status"><input type="hidden" name="fixed_product_id" value="' + id + '">';
            document.body.appendChild(f);
            f.submit();
        }
        
        function importProducts() {
            if (confirm('将从产品库(Products表)导入所有ProductType="Fixed"的产品。已存在的不会重复导入。')) {
                var f = document.createElement('form');
                f.method = 'post';
                f.innerHTML = '<input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>"><input type="hidden" name="action" value="import_products">';
                document.body.appendChild(f);
                f.submit();
            }
        }
        
        // 点击遮罩关闭
        document.getElementById('productModal').addEventListener('click', function(e) { if (e.target === this) closeModal(); });
        document.getElementById('deleteConfirm').addEventListener('click', function(e) { if (e.target === this) closeDeleteConfirm(); });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
