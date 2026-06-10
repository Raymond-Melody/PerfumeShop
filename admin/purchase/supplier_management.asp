<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<%
Call OpenConnection()

' ========== SafeNum函数：安全处理数值空值 ==========
Function SafeNum(val)
    If IsNull(val) Then
        SafeNum = 0
    ElseIf val = "" Then
        SafeNum = 0
    ElseIf Not IsNumeric(val) Then
        SafeNum = 0
    Else
        SafeNum = CDbl(val)
    End If
End Function

' ========== SafeOutput函数：安全输出到JavaScript/HTML（转义单引号+HTML编码）==========
Function SafeOutput(str)
    If IsNull(str) Or str = "" Then
        SafeOutput = ""
        Exit Function
    End If
    ' 先转义单引号为 \' (JS安全)，再HTML编码
    Dim tmp
    tmp = Replace(CStr(str), "'", "\'")
    SafeOutput = Server.HTMLEncode(tmp)
End Function

' ========== 获取供应商详情（用于详情弹窗）==========
Dim detailSupplierId
detailSupplierId = Request.QueryString("detailId")
Dim rsDetail, detailExists
detailExists = False
If detailSupplierId <> "" Then
    If IsNumeric(detailSupplierId) Then
        Set rsDetail = ExecuteQuery("SELECT * FROM Suppliers WHERE SupplierID = " & CInt(detailSupplierId))
        If Not rsDetail Is Nothing Then
            If Not rsDetail.EOF Then
                detailExists = True
            End If
        End If
' ========== V12: 获取供应商合作历史与评分 ==========
Dim rsSupplierOrders, supplierRatingAvg, supplierEvalCount, supplierActiveContracts
If detailExists Then
    ' 最近5笔采购订单
    Set rsSupplierOrders = ExecuteQuery("SELECT TOP 5 PurchaseID, PurchaseNo, OrderDate, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status FROM PurchaseOrders WHERE SupplierID=" & CInt(detailSupplierId) & " ORDER BY OrderDate DESC")
    
    ' 供应商综合评分
    Dim ratingRS : Set ratingRS = ExecuteQuery("SELECT AVG(CAST(OverallScore AS FLOAT)) as AvgScore, COUNT(*) as EvalCount FROM SupplierEvaluations WHERE SupplierID=" & CInt(detailSupplierId))
    If Not ratingRS Is Nothing Then
        If Not ratingRS.EOF Then
            supplierRatingAvg = SafeNum(ratingRS("AvgScore"))
            supplierEvalCount = SafeNum(ratingRS("EvalCount"))
        End If
        ratingRS.Close : Set ratingRS = Nothing
    End If
    
    ' 活跃合同数
    supplierActiveContracts = SafeNum(GetScalar("SELECT COUNT(*) FROM SupplierContracts WHERE SupplierID=" & CInt(detailSupplierId) & " AND Status='Active'"))
End If
    End If
End If

' ==================== POST 处理：CRUD ====================
Dim action, supplierId, supplierName, contactPerson, phone, email, address, category, notes, isActiveVal
action = Request.Form("action")

' 权限检查：PURCHASE_STAFF只能查看和编辑，不能新增/禁用
Dim canModify
canModify = False
If Session("AdminRoleCode") = "SUPER_ADMIN" Then
    canModify = True
ElseIf Session("AdminRoleCode") = "PURCHASE_MANAGER" Then
    canModify = True
End If

If action = "add" And canModify Then
    supplierName = SafeSQL(Request.Form("supplierName"))
    contactPerson = SafeSQL(Request.Form("contactPerson"))
    phone = SafeSQL(Request.Form("phone"))
    email = SafeSQL(Request.Form("email"))
    address = SafeSQL(Request.Form("address"))
    category = SafeSQL(Request.Form("category"))
    notes = SafeSQL(Request.Form("notes"))
    
    If supplierName <> "" Then
        Dim insertSql
        insertSql = "INSERT INTO Suppliers (SupplierName, ContactPerson, Phone, Email, Address, Category, Notes, IsActive, CreatedAt) VALUES (" & _
            "'" & supplierName & "', " & _
            "'" & contactPerson & "', " & _
            "'" & phone & "', " & _
            "'" & email & "', " & _
            "'" & address & "', " & _
            "'" & category & "', " & _
            "'" & notes & "', " & _
            "1, GETDATE())"
        If ExecuteNonQuery(insertSql) Then
            Response.Redirect "supplier_management.asp?msg=" & Server.URLEncode("供应商添加成功")
        Else
            Response.Redirect "supplier_management.asp?msg=" & Server.URLEncode("供应商添加失败")
        End If
    End If
ElseIf action = "edit" Then
    supplierId = Request.Form("supplierId")
    supplierName = SafeSQL(Request.Form("supplierName"))
    contactPerson = SafeSQL(Request.Form("contactPerson"))
    phone = SafeSQL(Request.Form("phone"))
    email = SafeSQL(Request.Form("email"))
    address = SafeSQL(Request.Form("address"))
    category = SafeSQL(Request.Form("category"))
    notes = SafeSQL(Request.Form("notes"))
    
    If IsNumeric(supplierId) Then
        If supplierName <> "" Then
            Dim updateSql
            updateSql = "UPDATE Suppliers SET " & _
                "SupplierName = '" & supplierName & "', " & _
                "ContactPerson = '" & contactPerson & "', " & _
                "Phone = '" & phone & "', " & _
                "Email = '" & email & "', " & _
                "Address = '" & address & "', " & _
                "Category = '" & category & "', " & _
                "Notes = '" & notes & "' " & _
                "WHERE SupplierID = " & CInt(supplierId)
            If ExecuteNonQuery(updateSql) Then
                Response.Redirect "supplier_management.asp?msg=" & Server.URLEncode("供应商信息更新成功")
            Else
                Response.Redirect "supplier_management.asp?msg=" & Server.URLEncode("供应商信息更新失败")
            End If
        End If
    End If
ElseIf action = "toggle" And canModify Then
    supplierId = Request.Form("supplierId")
    If IsNumeric(supplierId) Then
        Dim toggleSql
        toggleSql = "UPDATE Suppliers SET IsActive = IIF(IsActive <> 0, 0, 1) WHERE SupplierID = " & CInt(supplierId)
        If ExecuteNonQuery(toggleSql) Then
            Response.Redirect "supplier_management.asp?msg=" & Server.URLEncode("状态切换成功")
        Else
            Response.Redirect "supplier_management.asp?msg=" & Server.URLEncode("状态切换失败")
        End If
    End If
End If

' ========== 合同管理 POST ==========
If action = "contract_add" And canModify Then
    Dim cSupplierID, cContractNo, cContractName, cStartDate, cEndDate, cTotalAmount, cPaymentTerms, cTermsSummary
    cSupplierID = SafeNum(Request.Form("supplier_id"))
    cContractNo = Trim(Request.Form("contract_no"))
    cContractName = Trim(Request.Form("contract_name"))
    cStartDate = Trim(Request.Form("start_date"))
    cEndDate = Trim(Request.Form("end_date"))
    cTotalAmount = SafeNum(Request.Form("total_amount"))
    cPaymentTerms = Trim(Request.Form("payment_terms"))
    cTermsSummary = Trim(Request.Form("terms_summary"))
    
    If cSupplierID > 0 And cContractName <> "" Then
        conn.Execute "INSERT INTO SupplierContracts (SupplierID, ContractNo, ContractName, StartDate, EndDate, TotalAmount, PaymentTerms, TermsSummary, Status) VALUES (" & _
            cSupplierID & ",'" & SafeSQL(cContractNo) & "','" & SafeSQL(cContractName) & "'," & _
            IIf(cStartDate<>"", "'" & cStartDate & "'", "NULL") & "," & _
            IIf(cEndDate<>"", "'" & cEndDate & "'", "NULL") & "," & _
            cTotalAmount & ",'" & SafeSQL(cPaymentTerms) & "','" & SafeSQL(cTermsSummary) & "','Active')"
        Response.Redirect "supplier_management.asp?tab=contracts&msg=" & Server.URLEncode("合同添加成功")
        Response.End
    End If
ElseIf action = "contract_toggle" And canModify Then
    Dim cID : cID = SafeNum(Request.Form("contract_id"))
    If cID > 0 Then
        conn.Execute "UPDATE SupplierContracts SET Status=IIF(Status='Active','Inactive','Active'), UpdatedAt=GETDATE() WHERE ContractID=" & cID
        Response.Redirect "supplier_management.asp?tab=contracts&msg=" & Server.URLEncode("状态切换成功")
        Response.End
    End If
ElseIf action = "eval_add" Then
    Dim evSupplierID, evQuality, evDelivery, evPrice, evService, evComments
    evSupplierID = SafeNum(Request.Form("supplier_id"))
    evQuality = SafeNum(Request.Form("quality_score"))
    evDelivery = SafeNum(Request.Form("delivery_score"))
    evPrice = SafeNum(Request.Form("price_score"))
    evService = SafeNum(Request.Form("service_score"))
    evComments = Trim(Request.Form("comments"))
    
    If evSupplierID > 0 Then
        Dim evOverall : evOverall = Int((evQuality + evDelivery + evPrice + evService) / 4)
        Dim evRating
        If evOverall >= 90 Then
            evRating = "A"
        ElseIf evOverall >= 75 Then
            evRating = "B"
        ElseIf evOverall >= 60 Then
            evRating = "C"
        Else
            evRating = "D"
        End If
        conn.Execute "INSERT INTO SupplierEvaluations (SupplierID, EvaluatedBy, QualityScore, DeliveryScore, PriceScore, ServiceScore, OverallScore, Rating, Comments, Period) VALUES (" & _
            evSupplierID & ",'" & SafeSQL(Session("AdminRealName")) & "'," & evQuality & "," & evDelivery & "," & evPrice & "," & evService & "," & evOverall & ",'" & evRating & "','" & SafeSQL(evComments) & "','" & Year(Now) & "Q" & Int((Month(Now)-1)/3)+1 & "')"
        Response.Redirect "supplier_management.asp?tab=evals&msg=" & Server.URLEncode("评估提交成功")
        Response.End
    End If
End If

' ========== Tab参数 ==========
Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "suppliers"

' ========== 合同数据 ==========
Dim rsContracts
On Error Resume Next
Set rsContracts = conn.Execute("SELECT sc.*, s.SupplierName FROM SupplierContracts sc LEFT JOIN Suppliers s ON sc.SupplierID = s.SupplierID ORDER BY sc.CreatedAt DESC")
If Err.Number <> 0 Then Err.Clear : Set rsContracts = Nothing
On Error GoTo 0

' ========== 评估数据 ==========
Dim rsEvals
On Error Resume Next
Set rsEvals = conn.Execute("SELECT se.*, s.SupplierName FROM SupplierEvaluations se LEFT JOIN Suppliers s ON se.SupplierID = s.SupplierID ORDER BY se.CreatedAt DESC")
If Err.Number <> 0 Then Err.Clear : Set rsEvals = Nothing
On Error GoTo 0
Dim totalSuppliers, activeSuppliers, spiceSuppliers, packageSuppliers, logisticsSuppliers, printSuppliers, spraySuppliers
totalSuppliers = GetScalar("SELECT COUNT(*) FROM Suppliers")
activeSuppliers = GetScalar("SELECT COUNT(*) FROM Suppliers WHERE IsActive <> 0")
spiceSuppliers = GetScalar("SELECT COUNT(*) FROM Suppliers WHERE Category = '香料供应商'")
packageSuppliers = GetScalar("SELECT COUNT(*) FROM Suppliers WHERE Category = '包装供应商'")
logisticsSuppliers = GetScalar("SELECT COUNT(*) FROM Suppliers WHERE Category = '物流合作商'")
printSuppliers = GetScalar("SELECT COUNT(*) FROM Suppliers WHERE Category = '印刷供应商'")
spraySuppliers = GetScalar("SELECT COUNT(*) FROM Suppliers WHERE Category = '喷头供应商'")

' ==================== 获取筛选参数 ====================
Dim filterCategory, filterStatus, searchKeyword
filterCategory = Request.QueryString("category")
filterStatus = Request.QueryString("status")
searchKeyword = Request.QueryString("search")

' ==================== 构建查询条件 ====================
Dim whereClause, hasCondition
whereClause = ""
hasCondition = False

If filterCategory <> "" Then
    whereClause = whereClause & "Category = '" & SafeSQL(filterCategory) & "'"
    hasCondition = True
End If

If filterStatus = "active" Then
    If hasCondition Then
        whereClause = whereClause & " AND "
    End If
    whereClause = whereClause & "IsActive <> 0"
    hasCondition = True
ElseIf filterStatus = "inactive" Then
    If hasCondition Then
        whereClause = whereClause & " AND "
    End If
    whereClause = whereClause & "IsActive = 0"
    hasCondition = True
End If

If searchKeyword <> "" Then
    If hasCondition Then
        whereClause = whereClause & " AND "
    End If
    whereClause = whereClause & "(SupplierName LIKE '%" & SafeSQL(searchKeyword) & "%' OR ContactPerson LIKE '%" & SafeSQL(searchKeyword) & "%')"
    hasCondition = True
End If

If hasCondition Then
    whereClause = "WHERE " & whereClause
End If

' ==================== 获取供应商列表 ====================
Dim rsSuppliers, supplierList(), supplierCount, i
supplierCount = 0
Set rsSuppliers = ExecuteQuery("SELECT * FROM Suppliers " & whereClause & " ORDER BY CreatedAt DESC")

If Not rsSuppliers Is Nothing Then
    If Not rsSuppliers.EOF Then
        rsSuppliers.MoveLast
        supplierCount = rsSuppliers.RecordCount
        rsSuppliers.MoveFirst
        
        ReDim supplierList(supplierCount - 1, 9)
        i = 0
        Do While Not rsSuppliers.EOF
            supplierList(i, 0) = rsSuppliers("SupplierID")
            supplierList(i, 1) = rsSuppliers("SupplierName")
            supplierList(i, 2) = rsSuppliers("ContactPerson")
            supplierList(i, 3) = rsSuppliers("Phone")
            supplierList(i, 4) = rsSuppliers("Email")
            supplierList(i, 5) = rsSuppliers("Address")
            supplierList(i, 6) = rsSuppliers("Category")
            supplierList(i, 7) = rsSuppliers("Notes")
            supplierList(i, 8) = rsSuppliers("IsActive")
            supplierList(i, 9) = rsSuppliers("CreatedAt")
            i = i + 1
            rsSuppliers.MoveNext
        Loop
    End If
    rsSuppliers.Close
    Set rsSuppliers = Nothing
End If

' ==================== 获取供应商关联数据（采购订单和报价数量）====================
Function GetSupplierOrderCount(sid)
    Dim cnt
    cnt = 0
    If IsNumeric(sid) Then
        cnt = GetScalar("SELECT COUNT(*) FROM PurchaseOrders WHERE SupplierID = " & CInt(sid))
    End If
    GetSupplierOrderCount = cnt
End Function

Function GetSupplierPriceCount(sid)
    Dim cnt
    cnt = 0
    If IsNumeric(sid) Then
        cnt = GetScalar("SELECT COUNT(*) FROM SupplierPrices WHERE SupplierID = " & CInt(sid) & " AND IsActive <> 0")
    End If
    GetSupplierPriceCount = cnt
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>供应商管理 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 统计卡片 */
        .stats-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); 
            gap: 20px; 
            margin-bottom: 30px; 
        }
        .stat-card { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            padding: 20px; 
            border-radius: 12px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            text-align: center; 
            transition: transform 0.2s ease;
        }
        .stat-card:hover {
            transform: translateY(-3px);
        }
        .stat-card.total { border-top: 4px solid #FF9800; }
        .stat-card.active { border-top: 4px solid #4CAF50; }
        .stat-card.spice { border-top: 4px solid #FFB74D; }
        .stat-card.package { border-top: 4px solid #2196F3; }
        .stat-card.logistics { border-top: 4px solid #9C27B0; }
        .stat-card.printing { border-top: 4px solid #00BCD4; }
        .stat-card.spray { border-top: 4px solid #FF5722; }
        .stat-value { font-size: 28px; font-weight: bold; color: #fff; }
        .stat-label { color: #888; margin-top: 8px; font-size: 13px; }
        
        /* 筛选栏 */
        .filter-bar { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            padding: 20px; 
            border-radius: 12px; 
            margin-bottom: 20px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            display: flex; 
            gap: 15px; 
            flex-wrap: wrap; 
            align-items: center; 
        }
        .filter-bar select, .filter-bar input { 
            padding: 10px 15px; 
            border: 1px solid rgba(255,255,255,0.1); 
            border-radius: 8px; 
            font-size: 14px; 
            background: #1e1e32;
            color: #e0e0e0;
        }
        .filter-bar input[type="text"] { width: 200px; }
        .filter-bar button { 
            padding: 10px 20px; 
            border: none; 
            border-radius: 8px; 
            cursor: pointer; 
            font-size: 14px;
            transition: all 0.2s ease;
        }
        .btn-search { background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%); color: white; }
        .btn-search:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(255,152,0,0.3); }
        .btn-reset { background: #3a3a5c; color: #e0e0e0; }
        .btn-reset:hover { background: #4a4a6c; }
        .btn-add { 
            background: linear-gradient(135deg, #4CAF50 0%, #388E3C 100%); 
            color: white; 
            margin-left: auto;
        }
        .btn-add:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(76,175,80,0.3); }
        .btn-add.disabled {
            background: #3a3a5c;
            cursor: not-allowed;
        }
        
        /* 供应商表格 */
        .supplier-table { 
            width: 100%; 
            border-collapse: collapse; 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px; 
            overflow: hidden; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
        }
        .supplier-table th { 
            background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%);
            color: white; 
            padding: 15px; 
            text-align: left;
            font-weight: 600;
        }
        .supplier-table td { 
            padding: 15px; 
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #e0e0e0;
        }
        .supplier-table tr:hover { background: rgba(255,255,255,0.02); }
        
        .status-badge { 
            display: inline-block; 
            padding: 5px 12px; 
            border-radius: 12px; 
            font-size: 12px; 
            font-weight: 500; 
        }
        .status-active { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-inactive { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        
        .category-badge { 
            display: inline-block; 
            padding: 4px 12px; 
            border-radius: 10px; 
            font-size: 11px;
            font-weight: 500;
        }
        .category-spice { background: rgba(255,183,77,0.2); color: #FFB74D; }
        .category-package { background: rgba(33,150,243,0.2); color: #2196F3; }
        .category-logistics { background: rgba(156,39,176,0.2); color: #9C27B0; }
        .category-printing { background: rgba(0,188,212,0.2); color: #00BCD4; }
        .category-spray { background: rgba(255,87,34,0.2); color: #FF5722; }
        .category-other { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        
        .btn-action { 
            padding: 6px 12px; 
            border: none; 
            border-radius: 6px; 
            cursor: pointer; 
            font-size: 12px; 
            margin-right: 5px;
            transition: all 0.2s ease;
        }
        .btn-action:hover { transform: translateY(-2px); }
        .btn-edit { background: linear-gradient(135deg, #2196F3 0%, #1976D2 100%); color: white; }
        .btn-toggle { background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%); color: white; }
        .btn-toggle.disabled { background: #3a3a5c; cursor: not-allowed; }
        .btn-detail { background: linear-gradient(135deg, #9C27B0 0%, #7B1FA2 100%); color: white; }
        
        /* 模态框 */
        .modal { 
            display: none; 
            position: fixed; 
            top: 0; 
            left: 0; 
            width: 100%; 
            height: 100%; 
            background: rgba(0,0,0,0.7); 
            z-index: 1000;
            backdrop-filter: blur(5px);
        }
        .modal-content { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            width: 90%; 
            max-width: 500px; 
            margin: 50px auto; 
            padding: 30px; 
            border-radius: 12px; 
            max-height: 80vh; 
            overflow-y: auto;
            border: 1px solid rgba(255,255,255,0.1);
            box-shadow: 0 10px 40px rgba(0,0,0,0.5);
        }
        .modal-header { 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            margin-bottom: 25px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            padding-bottom: 15px;
        }
        .modal-header h3 { 
            margin: 0; 
            color: #fff;
            font-size: 20px;
        }
        .close-btn { 
            background: none; 
            border: none; 
            font-size: 28px; 
            cursor: pointer; 
            color: #888;
            transition: color 0.2s ease;
        }
        .close-btn:hover { color: #fff; }
        .form-group { margin-bottom: 20px; }
        .form-group label { 
            display: block; 
            margin-bottom: 8px; 
            font-weight: 500; 
            color: #b0b0b0; 
        }
        .form-group input, .form-group select, .form-group textarea { 
            width: 100%; 
            padding: 12px; 
            border: 1px solid rgba(255,255,255,0.1); 
            border-radius: 8px; 
            font-size: 14px; 
            box-sizing: border-box;
            background: #1e1e32;
            color: #e0e0e0;
        }
        .form-group input:focus, .form-group select:focus, .form-group textarea:focus {
            outline: none;
            border-color: #FF9800;
        }
        .form-group textarea { resize: vertical; min-height: 80px; }
        .required { color: #f44336; }
        .form-actions { 
            text-align: right; 
            margin-top: 25px;
            padding-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.1);
        }
        .form-actions button { 
            padding: 12px 25px; 
            border: none; 
            border-radius: 8px; 
            cursor: pointer; 
            font-size: 14px; 
            margin-left: 10px;
            transition: all 0.2s ease;
        }
        .btn-cancel { background: #3a3a5c; color: #e0e0e0; }
        .btn-cancel:hover { background: #4a4a6c; }
        .btn-save { 
            background: linear-gradient(135deg, #4CAF50 0%, #388E3C 100%); 
            color: white; 
        }
        .btn-save:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(76,175,80,0.3); }
        
        /* 详情模态框样式 */
        .detail-section {
            margin-bottom: 20px;
            padding-bottom: 20px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .detail-section:last-child {
            border-bottom: none;
        }
        .detail-section h4 {
            color: #FF9800;
            margin-bottom: 15px;
            font-size: 16px;
        }
        .detail-row {
            display: flex;
            margin-bottom: 10px;
        }
        .detail-label {
            width: 100px;
            color: #888;
            flex-shrink: 0;
        }
        .detail-value {
            color: #e0e0e0;
            flex: 1;
        }
        .detail-stats {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 15px;
            margin-top: 15px;
        }
        .detail-stat-item {
            background: rgba(255,255,255,0.05);
            padding: 15px;
            border-radius: 8px;
            text-align: center;
        }
        .detail-stat-value {
            font-size: 24px;
            font-weight: bold;
            color: #FF9800;
        }
        .detail-stat-label {
            font-size: 12px;
            color: #888;
            margin-top: 5px;
        }
        
        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-grid { grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: repeat(auto-fill, minmax(120px, 1fr)); }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .btn-add { margin-left: 0; margin-top: 10px; }
        }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-truck"></i> 供应商管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>供应商管理</span>
            </div>
            <!-- V8 Tab导航 -->
            <div class="tab-nav" style="display:flex; gap:0; margin-top:15px; background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:10px; overflow:hidden;">
                <a href="?tab=suppliers" class="tab-link" style="flex:1;padding:12px 20px;text-align:center;color:<%=IIf(currentTab="suppliers","#fff","#888")%>;text-decoration:none;font-weight:500;border-bottom:3px solid <%=IIf(currentTab="suppliers","#FF9800","transparent")%>;background:<%=IIf(currentTab="suppliers","rgba(255,152,0,0.1)","transparent")%>;"><i class="fas fa-truck"></i> 供应商列表</a>
                <a href="?tab=contracts" class="tab-link" style="flex:1;padding:12px 20px;text-align:center;color:<%=IIf(currentTab="contracts","#fff","#888")%>;text-decoration:none;font-weight:500;border-bottom:3px solid <%=IIf(currentTab="contracts","#FF9800","transparent")%>;background:<%=IIf(currentTab="contracts","rgba(255,152,0,0.1)","transparent")%>;"><i class="fas fa-file-contract"></i> 合同管理</a>
                <a href="?tab=evals" class="tab-link" style="flex:1;padding:12px 20px;text-align:center;color:<%=IIf(currentTab="evals","#fff","#888")%>;text-decoration:none;font-weight:500;border-bottom:3px solid <%=IIf(currentTab="evals","#FF9800","transparent")%>;background:<%=IIf(currentTab="evals","rgba(255,152,0,0.1)","transparent")%>;"><i class="fas fa-star"></i> 供应商评估</a>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success" style="background: rgba(76,175,80,0.2); color: #4CAF50; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card total">
                <div class="stat-value"><%= totalSuppliers %></div>
                <div class="stat-label">总供应商数</div>
            </div>
            <div class="stat-card active">
                <div class="stat-value"><%= activeSuppliers %></div>
                <div class="stat-label">活跃供应商</div>
            </div>
            <div class="stat-card spice">
                <div class="stat-value"><%= spiceSuppliers %></div>
                <div class="stat-label">香料供应商</div>
            </div>
            <div class="stat-card package">
                <div class="stat-value"><%= packageSuppliers %></div>
                <div class="stat-label">包装供应商</div>
            </div>
            <div class="stat-card logistics">
                <div class="stat-value"><%= logisticsSuppliers %></div>
                <div class="stat-label">物流合作商</div>
            </div>
            <div class="stat-card printing">
                <div class="stat-value"><%= printSuppliers %></div>
                <div class="stat-label">印刷供应商</div>
            </div>
            <div class="stat-card spray">
                <div class="stat-value"><%= spraySuppliers %></div>
                <div class="stat-label">喷头供应商</div>
            </div>
        </div>
        
        <!-- 筛选栏 -->
        <div class="filter-bar">
            <form method="get" action="supplier_management.asp" style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center; flex: 1;">
                <select name="category">
                    <option value="">全部类别</option>
                    <option value="香料供应商" <% If filterCategory = "香料供应商" Then Response.Write "selected" %>>香料供应商</option>
                    <option value="包装供应商" <% If filterCategory = "包装供应商" Then Response.Write "selected" %>>包装供应商</option>
                    <option value="物流合作商" <% If filterCategory = "物流合作商" Then Response.Write "selected" %>>物流合作商</option>
                    <option value="印刷供应商" <% If filterCategory = "印刷供应商" Then Response.Write "selected" %>>印刷供应商</option>
                    <option value="喷头供应商" <% If filterCategory = "喷头供应商" Then Response.Write "selected" %>>喷头供应商</option>
                </select>
                <select name="status">
                    <option value="">全部状态</option>
                    <option value="active" <% If filterStatus = "active" Then Response.Write "selected" %>>启用</option>
                    <option value="inactive" <% If filterStatus = "inactive" Then Response.Write "selected" %>>禁用</option>
                </select>
                <input type="text" name="search" placeholder="搜索名称或联系人" value="<%= HTMLEncode(searchKeyword) %>">
                <button type="submit" class="btn-search"><i class="fas fa-search"></i> 搜索</button>
                <a href="supplier_management.asp" class="btn-reset" style="text-decoration: none; display: inline-block; padding: 10px 20px;"><i class="fas fa-undo"></i> 重置</a>
            </form>
            <% If canModify Then %>
            <button class="btn-add" onclick="openAddModal()"><i class="fas fa-plus"></i> 添加供应商</button>
            <% Else %>
            <button class="btn-add disabled" disabled title="无权限"><i class="fas fa-plus"></i> 添加供应商</button>
            <% End If %>
        </div>
        
        <!-- 供应商列表 -->
        <table class="supplier-table">
            <thead>
                <tr>
                    <th>ID</th>
                    <th>供应商名称</th>
                    <th>联系人</th>
                    <th>电话</th>
                    <th>供应类别</th>
                    <th>状态</th>
                    <th>创建时间</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If supplierCount > 0 Then %>
                <% For i = 0 To supplierCount - 1 %>
                <tr>
                    <td>#<%= supplierList(i, 0) %></td>
                    <td><%= HTMLEncode(supplierList(i, 1)) %></td>
                    <td><%= HTMLEncode(supplierList(i, 2)) %></td>
                    <td><%= HTMLEncode(supplierList(i, 3)) %></td>
                    <td>
                        <% If supplierList(i, 6) = "香料供应商" Then %>
                        <span class="category-badge category-spice"><i class="fas fa-leaf"></i> 香料供应商</span>
                        <% ElseIf supplierList(i, 6) = "包装供应商" Then %>
                        <span class="category-badge category-package"><i class="fas fa-box"></i> 包装供应商</span>
                        <% ElseIf supplierList(i, 6) = "物流合作商" Then %>
                        <span class="category-badge category-logistics"><i class="fas fa-shipping-fast"></i> 物流合作商</span>
                        <% ElseIf supplierList(i, 6) = "印刷供应商" Then %>
                        <span class="category-badge category-printing"><i class="fas fa-print"></i> 印刷供应商</span>
                        <% ElseIf supplierList(i, 6) = "喷头供应商" Then %>
                        <span class="category-badge category-spray"><i class="fas fa-spray-can"></i> 喷头供应商</span>
                        <% Else %>
                        <span class="category-badge category-other"><%= HTMLEncode(supplierList(i, 6)) %></span>
                        <% End If %>
                    </td>
                    <td>
                        <% If supplierList(i, 8) <> 0 Then %>
                        <span class="status-badge status-active"><i class="fas fa-check-circle"></i> 启用</span>
                        <% Else %>
                        <span class="status-badge status-inactive"><i class="fas fa-ban"></i> 禁用</span>
                        <% End If %>
                    </td>
                    <td><% If IsDate(supplierList(i, 9)) Then Response.Write FormatDateTime(supplierList(i, 9), 2) End If %></td>
                    <td>
                        <button class="btn-action btn-detail" onclick="location.href='supplier_management.asp?detailId=<%= supplierList(i, 0) %>'">
                            <i class="fas fa-eye"></i> 详情
                        </button>
                        <button class="btn-action btn-edit" onclick="openEditModal('<%= supplierList(i, 0) %>', '<%= SafeOutput(supplierList(i, 1)) %>', '<%= SafeOutput(supplierList(i, 2)) %>', '<%= SafeOutput(supplierList(i, 3)) %>', '<%= SafeOutput(supplierList(i, 4)) %>', '<%= SafeOutput(supplierList(i, 5)) %>', '<%= SafeOutput(supplierList(i, 6)) %>', '<%= SafeOutput(supplierList(i, 7)) %>')">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <% If canModify Then %>
                        <form method="post" action="supplier_management.asp" style="display: inline;">
                            <input type="hidden" name="action" value="toggle">
                            <input type="hidden" name="supplierId" value="<%= supplierList(i, 0) %>">
                            <button type="submit" class="btn-action btn-toggle" onclick="return confirm('确定要切换该供应商的状态吗？')">
                                <i class="fas fa-exchange-alt"></i> <%= IIf(supplierList(i, 8) <> 0, "禁用", "启用") %>
                            </button>
                        </form>
                        <% Else %>
                        <button class="btn-action btn-toggle disabled" disabled title="无权限">
                            <i class="fas fa-exchange-alt"></i> <%= IIf(supplierList(i, 8) <> 0, "禁用", "启用") %>
                        </button>
                        <% End If %>
                    </td>
                </tr>
                <% Next %>
                <% Else %>
                <tr>
                    <td colspan="8" style="text-align: center; padding: 40px; color: #666;">
                        <i class="fas fa-inbox" style="font-size: 48px; display: block; margin-bottom: 15px;"></i>
                        暂无供应商数据
                    </td>
                </tr>
                <% End If %>
            </tbody>
        </table>
    </div>
    
    <!-- V8 合同管理 Tab -->
    <% If currentTab = "contracts" Then %>
    <div class="data-section" style="margin-top:0;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <h3 style="color:#fff; margin:0;"><i class="fas fa-file-contract" style="color:#FF9800;"></i> 供应商合同列表</h3>
            <button class="btn-add" onclick="openContractModal()"><i class="fas fa-plus"></i> 添加合同</button>
        </div>
        <table class="supplier-table">
            <thead>
                <tr>
                    <th>合同编号</th>
                    <th>合同名称</th>
                    <th>供应商</th>
                    <th>有效期</th>
                    <th>金额</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsContracts Is Nothing Then
                %>
                <tr><td colspan="7" style="text-align:center; padding:40px; color:#666;"><i class="fas fa-inbox"></i> 暂无合同数据</td></tr>
                <%
                ElseIf rsContracts.EOF Then
                %>
                <tr><td colspan="7" style="text-align:center; padding:40px; color:#666;"><i class="fas fa-inbox"></i> 暂无合同数据</td></tr>
                <%
                Else
                    Do While Not rsContracts.EOF
                        Dim ctStatus : ctStatus = rsContracts("Status") & ""
                %>
                <tr>
                    <td><%= Server.HTMLEncode(rsContracts("ContractNo") & "") %></td>
                    <td><strong><%= Server.HTMLEncode(rsContracts("ContractName") & "") %></strong></td>
                    <td><%= Server.HTMLEncode(rsContracts("SupplierName") & "") %></td>
                    <td>
                        <% If IsDate(rsContracts("StartDate")) Then Response.Write FormatDateTime(rsContracts("StartDate"),2) %>
                        ~ <% If IsDate(rsContracts("EndDate")) Then Response.Write FormatDateTime(rsContracts("EndDate"),2) %>
                    </td>
                    <td>¥<%= FormatNumber(SafeNum(rsContracts("TotalAmount")), 2) %></td>
                    <td>
                        <span class="status-badge <% If ctStatus="Active" Then %>status-active<% Else %>status-inactive<% End If %>">
                            <%= IIf(ctStatus="Active", "生效中", "已失效") %>
                        </span>
                    </td>
                    <td>
                        <form method="post" style="display:inline;">
                            <input type="hidden" name="action" value="contract_toggle">
                            <input type="hidden" name="contract_id" value="<%= rsContracts("ContractID") %>">
                            <button type="submit" class="btn-action btn-toggle" style="font-size:11px;">
                                <%= IIf(ctStatus="Active", "停用", "启用") %>
                            </button>
                        </form>
                    </td>
                </tr>
                <%
                        rsContracts.MoveNext
                    Loop
                    rsContracts.Close : Set rsContracts = Nothing
                End If
                %>
            </tbody>
        </table>
    </div>
    <% End If %>
    
    <!-- V8 供应商评估 Tab -->
    <% If currentTab = "evals" Then %>
    <div class="data-section" style="margin-top:0;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <h3 style="color:#fff; margin:0;"><i class="fas fa-star" style="color:#FF9800;"></i> 供应商评估记录</h3>
            <button class="btn-add" onclick="openEvalModal()"><i class="fas fa-plus"></i> 新增评估</button>
        </div>
        <table class="supplier-table">
            <thead>
                <tr>
                    <th>供应商</th>
                    <th>质量</th>
                    <th>交付</th>
                    <th>价格</th>
                    <th>服务</th>
                    <th>综合</th>
                    <th>评级</th>
                    <th>评估人</th>
                    <th>日期</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsEvals Is Nothing Then
                %>
                <tr><td colspan="9" style="text-align:center; padding:40px; color:#666;"><i class="fas fa-inbox"></i> 暂无评估记录</td></tr>
                <%
                ElseIf rsEvals.EOF Then
                %>
                <tr><td colspan="9" style="text-align:center; padding:40px; color:#666;"><i class="fas fa-inbox"></i> 暂无评估记录</td></tr>
                <%
                Else
                    Do While Not rsEvals.EOF
                        Dim evScore : evScore = SafeNum(rsEvals("OverallScore"))
                        evRating = rsEvals("Rating") & ""
                %>
                <tr>
                    <td><strong><%= Server.HTMLEncode(rsEvals("SupplierName") & "") %></strong></td>
                    <td><%= rsEvals("QualityScore") %></td>
                    <td><%= rsEvals("DeliveryScore") %></td>
                    <td><%= rsEvals("PriceScore") %></td>
                    <td><%= rsEvals("ServiceScore") %></td>
                    <td><strong><%= evScore %></strong></td>
                    <td>
                        <span class="status-badge" style="<% If evRating="A" Then %>background:rgba(76,175,80,0.2);color:#4CAF50;<% ElseIf evRating="B" Then %>background:rgba(33,150,243,0.2);color:#2196F3;<% ElseIf evRating="D" Then %>background:rgba(244,67,54,0.2);color:#F44336;<% Else %>background:rgba(255,152,0,0.2);color:#FF9800;<% End If %>">
                            <%= evRating %>级
                        </span>
                    </td>
                    <td><%= Server.HTMLEncode(rsEvals("EvaluatedBy") & "") %></td>
                    <td><% If IsDate(rsEvals("EvaluationDate")) Then Response.Write FormatDateTime(rsEvals("EvaluationDate"),2) %></td>
                </tr>
                <%
                        rsEvals.MoveNext
                    Loop
                    rsEvals.Close : Set rsEvals = Nothing
                End If
                %>
            </tbody>
        </table>
    </div>
    <% End If %>
    
    <!-- 添加供应商模态框 -->
    <div id="addModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-plus-circle"></i> 添加供应商</h3>
                <button class="close-btn" onclick="closeModal('addModal')">&times;</button>
            </div>
            <form method="post" action="supplier_management.asp">
                <input type="hidden" name="action" value="add">
                
                <div class="form-group">
                    <label>供应商名称 <span class="required">*</span></label>
                    <input type="text" name="supplierName" required placeholder="请输入供应商名称">
                </div>
                
                <div class="form-group">
                    <label>联系人</label>
                    <input type="text" name="contactPerson" placeholder="请输入联系人姓名">
                </div>
                
                <div class="form-group">
                    <label>联系电话</label>
                    <input type="text" name="phone" placeholder="请输入联系电话">
                </div>
                
                <div class="form-group">
                    <label>电子邮箱</label>
                    <input type="email" name="email" placeholder="请输入电子邮箱">
                </div>
                
                <div class="form-group">
                    <label>地址</label>
                    <input type="text" name="address" placeholder="请输入供应商地址">
                </div>
                
                <div class="form-group">
                    <label>供应类别</label>
                    <select name="category">
                        <option value="">请选择类别</option>
                        <option value="香料供应商">香料供应商</option>
                        <option value="包装供应商">包装供应商</option>
                        <option value="物流合作商">物流合作商</option>
                        <option value="印刷供应商">印刷供应商</option>
                        <option value="喷头供应商">喷头供应商</option>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>备注</label>
                    <textarea name="notes" placeholder="请输入备注信息..."></textarea>
                </div>
                
                <div class="form-actions">
                    <button type="button" class="btn-cancel" onclick="closeModal('addModal')">取消</button>
                    <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 编辑供应商模态框 -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-edit"></i> 编辑供应商</h3>
                <button class="close-btn" onclick="closeModal('editModal')">&times;</button>
            </div>
            <form method="post" action="supplier_management.asp">
                <input type="hidden" name="action" value="edit">
                <input type="hidden" name="supplierId" id="editSupplierId">
                
                <div class="form-group">
                    <label>供应商名称 <span class="required">*</span></label>
                    <input type="text" name="supplierName" id="editSupplierName" required placeholder="请输入供应商名称">
                </div>
                
                <div class="form-group">
                    <label>联系人</label>
                    <input type="text" name="contactPerson" id="editContactPerson" placeholder="请输入联系人姓名">
                </div>
                
                <div class="form-group">
                    <label>联系电话</label>
                    <input type="text" name="phone" id="editPhone" placeholder="请输入联系电话">
                </div>
                
                <div class="form-group">
                    <label>电子邮箱</label>
                    <input type="email" name="email" id="editEmail" placeholder="请输入电子邮箱">
                </div>
                
                <div class="form-group">
                    <label>地址</label>
                    <input type="text" name="address" id="editAddress" placeholder="请输入供应商地址">
                </div>
                
                <div class="form-group">
                    <label>供应类别</label>
                    <select name="category" id="editCategory">
                        <option value="">请选择类别</option>
                        <option value="香料供应商">香料供应商</option>
                        <option value="包装供应商">包装供应商</option>
                        <option value="物流合作商">物流合作商</option>
                        <option value="印刷供应商">印刷供应商</option>
                        <option value="喷头供应商">喷头供应商</option>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>备注</label>
                    <textarea name="notes" id="editNotes" placeholder="请输入备注信息..."></textarea>
                </div>
                
                <div class="form-actions">
                    <button type="button" class="btn-cancel" onclick="closeModal('editModal')">取消</button>
                    <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 供应商详情模态框 -->
    <div id="detailModal" class="modal" style="<%= IIf(detailExists, "display: block;", "display: none;") %>">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-info-circle"></i> 供应商详情</h3>
                <button class="close-btn" onclick="closeDetailModal()">&times;</button>
            </div>
            <% If detailExists Then %>
            <div class="detail-section">
                <h4><i class="fas fa-building"></i> 基本信息</h4>
                <div class="detail-row">
                    <div class="detail-label">供应商名称</div>
                    <div class="detail-value"><%= HTMLEncode(CStr(rsDetail("SupplierName"))) %></div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">联系人</div>
                    <div class="detail-value"><%= HTMLEncode(CStr(rsDetail("ContactPerson"))) %></div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">联系电话</div>
                    <div class="detail-value"><%= HTMLEncode(CStr(rsDetail("Phone"))) %></div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">电子邮箱</div>
                    <div class="detail-value"><%= HTMLEncode(CStr(rsDetail("Email"))) %></div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">地址</div>
                    <div class="detail-value"><%= HTMLEncode(CStr(rsDetail("Address"))) %></div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">供应类别</div>
                    <div class="detail-value">
                        <% 
                        Dim detailCategory
                        detailCategory = CStr(rsDetail("Category"))
                        If detailCategory = "香料供应商" Then
                            Response.Write "<span class='category-badge category-spice'><i class='fas fa-leaf'></i> 香料供应商</span>"
                        ElseIf detailCategory = "包装供应商" Then
                            Response.Write "<span class='category-badge category-package'><i class='fas fa-box'></i> 包装供应商</span>"
                        ElseIf detailCategory = "物流合作商" Then
                            Response.Write "<span class='category-badge category-logistics'><i class='fas fa-shipping-fast'></i> 物流合作商</span>"
                        ElseIf detailCategory = "印刷供应商" Then
                            Response.Write "<span class='category-badge category-printing'><i class='fas fa-print'></i> 印刷供应商</span>"
                        ElseIf detailCategory = "喷头供应商" Then
                            Response.Write "<span class='category-badge category-spray'><i class='fas fa-spray-can'></i> 喷头供应商</span>"
                        Else
                            Response.Write HTMLEncode(detailCategory)
                        End If
                        %>
                    </div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">状态</div>
                    <div class="detail-value">
                        <% If rsDetail("IsActive") <> 0 Then %>
                        <span class="status-badge status-active"><i class="fas fa-check-circle"></i> 启用</span>
                        <% Else %>
                        <span class="status-badge status-inactive"><i class="fas fa-ban"></i> 禁用</span>
                        <% End If %>
                    </div>
                </div>
            </div>
            
            <div class="detail-section">
                <h4><i class="fas fa-chart-bar"></i> 业务统计</h4>
                <div class="detail-stats">
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= GetSupplierOrderCount(rsDetail("SupplierID")) %></div>
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= supplierEvalCount %></div>
            
            <div class="detail-section">
                <h4><i class="fas fa-history"></i> 合作历史</h4>
                <%
                If Not rsSupplierOrders Is Nothing Then
                    If Not rsSupplierOrders.EOF Then
                %>
                <table class="data-table" style="font-size:12px;">
                    <thead>
                        <tr>
                            <th>订单号</th>
                            <th>日期</th>
                            <th style="text-align:right;">金额</th>
                            <th>状态</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        Do While Not rsSupplierOrders.EOF
                        %>
                        <tr>
                            <td><a href="purchase_orders.asp?view=<%= rsSupplierOrders("PurchaseID") %>" style="color:#FF9800;"><%= Server.HTMLEncode(rsSupplierOrders("PurchaseNo") & "") %></a></td>
                            <td><% If IsDate(rsSupplierOrders("OrderDate")) Then Response.Write FormatDateTime(rsSupplierOrders("OrderDate"), 2) %></td>
                            <td style="text-align:right;">¥<%= FormatNumber(SafeNum(rsSupplierOrders("TotalAmount")), 2) %></td>
                            <td><span class="status-badge status-<%= LCase(rsSupplierOrders("Status") & "") %>"><%= rsSupplierOrders("Status") & "" %></span></td>
                        </tr>
                        <%
                            rsSupplierOrders.MoveNext
                        Loop
                        rsSupplierOrders.Close : Set rsSupplierOrders = Nothing
                        %>
                    </tbody>
                </table>
                <%
                    Else
                        rsSupplierOrders.Close : Set rsSupplierOrders = Nothing
                %>
                <div style="color:#666;padding:10px 0;">暂无采购记录</div>
                <%
                    End If
                Else
                %>
                <div style="color:#666;padding:10px 0;">暂无采购记录</div>
                <% End If %>
            </div>
                        <div class="detail-stat-label">评估次数</div>
                    </div>
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= FormatNumber(supplierRatingAvg, 1) %></div>
                        <div class="detail-stat-label">综合评分</div>
                    </div>
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= supplierActiveContracts %></div>
                        <div class="detail-stat-label">活跃合同</div>
                    </div>
                        <div class="detail-stat-label">关联采购单</div>
                    </div>
                    <div class="detail-stat-item">
                        <div class="detail-stat-value"><%= GetSupplierPriceCount(rsDetail("SupplierID")) %></div>
                        <div class="detail-stat-label">有效报价</div>
                    </div>
                </div>
            </div>
            
            <% If Not IsNull(rsDetail("Notes")) And CStr(rsDetail("Notes")) <> "" Then %>
            <div class="detail-section">
                <h4><i class="fas fa-sticky-note"></i> 备注</h4>
                <div style="color: #e0e0e0; line-height: 1.6;">
                    <%= HTMLEncode(CStr(rsDetail("Notes"))) %>
                </div>
            </div>
            <% End If %>
            <% End If %>
        </div>
    </div>
    
    <!-- V8 合同添加模态框 -->
    <div id="contractModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-file-contract"></i> 添加合同</h3>
                <button class="close-btn" onclick="closeModal('contractModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="contract_add">
                <div class="form-group">
                    <label>供应商 <span class="required">*</span></label>
                    <select name="supplier_id" required>
                        <option value="">请选择供应商</option>
                        <%
                        Dim rsCS : Set rsCS = conn.Execute("SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive=1 ORDER BY SupplierName")
                        If Not rsCS Is Nothing Then
                            Do While Not rsCS.EOF
                        %>
                        <option value="<%= rsCS("SupplierID") %>"><%= Server.HTMLEncode(rsCS("SupplierName") & "") %></option>
                        <%
                                rsCS.MoveNext
                            Loop
                            rsCS.Close : Set rsCS = Nothing
                        End If
                        %>
                    </select>
                </div>
                <div class="form-group">
                    <label>合同编号</label>
                    <input type="text" name="contract_no" placeholder="CT-">
                </div>
                <div class="form-group">
                    <label>合同名称 <span class="required">*</span></label>
                    <input type="text" name="contract_name" required>
                </div>
                <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:15px;">
                    <div class="form-group">
                        <label>开始日期</label>
                        <input type="date" name="start_date">
                    </div>
                    <div class="form-group">
                        <label>结束日期</label>
                        <input type="date" name="end_date">
                    </div>
                </div>
                <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:15px;">
                    <div class="form-group">
                        <label>合同金额</label>
                        <input type="number" name="total_amount" step="0.01" value="0">
                    </div>
                    <div class="form-group">
                        <label>付款条件</label>
                        <input type="text" name="payment_terms" placeholder="如：货到30天付款">
                    </div>
                </div>
                <div class="form-group">
                    <label>条款摘要</label>
                    <textarea name="terms_summary" rows="3"></textarea>
                </div>
                <div class="form-actions" style="text-align:right;">
                    <button type="button" class="btn-cancel" onclick="closeModal('contractModal')">取消</button>
                    <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- V8 供应商评估模态框 -->
    <div id="evalModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-star"></i> 新增供应商评估</h3>
                <button class="close-btn" onclick="closeModal('evalModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="eval_add">
                <div class="form-group">
                    <label>供应商 <span class="required">*</span></label>
                    <select name="supplier_id" required>
                        <option value="">请选择供应商</option>
                        <%
                        Dim rsES : Set rsES = conn.Execute("SELECT SupplierID, SupplierName FROM Suppliers WHERE IsActive=1 ORDER BY SupplierName")
                        If Not rsES Is Nothing Then
                            Do While Not rsES.EOF
                        %>
                        <option value="<%= rsES("SupplierID") %>"><%= Server.HTMLEncode(rsES("SupplierName") & "") %></option>
                        <%
                                rsES.MoveNext
                            Loop
                            rsES.Close : Set rsES = Nothing
                        End If
                        %>
                    </select>
                </div>
                <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:15px;">
                    <div class="form-group">
                        <label>质量评分 (0-100)</label>
                        <input type="number" name="quality_score" min="0" max="100" value="80">
                    </div>
                    <div class="form-group">
                        <label>交付评分 (0-100)</label>
                        <input type="number" name="delivery_score" min="0" max="100" value="80">
                    </div>
                </div>
                <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:15px;">
                    <div class="form-group">
                        <label>价格评分 (0-100)</label>
                        <input type="number" name="price_score" min="0" max="100" value="80">
                    </div>
                    <div class="form-group">
                        <label>服务评分 (0-100)</label>
                        <input type="number" name="service_score" min="0" max="100" value="80">
                    </div>
                </div>
                <div class="form-group">
                    <label>评价备注</label>
                    <textarea name="comments" rows="3"></textarea>
                </div>
                <div class="form-actions" style="text-align:right;">
                    <button type="button" class="btn-cancel" onclick="closeModal('evalModal')">取消</button>
                    <button type="submit" class="btn-save"><i class="fas fa-save"></i> 保存</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        function openAddModal() {
            document.getElementById('addModal').style.display = 'block';
        }
        
        function openEditModal(id, name, contact, phone, email, address, category, notes) {
            document.getElementById('editSupplierId').value = id;
            document.getElementById('editSupplierName').value = name;
            document.getElementById('editContactPerson').value = contact;
            document.getElementById('editPhone').value = phone;
            document.getElementById('editEmail').value = email;
            document.getElementById('editAddress').value = address;
            document.getElementById('editCategory').value = category;
            document.getElementById('editNotes').value = notes;
            document.getElementById('editModal').style.display = 'block';
        }
        
        function closeModal(modalId) {
            document.getElementById(modalId).style.display = 'none';
        }
        
        function closeDetailModal() {
            document.getElementById('detailModal').style.display = 'none';
            if (window.history.replaceState) {
                window.history.replaceState({}, document.title, 'supplier_management.asp');
            }
        }
        
        function openContractModal() {
            document.getElementById('contractModal').style.display = 'block';
        }
        
        function openEvalModal() {
            document.getElementById('evalModal').style.display = 'block';
        }
        
        window.onclick = function(event) {
            if (event.target.classList.contains('modal')) {
                event.target.style.display = 'none';
                if (event.target.id === 'detailModal') {
                    if (window.history.replaceState) {
                        window.history.replaceState({}, document.title, 'supplier_management.asp');
                    }
                }
            }
        }
    </script>
</body>
</html>
<%
If detailExists Then
    rsDetail.Close
    Set rsDetail = Nothing
End If
Call CloseConnection()
%>
