<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/cost_engine.asp"-->
<%
' 安全数值转换函数
Function SafeNum(val)
    If IsNull(val) Or IsEmpty(val) Or val = "" Then
        SafeNum = 0
    Else
        On Error Resume Next
        SafeNum = CDbl(val)
        If Err.Number <> 0 Then
            SafeNum = 0
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

Call OpenConnection()

' ============================================
' 权限检查
' ============================================
Dim canEdit
canEdit = (Session("AdminRoleCode") = "FIN_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN")

' ============================================
' 获取当前Tab
' ============================================
Dim currentTab, vProductId, vProductName, vLastCost, vCurrentCost, vVariance, vVarianceRate
Dim isWarning, statusClass, statusText, noteKey2, varianceNote
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "products"

' 成本传导链标签页：预加载所有成本数据（避免N+1查询）
If currentTab = "chain" Then Call CE_PreloadAllCostData()

' ============================================
' 获取分页参数
' ============================================
Dim page, pageSize
page = CInt(SafeNum(IIf(Request.QueryString("page") = "", 1, Request.QueryString("page"))))
If page < 1 Then page = 1
pageSize = 20

' ============================================
' 处理表单提交
' ============================================
Dim action, msg, errMsg
action = Request.Form("action")
msg = ""
errMsg = ""

If action = "save_cost_method" AND canEdit Then
    ' 保存成本计价方式
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        Dim costMethod
        costMethod = Request.Form("costMethod")
        If costMethod = "FIFO" OR costMethod = "WEIGHTED" Then
            Dim checkSQL, updateSQL, insertSQL
            checkSQL = "SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = 'CostCalculationMethod'"
            If GetScalar(checkSQL) > 0 Then
                updateSQL = "UPDATE SiteSettings SET SettingValue = '" & SafeSQL(costMethod) & "' WHERE SettingKey = 'CostCalculationMethod'"
                ExecuteNonQuery updateSQL
            Else
                insertSQL = "INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('CostCalculationMethod', '" & SafeSQL(costMethod) & "')"
                ExecuteNonQuery insertSQL
            End If
            Call LogAdminAction("修改成本计价方式", "finance", "SiteSettings", "", "设置为: " & costMethod)
            msg = "计价方式保存成功"
        End If
    End If
End If

If action = "save_product_cost" AND canEdit Then
    ' 保存商品成本
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        Dim productId, costItemCount, i
        productId = CInt(SafeNum(Request.Form("productId")))
        costItemCount = CInt(SafeNum(Request.Form("costItemCount")))
        
        Dim totalBOMCost, totalPurchaseCost, totalPackagingCost, totalOtherCost
        ' 循环变量声明移到循环外（VBScript限制）
        Dim costType, costName, unitCost, quantity, totalCost, effectiveDate, insertCostSQL
        totalBOMCost = 0
        totalPurchaseCost = 0
        totalPackagingCost = 0
        totalOtherCost = 0
        
        ' 删除该商品旧的成本记录
        Dim deleteSQL
        deleteSQL = "DELETE FROM ProductCosts WHERE ProductID = " & productId
        ExecuteNonQuery deleteSQL
        
        ' 插入新的成本记录
        For i = 1 To costItemCount
            costType = Request.Form("costType_" & i)
            costName = Request.Form("costName_" & i)
            unitCost = SafeNum(Request.Form("unitCost_" & i))
            quantity = SafeNum(Request.Form("quantity_" & i))
            If quantity = 0 Then quantity = 1
            totalCost = unitCost * quantity
            effectiveDate = Request.Form("effectiveDate_" & i)
            If effectiveDate = "" Then effectiveDate = CStr(Date())
            
            If costName <> "" Then
                insertCostSQL = "INSERT INTO ProductCosts (ProductID, CostType, CostName, UnitCost, Quantity, TotalCost, EffectiveDate, CreatedBy, CreatedAt) VALUES (" & _
                    productId & ", '" & SafeSQL(costType) & "', '" & SafeSQL(costName) & "', " & unitCost & ", " & quantity & ", " & totalCost & ", #" & effectiveDate & "#, '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
                ExecuteNonQuery insertCostSQL
                
                ' 累加各类成本
                Select Case costType
                    Case "BOM": totalBOMCost = totalBOMCost + totalCost
                    Case "Purchase": totalPurchaseCost = totalPurchaseCost + totalCost
                    Case "Packaging": totalPackagingCost = totalPackagingCost + totalCost
                    Case "Other": totalOtherCost = totalOtherCost + totalCost
                End Select
            End If
        Next
        
        ' 更新Products表的汇总成本字段
        Dim totalUnitCost
        totalUnitCost = totalBOMCost + totalPurchaseCost + totalPackagingCost + totalOtherCost
        Dim updateProductSQL
        updateProductSQL = "UPDATE Products SET BOMCost = " & totalBOMCost & ", UnitCost = " & totalUnitCost & " WHERE ProductID = " & productId
        ExecuteNonQuery updateProductSQL
        
        Call LogAdminAction("更新商品成本", "finance", "ProductCosts", CStr(productId), "更新商品成本信息")
        msg = "成本信息保存成功"
    End If
End If

If action = "auto_calc_all" AND canEdit Then
    ' 批量自动计算所有产品的成本
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        On Error Resume Next
        Err.Clear
        Call CE_UpdateAllProductCosts()
        If Err.Number = 0 Then
            msg = "所有产品BOM/单位成本已自动更新完成。已更新 " & Session("CE_LastUpdateCount") & " 个产品。"
            Call LogAdminAction("批量自动计算成本", "finance", "Products", "", "触发批量成本传导")
        Else
            errMsg = "成本计算过程中出现错误: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End If

If action = "auto_calc_orders" AND canEdit Then
    ' 批量更新所有订单的成本和利润
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        On Error Resume Next
        Err.Clear
        Call CE_UpdateAllOrderCosts()
        If Err.Number = 0 Then
            msg = "所有已支付订单的成本和利润已自动更新完成。"
            Call LogAdminAction("批量更新订单成本", "finance", "Orders", "", "触发订单成本传导")
        Else
            errMsg = "订单成本更新过程中出现错误: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End If

If action = "save_variance_note" AND canEdit Then
    ' 保存异动归因备注
    If Not ValidateCSRFToken() Then
        errMsg = "安全验证失败"
    Else
        Dim noteProductId
        noteProductId = CInt(SafeNum(Request.Form("productId")))
        varianceNote = Request.Form("varianceNote")
        
        ' 将备注存入SiteSettings，Key格式：CostVarianceNote_ProductID_年月
        Dim currentYM, noteKey
        currentYM = Year(Date()) & Right("0" & Month(Date()), 2)
        noteKey = "CostVarianceNote_" & noteProductId & "_" & currentYM
        
        Dim checkNoteSQL
        checkNoteSQL = "SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = '" & SafeSQL(noteKey) & "'"
        If GetScalar(checkNoteSQL) > 0 Then
            Dim updateNoteSQL
            updateNoteSQL = "UPDATE SiteSettings SET SettingValue = '" & SafeSQL(varianceNote) & "' WHERE SettingKey = '" & SafeSQL(noteKey) & "'"
            ExecuteNonQuery updateNoteSQL
        Else
            Dim insertNoteSQL
            insertNoteSQL = "INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('" & SafeSQL(noteKey) & "', '" & SafeSQL(varianceNote) & "')"
            ExecuteNonQuery insertNoteSQL
        End If
        
        msg = "归因备注保存成功"
    End If
End If

' ============================================
' 获取当前计价方式
' ============================================
Dim currentCostMethod
currentCostMethod = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'CostCalculationMethod'")
If currentCostMethod = "0" OR currentCostMethod = "" Then currentCostMethod = "FIFO"

' ============================================
' 获取商品总数（用于分页）
' ============================================
Dim totalProducts, totalPages
totalProducts = CInt(SafeNum(GetScalar("SELECT COUNT(*) FROM Products")))
totalPages = Int((totalProducts + pageSize - 1) / pageSize)
If totalPages < 1 Then totalPages = 1
If page > totalPages Then page = totalPages

Dim offset
offset = (page - 1) * pageSize

Call LogAdminAction("查看成本管理", "finance", "", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>成本管理 - 财务管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 暗色主题基础样式 */
        body { background: #1a1a2e; color: #e0e0e0; }
        .main-content { padding: 30px; margin-left: 260px; }
        
        /* 页面标题 */
        .page-header { margin-bottom: 25px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 12px; }
        .breadcrumb { color: #888; font-size: 14px; margin-top: 8px; }
        .breadcrumb a { color: #00bcd4; text-decoration: none; }
        
        /* 计价方式选择区 */
        .cost-method-banner { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); 
            border: 1px solid rgba(255,255,255,0.06); border-radius: 12px; padding: 20px; margin-bottom: 25px;
            display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 20px;
        }
        .cost-method-banner .method-info { display: flex; align-items: center; gap: 15px; }
        .cost-method-banner .method-info i { font-size: 28px; color: #ffa726; }
        .cost-method-banner .method-text h4 { margin: 0; color: #e0e0e0; font-size: 16px; }
        .cost-method-banner .method-text p { margin: 5px 0 0; color: #888; font-size: 13px; }
        .cost-method-selector { display: flex; align-items: center; gap: 15px; }
        .cost-method-selector label { color: #b0b0b0; font-size: 14px; }
        .cost-method-selector select { 
            padding: 10px 15px; border: 2px solid #3a3a4a; border-radius: 8px; 
            background: #1a1a2e; color: #e0e0e0; font-size: 14px; min-width: 180px;
        }
        .cost-method-selector select:focus { border-color: #00bcd4; outline: none; }
        
        /* Tab导航 */
        .tab-nav { display: flex; border-bottom: 2px solid #3a3a4a; margin-bottom: 25px; }
        .tab-nav a { 
            padding: 15px 30px; color: #888; text-decoration: none; font-size: 15px;
            border-bottom: 3px solid transparent; transition: all 0.3s;
            display: flex; align-items: center; gap: 8px;
        }
        .tab-nav a:hover { color: #e0e0e0; }
        .tab-nav a.active { color: #00bcd4; border-bottom-color: #00bcd4; }
        
        /* 内容卡片 */
        .content-card { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; padding: 25px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.06);
        }
        
        /* 表格样式 */
        .data-table { width: 100%; border-collapse: collapse; }
        .data-table th, .data-table td { 
            padding: 15px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.06); 
        }
        .data-table th { 
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%); color: white; font-weight: 600; font-size: 13px;
        }
        .data-table td { color: #e0e0e0; font-size: 14px; }
        .data-table tr:hover { background: rgba(255,255,255,0.05); }
        .data-table tr.warning-row { background: rgba(244, 67, 54, 0.15) !important; }
        .data-table tr.warning-row:hover { background: rgba(244, 67, 54, 0.25) !important; }
        
        /* 成本数值样式 */
        .cost-value { font-family: 'Courier New', monospace; font-weight: 600; }
        .cost-bom { color: #4CAF50; }
        .cost-purchase { color: #2196F3; }
        .cost-packaging { color: #FF9800; }
        .cost-total { color: #e0e0e0; font-size: 16px; }
        
        /* 变动指示器 */
        .variance-up { color: #f44336; }
        .variance-down { color: #4CAF50; }
        .variance-neutral { color: #888; }
        
        /* 状态标签 */
        .status-badge { 
            display: inline-block; padding: 4px 12px; border-radius: 12px; 
            font-size: 12px; font-weight: 500; 
        }
        .status-normal { background: #1b5e20; color: #81c784; }
        .status-warning { background: #5e1b1b; color: #e57373; }
        

        /* 分页 */
        .pagination { 
            display: flex; justify-content: center; align-items: center; 
            gap: 10px; margin-top: 25px; 
        }
        .pagination a, .pagination span { 
            padding: 8px 14px; border-radius: 6px; text-decoration: none; 
            font-size: 14px; min-width: 40px; text-align: center;
        }
        .pagination a { background: #3a3a4a; color: #e0e0e0; }
        .pagination a:hover { background: #4a4a5a; }
        .pagination .current { background: #00bcd4; color: white; }
        .pagination .disabled { background: #2a2a3a; color: #666; cursor: not-allowed; }
        
        /* 模态框 */
        .modal { 
            display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; 
            background: rgba(0,0,0,0.7); z-index: 1000; align-items: center; justify-content: center;
        }
        .modal.active { display: flex; }
        .modal-content { 
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; width: 90%; max-width: 700px; 
            max-height: 90vh; overflow: hidden; border: 1px solid rgba(255,255,255,0.06);
        }
        .modal-header { 
            padding: 20px 25px; border-bottom: 1px solid rgba(255,255,255,0.06); 
            display: flex; justify-content: space-between; align-items: center;
        }
        .modal-header h3 { margin: 0; color: #e0e0e0; font-size: 18px; }
        .modal-close { 
            background: none; border: none; color: #888; font-size: 24px; cursor: pointer;
        }
        .modal-close:hover { color: #e0e0e0; }
        .modal-body { padding: 25px; max-height: 60vh; overflow-y: auto; }
        .modal-footer { 
            padding: 20px 25px; border-top: 1px solid rgba(255,255,255,0.06); 
            display: flex; justify-content: flex-end; gap: 10px;
        }
        
        /* 表单样式 */
        .form-group { margin-bottom: 20px; }
        .form-group label { 
            display: block; margin-bottom: 8px; color: #b0b0b0; font-size: 14px; font-weight: 500;
        }
        .form-group input, .form-group select { 
            width: 100%; padding: 12px 15px; border: 2px solid #3a3a4a; border-radius: 8px; 
            font-size: 14px; background: #1a1a2e; color: #e0e0e0; box-sizing: border-box;
        }
        .form-group input:focus, .form-group select:focus { border-color: #00bcd4; outline: none; }
        .form-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; }
        
        /* 成本项卡片 */
        .cost-item { 
            background: #1a1a2e; border: 1px solid #3a3a4a; border-radius: 8px; 
            padding: 15px; margin-bottom: 15px; position: relative;
        }
        .cost-item-header { 
            display: flex; justify-content: space-between; align-items: center; 
            margin-bottom: 15px;
        }
        .cost-item-title { font-weight: 600; color: #00bcd4; }
        .cost-item-remove { 
            background: none; border: none; color: #f44336; cursor: pointer; font-size: 18px;
        }
        
        /* 提示信息 */
        .alert { 
            padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; 
            display: flex; align-items: center; gap: 10px;
        }
        .alert-success { background: rgba(76, 175, 80, 0.15); color: #81c784; border: 1px solid rgba(76, 175, 80, 0.3); }
        .alert-error { background: rgba(244, 67, 54, 0.15); color: #e57373; border: 1px solid rgba(244, 67, 54, 0.3); }
        .alert-warning { background: rgba(255, 152, 0, 0.15); color: #ffb74d; border: 1px solid rgba(255, 152, 0, 0.3); }
        
        /* 只读遮罩 */
        .readonly-mask { position: relative; }
        .readonly-mask::after { 
            content: "只读权限"; position: absolute; top: 0; left: 0; right: 0; bottom: 0; 
            background: rgba(26,26,46,0.85); display: flex; align-items: center; justify-content: center; 
            font-size: 16px; color: #888; border-radius: 12px;
        }
        
        /* 历史记录时间线 */
        .timeline { position: relative; padding-left: 30px; }
        .timeline::before { 
            content: ""; position: absolute; left: 8px; top: 0; bottom: 0; 
            width: 2px; background: #3a3a4a;
        }
        .timeline-item { position: relative; margin-bottom: 20px; }
        .timeline-item::before { 
            content: ""; position: absolute; left: -26px; top: 4px; 
            width: 12px; height: 12px; border-radius: 50%; background: #00bcd4;
        }
        .timeline-date { font-size: 12px; color: #888; margin-bottom: 5px; }
        .timeline-content { background: #1a1a2e; padding: 15px; border-radius: 8px; }
        .timeline-title { font-weight: 600; color: #e0e0e0; margin-bottom: 8px; }
        .timeline-detail { font-size: 13px; color: #b0b0b0; }
        
        /* 归因输入框 */
        .variance-note-input { 
            width: 100%; padding: 8px 12px; border: 1px solid #3a3a4a; border-radius: 6px;
            background: #1a1a2e; color: #e0e0e0; font-size: 13px;
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-coins"></i> 成本管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">财务中心</a> / <span>成本管理</span>
            </div>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-success"><i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(msg) %></div>
        <% End If %>
        
        <% If errMsg <> "" Then %>
        <div class="alert alert-error"><i class="fas fa-times-circle"></i> <%= Server.HTMLEncode(errMsg) %></div>
        <% End If %>
        
        <!-- 计价方式选择 -->
        <div class="cost-method-banner">
            <div class="method-info">
                <i class="fas fa-calculator"></i>
                <div class="method-text">
                    <h4>出库成本计价规则</h4>
                    <p>选择系统计算商品出库成本的方法</p>
                </div>
            </div>
            <form method="post" action="cost_management.asp?tab=<%= currentTab %>" class="cost-method-selector">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="save_cost_method">
                <label>计价方式：</label>
                <select name="costMethod" <%= IIf(canEdit, "", "disabled") %>>
                    <option value="FIFO" <%= IIf(currentCostMethod="FIFO", "selected", "") %>>先进先出 (FIFO)</option>
                    <option value="WEIGHTED" <%= IIf(currentCostMethod="WEIGHTED", "selected", "") %>>移动加权平均</option>
                </select>
                <% If canEdit Then %>
                <button type="submit" class="btn btn-primary"><i class="fas fa-save"></i> 保存</button>
                <% End If %>
            </form>
        </div>
        
        <!-- Tab导航 -->
        <div class="tab-nav">
            <a href="?tab=products" class="<%= IIf(currentTab="products", "active", "") %>">
                <i class="fas fa-box"></i> 商品成本维护
            </a>
            <a href="?tab=variance" class="<%= IIf(currentTab="variance", "active", "") %>">
                <i class="fas fa-chart-line"></i> 成本异动监控
            </a>
            <a href="?tab=history" class="<%= IIf(currentTab="history", "active", "") %>">
                <i class="fas fa-history"></i> 成本变更历史
            </a>
            <a href="?tab=chain" class="<%= IIf(currentTab="chain", "active", "") %>">
                <i class="fas fa-link"></i> 成本传导链
            </a>
        </div>
        
        <!-- Tab内容 -->
        <div class="content-card">
            <% If currentTab = "products" Then %>
                <!-- Tab 1: 商品成本维护 -->
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>商品名称</th>
                            <th>BOM成本</th>
                            <th>采购价</th>
                            <th>包材成本</th>
                            <th>其他成本</th>
                            <th>单位总成本</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        Dim rsProducts, productSQL
                        productSQL = "SELECT TOP " & pageSize & " p.ProductID, p.ProductName, p.BOMCost, p.UnitCost, " & _
                            "(SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID = p.ProductID AND CostType = 'Purchase') AS PurchaseCost, " & _
                            "(SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID = p.ProductID AND CostType = 'Packaging') AS PackagingCost, " & _
                            "(SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID = p.ProductID AND CostType = 'Other') AS OtherCost " & _
                            "FROM Products p " & _
                            "WHERE p.ProductID NOT IN (SELECT TOP " & offset & " ProductID FROM Products ORDER BY ProductID) " & _
                            "ORDER BY p.ProductID"
                        
                        Dim pId, pName, pBOMCost, pUnitCost, pPurchaseCost, pPackagingCost, pOtherCost
                        Set rsProducts = ExecuteQuery(productSQL)
                        If Not rsProducts Is Nothing Then
                            Do While Not rsProducts.EOF
                                pId = rsProducts("ProductID")
                                pName = rsProducts("ProductName")
                                pBOMCost = SafeNum(rsProducts("BOMCost"))
                                pUnitCost = SafeNum(rsProducts("UnitCost"))
                                pPurchaseCost = SafeNum(rsProducts("PurchaseCost"))
                                pPackagingCost = SafeNum(rsProducts("PackagingCost"))
                                pOtherCost = SafeNum(rsProducts("OtherCost"))
                        %>
                        <tr>
                            <td><strong><%= Server.HTMLEncode(pName) %></strong></td>
                            <td class="cost-value cost-bom">¥<%= FormatNumber(pBOMCost, 2) %></td>
                            <td class="cost-value cost-purchase">¥<%= FormatNumber(pPurchaseCost, 2) %></td>
                            <td class="cost-value cost-packaging">¥<%= FormatNumber(pPackagingCost, 2) %></td>
                            <td class="cost-value">¥<%= FormatNumber(pOtherCost, 2) %></td>
                            <td class="cost-value cost-total">¥<%= FormatNumber(pUnitCost, 2) %></td>
                            <td>
                                <button class="btn btn-primary btn-sm" onclick="openCostModal(<%= pId %>, '<%= SafeOutput(pName) %>')">
                                    <i class="fas fa-edit"></i> 编辑成本
                                </button>
                            </td>
                        </tr>
                        <%
                                rsProducts.MoveNext
                            Loop
                            rsProducts.Close
                            Set rsProducts = Nothing
                        End If
                        %>
                    </tbody>
                </table>
                
                <!-- 分页 -->
                <div class="pagination">
                    <% If page > 1 Then %>
                        <a href="?tab=products&page=<%= page-1 %>"><i class="fas fa-chevron-left"></i></a>
                    <% Else %>
                        <span class="disabled"><i class="fas fa-chevron-left"></i></span>
                    <% End If %>
                    
                    <% 
                    Dim pStart, pEnd
                    pStart = IIf(page - 2 < 1, 1, page - 2)
                    pEnd = IIf(pStart + 4 > totalPages, totalPages, pStart + 4)
                    If pEnd - pStart < 4 Then pStart = IIf(pEnd - 4 < 1, 1, pEnd - 4)
                    
                    For i = pStart To pEnd 
                    %>
                        <% If i = page Then %>
                            <span class="current"><%= i %></span>
                        <% Else %>
                            <a href="?tab=products&page=<%= i %>"><%= i %></a>
                        <% End If %>
                    <% Next %>
                    
                    <% If page < totalPages Then %>
                        <a href="?tab=products&page=<%= page+1 %>"><i class="fas fa-chevron-right"></i></a>
                    <% Else %>
                        <span class="disabled"><i class="fas fa-chevron-right"></i></span>
                    <% End If %>
                </div>
                
            <% ElseIf currentTab = "variance" Then %>
                <!-- Tab 2: 成本异动监控 -->
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>商品名称</th>
                            <th>上月成本</th>
                            <th>本月成本</th>
                            <th>变动额</th>
                            <th>变动率</th>
                            <th>状态</th>
                            <th>归因备注</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                        ' 获取本月和上月的成本对比数据
                        Dim currentMonth, lastMonth
                        currentMonth = Year(Date()) & Right("0" & Month(Date()), 2)
                        If Month(Date()) = 1 Then
                            lastMonth = (Year(Date()) - 1) & "12"
                        Else
                            lastMonth = Year(Date()) & Right("0" & (Month(Date()) - 1), 2)
                        End If
                        
                        Dim rsVariance, varianceSQL
                        ' 使用子查询获取本月和上月成本
                        varianceSQL = "SELECT p.ProductID, p.ProductName, " & _
                            "(SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID = p.ProductID AND CONVERT(VARCHAR(6), CreatedAt, 112) = '" & lastMonth & "') AS LastMonthCost, " & _
                            "(SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID = p.ProductID AND CONVERT(VARCHAR(6), CreatedAt, 112) = '" & currentMonth & "') AS CurrentMonthCost, " & _
                            "p.UnitCost AS CurrentTotalCost " & _
                            "FROM Products p " & _
                            "ORDER BY p.ProductID"
                        
                        Set rsVariance = ExecuteQuery(varianceSQL)
                        If Not rsVariance Is Nothing Then
                            Do While Not rsVariance.EOF
                                vProductId = rsVariance("ProductID")
                                vProductName = rsVariance("ProductName")
                                vLastCost = SafeNum(rsVariance("LastMonthCost"))
                                If IsNull(rsVariance("CurrentMonthCost")) Or (rsVariance("CurrentMonthCost") & "") = "" Then
                                    vCurrentCost = SafeNum(rsVariance("CurrentTotalCost"))
                                Else
                                    vCurrentCost = SafeNum(rsVariance("CurrentMonthCost"))
                                End If
                                
                                ' 如果没有上月数据，使用当前UnitCost作为上月成本
                                If vLastCost = 0 Then vLastCost = vCurrentCost
                                
                                vVariance = vCurrentCost - vLastCost
                                If vLastCost > 0 Then
                                    vVarianceRate = (vVariance / vLastCost) * 100
                                Else
                                    vVarianceRate = 0
                                End If
                                
                                isWarning = Abs(vVarianceRate) > 5
                                statusClass = IIf(isWarning, "status-warning", "status-normal")
                                statusText = IIf(isWarning, "预警", "正常")
                                
                                ' 获取归因备注
                                noteKey2 = "CostVarianceNote_" & vProductId & "_" & currentMonth
                                varianceNote = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = '" & noteKey2 & "'")
                                If varianceNote = "0" Then varianceNote = ""
                        %>
                        <tr <%= IIf(isWarning, "class='warning-row'", "") %>>
                            <td><strong><%= Server.HTMLEncode(vProductName) %></strong></td>
                            <td class="cost-value">¥<%= FormatNumber(vLastCost, 2) %></td>
                            <td class="cost-value">¥<%= FormatNumber(vCurrentCost, 2) %></td>
                            <td class="cost-value <%= IIf(vVariance > 0, "variance-up", IIf(vVariance < 0, "variance-down", "variance-neutral")) %>">
                                <%= IIf(vVariance > 0, "+", "") %><%= FormatNumber(vVariance, 2) %>
                            </td>
                            <td class="cost-value <%= IIf(vVarianceRate > 0, "variance-up", IIf(vVarianceRate < 0, "variance-down", "variance-neutral")) %>">
                                <%= IIf(vVarianceRate > 0, "+", "") %><%= FormatNumber(vVarianceRate, 1) %>%
                            </td>
                            <td><span class="status-badge <%= statusClass %>"><%= statusText %></span></td>
                            <td>
                                <% If canEdit Then %>
                                <form method="post" action="?tab=variance" style="display:flex;gap:8px;">
                                    <%= GetCSRFTokenField() %>
                                    <input type="hidden" name="action" value="save_variance_note">
                                    <input type="hidden" name="productId" value="<%= vProductId %>">
                                    <input type="text" name="varianceNote" class="variance-note-input" 
                                           value="<%= Server.HTMLEncode(varianceNote) %>" 
                                           placeholder="如：采购价上涨/物流涨价...">
                                    <button type="submit" class="btn btn-sm btn-secondary"><i class="fas fa-save"></i></button>
                                </form>
                                <% Else %>
                                <span style="color:#888;"><%= IIf(varianceNote <> "", Server.HTMLEncode(varianceNote), "-") %></span>
                                <% End If %>
                            </td>
                        </tr>
                        <%
                                rsVariance.MoveNext
                            Loop
                            rsVariance.Close
                            Set rsVariance = Nothing
                        End If
                        %>
                    </tbody>
                </table>
                
            <% ElseIf currentTab = "chain" Then %>
                <!-- Tab 4: 成本传导链 -->
                <div style="padding:10px;">
                    <% If canEdit Then %>
                    <div style="display:flex;gap:15px;margin-bottom:25px;flex-wrap:wrap;">
                        <form method="post" action="?tab=chain" style="display:inline;">
                            <%= GetCSRFTokenField() %>
                            <input type="hidden" name="action" value="auto_calc_all">
                            <button type="submit" class="btn btn-primary" onclick="return confirm('确认要自动计算所有产品的BOM/单位成本？此操作将覆盖手动输入的成本。')">
                                <i class="fas fa-calculator"></i> 自动计算全部产品成本
                            </button>
                        </form>
                        <form method="post" action="?tab=chain" style="display:inline;">
                            <%= GetCSRFTokenField() %>
                            <input type="hidden" name="action" value="auto_calc_orders">
                            <button type="submit" class="btn btn-primary" onclick="return confirm('确认要自动计算所有订单的利润？')">
                                <i class="fas fa-receipt"></i> 自动更新全部订单利润
                            </button>
                        </form>
                    </div>
                    <% End If %>
                                
                    <!-- 成本传导链路可视化 -->
                    <div style="background:#1a1a2e;border-radius:12px;padding:30px;margin-bottom:25px;border:1px solid rgba(255,255,255,0.06);">
                        <h4 style="color:#00bcd4;margin-bottom:25px;display:flex;align-items:center;gap:10px;">
                            <i class="fas fa-project-diagram"></i> 成本自动传导链路
                        </h4>
                        <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:15px;align-items:start;">
                            <!-- 节点1: 采购原料 -->
                            <div style="background:#2d2d44;border-radius:10px;padding:15px;text-align:center;border-top:3px solid #4CAF50;">
                                <i class="fas fa-truck-loading" style="font-size:28px;color:#4CAF50;margin-bottom:10px;"></i>
                                <h5 style="color:#e0e0e0;margin:5px 0;">① 采购原料</h5>
                                <p style="color:#888;font-size:11px;margin:8px 0;">SupplierPrices</p>
                                <div style="background:#1a1a2e;border-radius:6px;padding:8px;font-size:11px;color:#aaa;">
                                    原料单价<br>
                                    <span style="color:#4CAF50;font-weight:600;">RawMaterialInventory</span>
                                </div>
                            </div>
                            <!-- 箭头 -->
                            <div style="display:flex;align-items:center;justify-content:center;height:100%;">
                                <i class="fas fa-arrow-right" style="font-size:24px;color:#00bcd4;"></i>
                            </div>
                            <!-- 节点2: 香调生产 -->
                            <div style="background:#2d2d44;border-radius:10px;padding:15px;text-align:center;border-top:3px solid #2196F3;">
                                <i class="fas fa-flask" style="font-size:28px;color:#2196F3;margin-bottom:10px;"></i>
                                <h5 style="color:#e0e0e0;margin:5px 0;">② 香调/基香</h5>
                                <p style="color:#888;font-size:11px;margin:8px 0;">RecipeAccords/NoteIngredients</p>
                                <div style="background:#1a1a2e;border-radius:6px;padding:8px;font-size:11px;color:#aaa;">
                                    原料配比 × 单价<br>
                                    <span style="color:#2196F3;font-weight:600;">CE_CalculateNoteCost()</span>
                                </div>
                            </div>
                            <!-- 箭头 -->
                            <div style="display:flex;align-items:center;justify-content:center;height:100%;">
                                <i class="fas fa-arrow-right" style="font-size:24px;color:#00bcd4;"></i>
                            </div>
                            <!-- 节点3: 产品BOM -->
                            <div style="background:#2d2d44;border-radius:10px;padding:15px;text-align:center;border-top:3px solid #FF9800;">
                                <i class="fas fa-box" style="font-size:28px;color:#FF9800;margin-bottom:10px;"></i>
                                <h5 style="color:#e0e0e0;margin:5px 0;">③ 产品BOM</h5>
                                <p style="color:#888;font-size:11px;margin:8px 0;">ProductNoteRatios</p>
                                <div style="background:#1a1a2e;border-radius:6px;padding:8px;font-size:11px;color:#aaa;">
                                    香调成本 × 配比% + 瓶身<br>
                                    <span style="color:#FF9800;font-weight:600;">CE_CalculateProductBOMCost()</span>
                                </div>
                            </div>
                            <!-- 箭头 -->
                            <div style="display:flex;align-items:center;justify-content:center;height:100%;">
                                <i class="fas fa-arrow-right" style="font-size:24px;color:#00bcd4;"></i>
                            </div>
                            <!-- 节点4: 销售订单 -->
                            <div style="background:#2d2d44;border-radius:10px;padding:15px;text-align:center;border-top:3px solid #9C27B0;">
                                <i class="fas fa-shopping-cart" style="font-size:28px;color:#9C27B0;margin-bottom:10px;"></i>
                                <h5 style="color:#e0e0e0;margin:5px 0;">④ 销售订单</h5>
                                <p style="color:#888;font-size:11px;margin:8px 0;">Orders.CostAmount</p>
                                <div style="background:#1a1a2e;border-radius:6px;padding:8px;font-size:11px;color:#aaa;">
                                    售价 - 成本 - 运费<br>
                                    <span style="color:#9C27B0;font-weight:600;">CE_UpdateOrderCosts()</span>
                                </div>
                            </div>
                        </div>
                    </div>
                                
                    <!-- 成本传导状态 -->
                    <%
                    ' 使用预加载的统计数据（无需额外DB查询）
                    Dim chainSummaryJSON
                    chainSummaryJSON = "{""totalProducts"":" & CE_Stats("totalProducts") & ",""updatedProducts"":" & CE_Stats("updatedProducts") _
                        & ",""totalOrders"":" & CE_Stats("totalValidOrders") & ",""updatedOrders"":" & CE_Stats("updatedOrders") _
                        & ",""lastUpdate"":""" & Session("CE_LastUpdateTime") & """}"
                    %>
                    <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(200px, 1fr));gap:15px;margin-bottom:25px;">
                        <div style="background:#2d2d44;border-radius:10px;padding:20px;text-align:center;border:1px solid rgba(255,255,255,0.06);">
                            <div style="font-size:12px;color:#888;margin-bottom:8px;">产品总数</div>
                            <div style="font-size:28px;font-weight:700;color:#e0e0e0;"><%= CE_Stats("totalProducts") %></div>
                            <div style="font-size:11px;color:#4CAF50;margin-top:5px;">已更新成本: <%= CE_Stats("updatedProducts") %></div>
                        </div>
                        <div style="background:#2d2d44;border-radius:10px;padding:20px;text-align:center;border:1px solid rgba(255,255,255,0.06);">
                            <div style="font-size:12px;color:#888;margin-bottom:8px;">订单总数</div>
                            <div style="font-size:28px;font-weight:700;color:#e0e0e0;"><%= CE_Stats("allOrders") %></div>
                            <div style="font-size:11px;color:#9C27B0;margin-top:5px;">已计算利润: <%= CE_Stats("updatedOrders") %></div>
                        </div>
                        <div style="background:#2d2d44;border-radius:10px;padding:20px;text-align:center;border:1px solid rgba(255,255,255,0.06);">
                            <div style="font-size:12px;color:#888;margin-bottom:8px;">原材料种类</div>
                            <div style="font-size:28px;font-weight:700;color:#e0e0e0;"><%= CE_Stats("rawMaterials") %></div>
                            <div style="font-size:11px;color:#888;margin-top:5px;">原料库存</div>
                        </div>
                        <div style="background:#2d2d44;border-radius:10px;padding:20px;text-align:center;border:1px solid rgba(255,255,255,0.06);">
                            <div style="font-size:12px;color:#888;margin-bottom:8px;">上次自动更新</div>
                            <div style="font-size:16px;font-weight:600;color:#e0e0e0;"><%= IIf(Session("CE_LastUpdateTime") <> "", FormatDateField(Session("CE_LastUpdateTime")), "尚未执行") %></div>
                            <div style="font-size:11px;color:#888;margin-top:5px;">点击上方按钮触发</div>
                        </div>
                    </div>
                                
                    <!-- 香调成本明细表 -->
                    <div style="margin-top:25px;">
                        <h4 style="color:#e0e0e0;margin-bottom:15px;display:flex;align-items:center;gap:10px;">
                            <i class="fas fa-list"></i> 香调(Note)成本明细
                        </h4>
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>NoteID</th>
                                    <th>香调名称</th>
                                    <th>类型</th>
                                    <th>自动计算成本</th>
                                    <th>状态</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Dim rsChainNotes, chainNoteSQL, noteId, noteName, noteType, noteCost, chainCostStr
                                chainNoteSQL = "SELECT NoteID, NoteName, NoteType FROM FragranceNotes WHERE IsActive=1 ORDER BY NoteID"
                                Set rsChainNotes = ExecuteQuery(chainNoteSQL)
                                If Not rsChainNotes Is Nothing Then
                                    Do While Not rsChainNotes.EOF
                                        noteId = rsChainNotes("NoteID")
                                        noteName = rsChainNotes("NoteName")
                                        noteType = rsChainNotes("NoteType")
                                        noteCost = CE_GetCachedNoteCost(noteId)
                                %>
                                <tr>
                                    <td><%= noteId %></td>
                                    <td><strong><%= HTMLEncode(noteName) %></strong></td>
                                    <td><span style="color:#888;"><%= HTMLEncode(noteType) %></span></td>
                                    <td class="cost-value">¥<%= FormatNumber(noteCost, 4) %></td>
                                    <td><span class="status-badge <%= IIf(noteCost > 0, "status-normal", "status-warning") %>"><%= IIf(noteCost > 0, "已传导", "待核算") %></span></td>
                                </tr>
                                <%
                                        rsChainNotes.MoveNext
                                    Loop
                                    rsChainNotes.Close
                                    Set rsChainNotes = Nothing
                                End If
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
                            
            <% ElseIf currentTab = "history" Then %>
                <!-- Tab 3: 成本变更历史 -->
                <div class="timeline">
                    <%
                    Dim rsHistory, historySQL
                    historySQL = "SELECT pc.*, p.ProductName " & _
                        "FROM ProductCosts pc " & _
                        "LEFT JOIN Products p ON pc.ProductID = p.ProductID " & _
                        "ORDER BY pc.CreatedAt DESC"
                    
                    Dim hDate, hProduct, hType, hName, hCost, hCreator, typeLabel, typeIcon
                    Set rsHistory = ExecuteQuery(historySQL)
                    If Not rsHistory Is Nothing Then
                        Do While Not rsHistory.EOF
                            hDate = rsHistory("CreatedAt")
                            hProduct = IIf(IsNull(rsHistory("ProductName")), "未知商品", rsHistory("ProductName"))
                            hType = rsHistory("CostType")
                            hName = rsHistory("CostName")
                            hCost = SafeNum(rsHistory("TotalCost"))
                            hCreator = IIf(IsNull(rsHistory("CreatedBy")), "系统", rsHistory("CreatedBy"))
                            
                            Select Case hType
                                Case "BOM": typeLabel = "BOM成本": typeIcon = "fa-cubes"
                                Case "Purchase": typeLabel = "采购成本": typeIcon = "fa-shopping-cart"
                                Case "Packaging": typeLabel = "包材成本": typeIcon = "fa-box-open"
                                Case Else: typeLabel = "其他成本": typeIcon = "fa-tag"
                            End Select
                    %>
                    <div class="timeline-item">
                        <div class="timeline-date"><%= FormatDateField(hDate) %></div>
                        <div class="timeline-content">
                            <div class="timeline-title">
                                <i class="fas <%= typeIcon %>"></i> 
                                <%= Server.HTMLEncode(hProduct) %> - <%= typeLabel %>
                            </div>
                            <div class="timeline-detail">
                                成本项：<%= Server.HTMLEncode(hName) %> | 
                                金额：<span class="cost-value">¥<%= FormatNumber(hCost, 2) %></span> | 
                                操作人：<%= Server.HTMLEncode(hCreator) %>
                            </div>
                        </div>
                    </div>
                    <%
                            rsHistory.MoveNext
                        Loop
                        rsHistory.Close
                        Set rsHistory = Nothing
                    End If
                    %>
                </div>
            <% End If %>
        </div>
    </div>
    
    <!-- 成本编辑模态框 -->
    <div id="costModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fas fa-edit"></i> 编辑商品成本 - <span id="modalProductName"></span></h3>
                <button class="modal-close" onclick="closeCostModal()">&times;</button>
            </div>
            <form method="post" action="?tab=products">
                <div class="modal-body">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="save_product_cost">
                    <input type="hidden" name="productId" id="modalProductId" value="">
                    <input type="hidden" name="costItemCount" id="costItemCount" value="0">
                    
                    <div id="costItemsContainer">
                        <!-- 成本项将动态添加到这里 -->
                    </div>
                    
                    <% If canEdit Then %>
                    <button type="button" class="btn btn-secondary" onclick="addCostItem()">
                        <i class="fas fa-plus"></i> 添加成本项
                    </button>
                    <% End If %>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" onclick="closeCostModal()">取消</button>
                    <% If canEdit Then %>
                    <button type="submit" class="btn btn-primary"><i class="fas fa-save"></i> 保存</button>
                    <% End If %>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        let costItemIndex = 0;
        
        function openCostModal(productId, productName) {
            document.getElementById('modalProductId').value = productId;
            document.getElementById('modalProductName').textContent = productName;
            document.getElementById('costItemsContainer').innerHTML = '';
            costItemIndex = 0;
            
            // 加载现有成本数据
            loadExistingCosts(productId);
            
            document.getElementById('costModal').classList.add('active');
        }
        
        function closeCostModal() {
            document.getElementById('costModal').classList.remove('active');
        }
        
        function addCostItem(type, name, unitCost, quantity, effectiveDate) {
            costItemIndex++;
            type = type || 'BOM';
            name = name || '';
            unitCost = unitCost || '';
            quantity = quantity || '1';
            effectiveDate = effectiveDate || new Date().toISOString().split('T')[0];
            
            const container = document.getElementById('costItemsContainer');
            const itemDiv = document.createElement('div');
            itemDiv.className = 'cost-item';
            itemDiv.innerHTML = `
                <div class="cost-item-header">
                    <span class="cost-item-title">成本项 #${costItemIndex}</span>
                    <button type="button" class="cost-item-remove" onclick="removeCostItem(this)">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>成本类型</label>
                        <select name="costType_${costItemIndex}" required>
                            <option value="BOM" ${type === 'BOM' ? 'selected' : ''}>BOM成本</option>
                            <option value="Purchase" ${type === 'Purchase' ? 'selected' : ''}>采购价</option>
                            <option value="Packaging" ${type === 'Packaging' ? 'selected' : ''}>包材成本</option>
                            <option value="Other" ${type === 'Other' ? 'selected' : ''}>其他</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>成本名称</label>
                        <input type="text" name="costName_${costItemIndex}" value="${name}" placeholder="如：原料A、包装盒" required>
                    </div>
                    <div class="form-group">
                        <label>单位成本 (¥)</label>
                        <input type="number" name="unitCost_${costItemIndex}" value="${unitCost}" step="0.01" min="0" required onchange="updateTotal(this)">
                    </div>
                    <div class="form-group">
                        <label>用量</label>
                        <input type="number" name="quantity_${costItemIndex}" value="${quantity}" step="0.01" min="0.01" required onchange="updateTotal(this)">
                    </div>
                </div>
                <div class="form-group">
                    <label>生效日期</label>
                    <input type="date" name="effectiveDate_${costItemIndex}" value="${effectiveDate}" required>
                </div>
            `;
            container.appendChild(itemDiv);
            document.getElementById('costItemCount').value = costItemIndex;
        }
        
        function removeCostItem(btn) {
            btn.closest('.cost-item').remove();
            // 重新编号
            const items = document.querySelectorAll('.cost-item');
            items.forEach((item, idx) => {
                item.querySelector('.cost-item-title').textContent = '成本项 #' + (idx + 1);
            });
            costItemIndex = items.length;
            document.getElementById('costItemCount').value = costItemIndex;
        }
        
        function updateTotal(input) {
            // 自动计算总成本的逻辑（前端展示用）
            const item = input.closest('.cost-item');
            const unitCost = parseFloat(item.querySelector('[name^="unitCost_"]').value) || 0;
            const quantity = parseFloat(item.querySelector('[name^="quantity_"]').value) || 0;
            // 总成本在服务器端计算
        }
        
        function loadExistingCosts(productId) {
            // 通过AJAX或页面内嵌数据加载现有成本
            // 这里简化处理，实际项目中可以使用AJAX
            <% 
            Dim rsCostData, costDataSQL
            costDataSQL = "SELECT ProductID, CostType, CostName, UnitCost, Quantity, EffectiveDate FROM ProductCosts ORDER BY ProductID, CostID"
            Dim cdProductId, cdType, cdName, cdUnitCost, cdQuantity, cdEffectiveDate
            Set rsCostData = ExecuteQuery(costDataSQL)
            If Not rsCostData Is Nothing Then
                Do While Not rsCostData.EOF
                    cdProductId = rsCostData("ProductID")
                    cdType = rsCostData("CostType")
                    cdName = rsCostData("CostName")
                    cdUnitCost = rsCostData("UnitCost")
                    cdQuantity = rsCostData("Quantity")
                    cdEffectiveDate = FormatDateField(rsCostData("EffectiveDate"))
                    If cdEffectiveDate = "-" Then cdEffectiveDate = Date()
            %>
            if (productId === <%= cdProductId %>) {
                addCostItem('<%= cdType %>', '<%= SafeOutput(cdName) %>', <%= cdUnitCost %>, <%= cdQuantity %>, '<%= cdEffectiveDate %>');
            }
            <%
                    rsCostData.MoveNext
                Loop
                rsCostData.Close
                Set rsCostData = Nothing
            End If
            %>
        }
        
        // 点击模态框外部关闭
        document.getElementById('costModal').addEventListener('click', function(e) {
            if (e.target === this) closeCostModal();
        });
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
