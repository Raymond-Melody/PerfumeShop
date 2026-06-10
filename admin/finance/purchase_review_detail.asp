<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
' ========== SafeNum函数 ==========
Function SafeNum(val)
    If IsNull(val) Then
        SafeNum = 0
    ElseIf val = "" Then
        SafeNum = 0
    ElseIf Not IsNumeric(val) Then
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

' ========== 获取供应商名称 ==========
Function GetSupplierName(supplierID)
    Dim sql, rs, result
    result = "未知供应商"
    On Error Resume Next
    sql = "SELECT SupplierName FROM Suppliers WHERE SupplierID=" & SafeNum(supplierID)
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs.EOF Then
            If Not IsNull(rs("SupplierName")) Then
                result = CStr(rs("SupplierName"))
            End If
        End If
        rs.Close
        Set rs = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
    GetSupplierName = result
End Function

' ========== 获取供应商信息 ==========
Function GetSupplierInfo(supplierID)
    Dim sql, rs, result
    result = ""
    On Error Resume Next
    sql = "SELECT ContactName, ContactPhone, Address FROM Suppliers WHERE SupplierID=" & SafeNum(supplierID)
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs.EOF Then
            Dim contactName, contactPhone, address
            contactName = ""
            contactPhone = ""
            address = ""
            If Not IsNull(rs("ContactName")) Then contactName = CStr(rs("ContactName"))
            If Not IsNull(rs("ContactPhone")) Then contactPhone = CStr(rs("ContactPhone"))
            If Not IsNull(rs("Address")) Then address = CStr(rs("Address"))
            result = "联系人：" & contactName & " | 电话：" & contactPhone & " | 地址：" & address
        End If
        rs.Close
        Set rs = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
    GetSupplierInfo = result
End Function

' ========== 获取提交人名称 ==========
Function GetCreatorName(adminID)
    Dim sql, rs, result
    result = "未知"
    On Error Resume Next
    sql = "SELECT AdminName FROM AdminUsers WHERE AdminID=" & SafeNum(adminID)
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs.EOF Then
            If Not IsNull(rs("AdminName")) Then
                result = CStr(rs("AdminName"))
            End If
        End If
        rs.Close
        Set rs = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
    GetCreatorName = result
End Function

' ========== 获取状态中文名称 ==========
Function GetStatusName(statusCode)
    If statusCode = "Draft" Then
        GetStatusName = "草稿"
    ElseIf statusCode = "Submitted" Then
        GetStatusName = "待审核"
    ElseIf statusCode = "FinanceApproved" Then
        GetStatusName = "已审批"
    ElseIf statusCode = "Ordered" Then
        GetStatusName = "已下单"
    ElseIf statusCode = "Received" Then
        GetStatusName = "已收货"
    ElseIf statusCode = "Completed" Then
        GetStatusName = "已完成"
    ElseIf statusCode = "Rejected" Then
        GetStatusName = "已拒绝"
    Else
        GetStatusName = statusCode
    End If
End Function

' ========== 获取分类名称 ==========
Function GetCategoryName(catCode)
    If catCode = "RAW" Then
        GetCategoryName = "原材料"
    ElseIf catCode = "BASE" Then
        GetCategoryName = "基香原料"
    ElseIf catCode = "PACK" Then
        GetCategoryName = "包装材料"
    ElseIf catCode = "MARKET" Then
        GetCategoryName = "营销物料"
    Else
        GetCategoryName = catCode
    End If
End Function

' ========== 获取采购订单详情 ==========
Dim purchaseID
purchaseID = SafeNum(Request.QueryString("id"))

If purchaseID = 0 Then
    Response.Write "<div style='color: #ff5252;'>无效的采购单ID</div>"
    Response.End
End If

Dim orderSQL, orderRS
orderSQL = "SELECT PurchaseID, PurchaseNo, SupplierID, CategoryCode, CAST(ISNULL(TotalAmount,0) AS FLOAT) as TotalAmount, Status, CreatedBy, CreatedAt, Remarks, ExpectedDate FROM PurchaseOrders WHERE PurchaseID=" & purchaseID
Set orderRS = ExecuteQuery(orderSQL)

If orderRS Is Nothing Then
    Response.Write "<div style='color: #ff5252;'>查询失败</div>"
    Response.End
End If

If orderRS.EOF Then
    Response.Write "<div style='color: #ff5252;'>采购订单不存在</div>"
    orderRS.Close
    Set orderRS = Nothing
    Response.End
End If

' 安全获取订单信息
Dim pNo, pSupplierID, pCategory, pAmount, pStatus, pCreator, pDate, pRemarks, pExpectedDate
pNo = ""
pSupplierID = 0
pCategory = ""
pAmount = 0
pStatus = ""
pCreator = 0
pDate = ""
pRemarks = ""
pExpectedDate = ""

On Error Resume Next
If Not IsNull(orderRS("PurchaseNo")) Then pNo = CStr(orderRS("PurchaseNo"))
If Not IsNull(orderRS("SupplierID")) Then pSupplierID = SafeNum(orderRS("SupplierID"))
If Not IsNull(orderRS("CategoryCode")) Then pCategory = CStr(orderRS("CategoryCode"))
If Not IsNull(orderRS("TotalAmount")) Then pAmount = SafeNum(orderRS("TotalAmount"))
If Not IsNull(orderRS("Status")) Then pStatus = CStr(orderRS("Status"))
If Not IsNull(orderRS("CreatedBy")) Then pCreator = SafeNum(orderRS("CreatedBy"))
If Not IsNull(orderRS("CreatedAt")) Then pDate = CStr(orderRS("CreatedAt"))
If Not IsNull(orderRS("Remarks")) Then pRemarks = CStr(orderRS("Remarks"))
If Not IsNull(orderRS("ExpectedDate")) Then pExpectedDate = CStr(orderRS("ExpectedDate"))
On Error GoTo 0

orderRS.Close
Set orderRS = Nothing
%>

<!-- 订单基本信息 -->
<div class="info-grid">
    <div class="info-item">
        <div class="info-label">采购单号</div>
        <div class="info-value"><%= Server.HTMLEncode(pNo) %></div>
    </div>
    <div class="info-item">
        <div class="info-label">供应商</div>
        <div class="info-value"><%= Server.HTMLEncode(GetSupplierName(pSupplierID)) %></div>
    </div>
    <div class="info-item">
        <div class="info-label">采购分类</div>
        <div class="info-value"><%= Server.HTMLEncode(GetCategoryName(pCategory)) %></div>
    </div>
    <div class="info-item">
        <div class="info-label">订单状态</div>
        <div class="info-value"><%= Server.HTMLEncode(GetStatusName(pStatus)) %></div>
    </div>
    <div class="info-item">
        <div class="info-label">总金额</div>
        <div class="info-value" style="color: #00bcd4; font-size: 18px;">¥<%= FormatNumber(pAmount, 2) %></div>
    </div>
    <div class="info-item">
        <div class="info-label">提交人</div>
        <div class="info-value"><%= Server.HTMLEncode(GetCreatorName(pCreator)) %></div>
    </div>
    <div class="info-item">
        <div class="info-label">提交日期</div>
        <div class="info-value"><%= Server.HTMLEncode(pDate) %></div>
    </div>
    <div class="info-item">
        <div class="info-label">预计到货</div>
        <div class="info-value"><%= Server.HTMLEncode(pExpectedDate) %></div>
    </div>
</div>

<!-- 供应商详细信息 -->
<div style="background: #1a1a2e; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
    <div style="color: #888; font-size: 12px; margin-bottom: 5px;">供应商信息</div>
    <div style="color: #e0e0e0; font-size: 13px;"><%= Server.HTMLEncode(GetSupplierInfo(pSupplierID)) %></div>
</div>

<!-- 备注信息 -->
<% If pRemarks <> "" Then %>
<div style="background: #1a1a2e; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
    <div style="color: #888; font-size: 12px; margin-bottom: 5px;">备注</div>
    <div style="color: #e0e0e0; font-size: 13px;"><%= Server.HTMLEncode(pRemarks) %></div>
</div>
<% End If %>

<!-- 明细列表 -->
<h4 style="color: #e0e0e0; margin: 20px 0 15px 0;"><i class="fas fa-list"></i> 采购明细</h4>
<table class="detail-table">
    <thead>
        <tr>
            <th>序号</th>
            <th>物品名称</th>
            <th>规格</th>
            <th>数量</th>
            <th>单价</th>
            <th>小计</th>
        </tr>
    </thead>
    <tbody>
    <% 
    Dim detailSQL, detailRS, rowNum
    detailSQL = "SELECT DetailID, PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, CAST(ISNULL(UnitPrice,0) AS FLOAT) as UnitPrice, CAST(ISNULL(TotalPrice,0) AS FLOAT) as TotalPrice, ReceivedQty FROM PurchaseOrderDetails WHERE PurchaseID=" & purchaseID & " ORDER BY DetailID"
    Set detailRS = ExecuteQuery(detailSQL)
    rowNum = 0
    
    If Not detailRS Is Nothing Then
        Do While Not detailRS.EOF
            rowNum = rowNum + 1
            Dim itemName, itemSpec, quantity, unitPrice, totalPrice
            itemName = ""
            itemSpec = ""
            quantity = 0
            unitPrice = 0
            totalPrice = 0
            
            On Error Resume Next
            If Not IsNull(detailRS("ItemName")) Then itemName = CStr(detailRS("ItemName"))
            If Not IsNull(detailRS("Specification")) Then itemSpec = CStr(detailRS("Specification"))
            If Not IsNull(detailRS("Quantity")) Then quantity = SafeNum(detailRS("Quantity"))
            If Not IsNull(detailRS("UnitPrice")) Then unitPrice = SafeNum(detailRS("UnitPrice"))
            If Not IsNull(detailRS("TotalPrice")) Then totalPrice = SafeNum(detailRS("TotalPrice"))
            On Error GoTo 0
    %>
        <tr>
            <td><%= rowNum %></td>
            <td><%= Server.HTMLEncode(itemName) %></td>
            <td><%= Server.HTMLEncode(itemSpec) %></td>
            <td><%= quantity %></td>
            <td>¥<%= FormatNumber(unitPrice, 2) %></td>
            <td>¥<%= FormatNumber(totalPrice, 2) %></td>
        </tr>
    <% 
            detailRS.MoveNext
        Loop
        detailRS.Close
        Set detailRS = Nothing
    End If
    
    If rowNum = 0 Then
    %>
        <tr>
            <td colspan="6" style="text-align: center; color: #888;">暂无明细数据</td>
        </tr>
    <% End If %>
    </tbody>
    <tfoot>
        <tr style="background: #2d2d44;">
            <td colspan="5" style="text-align: right; font-weight: bold;">合计：</td>
            <td style="color: #00bcd4; font-weight: bold;">¥<%= FormatNumber(pAmount, 2) %></td>
        </tr>
    </tfoot>
</table>

<!-- 审核历史 -->
<h4 style="color: #e0e0e0; margin: 30px 0 15px 0;"><i class="fas fa-history"></i> 审核记录</h4>
<table class="detail-table">
    <thead>
        <tr>
            <th>审核时间</th>
            <th>审核人</th>
            <th>审核结果</th>
            <th>审核金额</th>
            <th>成本归类</th>
            <th>审核意见</th>
        </tr>
    </thead>
    <tbody>
    <% 
    Dim reviewSQL, reviewRS, hasReview
    reviewSQL = "SELECT * FROM PurchaseCostReview WHERE PurchaseID=" & purchaseID & " ORDER BY ReviewedAt DESC"
    Set reviewRS = ExecuteQuery(reviewSQL)
    hasReview = False
    
    If Not reviewRS Is Nothing Then
        Do While Not reviewRS.EOF
            hasReview = True
            Dim rReviewer, rStatus, rAmount, rAllocation, rComments, rDate
            rReviewer = 0
            rStatus = ""
            rAmount = 0
            rAllocation = ""
            rComments = ""
            rDate = ""
            
            On Error Resume Next
            If Not IsNull(reviewRS("ReviewerID")) Then rReviewer = SafeNum(reviewRS("ReviewerID"))
            If Not IsNull(reviewRS("ReviewStatus")) Then rStatus = CStr(reviewRS("ReviewStatus"))
            If Not IsNull(reviewRS("ReviewAmount")) Then rAmount = SafeNum(reviewRS("ReviewAmount"))
            If Not IsNull(reviewRS("CostAllocation")) Then rAllocation = CStr(reviewRS("CostAllocation"))
            If Not IsNull(reviewRS("ReviewComments")) Then rComments = CStr(reviewRS("ReviewComments"))
            If Not IsNull(reviewRS("ReviewedAt")) Then rDate = CStr(reviewRS("ReviewedAt"))
            On Error GoTo 0
            
            Dim rStatusText, rAllocationText
            If rStatus = "Approved" Then
                rStatusText = "<span style='color: #4CAF50;'>已通过</span>"
            ElseIf rStatus = "Rejected" Then
                rStatusText = "<span style='color: #ff5252;'>已驳回</span>"
            Else
                rStatusText = rStatus
            End If
            
            If rAllocation = "PRODUCT_COST" Then
                rAllocationText = "产品成本"
            ElseIf rAllocation = "OPERATION_COST" Then
                rAllocationText = "运营成本"
            ElseIf rAllocation = "MARKETING_COST" Then
                rAllocationText = "营销成本"
            Else
                rAllocationText = rAllocation
            End If
    %>
        <tr>
            <td><%= Server.HTMLEncode(rDate) %></td>
            <td><%= Server.HTMLEncode(GetCreatorName(rReviewer)) %></td>
            <td><%= rStatusText %></td>
            <td>¥<%= FormatNumber(rAmount, 2) %></td>
            <td><%= Server.HTMLEncode(rAllocationText) %></td>
            <td><%= Server.HTMLEncode(rComments) %></td>
        </tr>
    <% 
            reviewRS.MoveNext
        Loop
        reviewRS.Close
        Set reviewRS = Nothing
    End If
    
    If Not hasReview Then
    %>
        <tr>
            <td colspan="6" style="text-align: center; color: #888;">暂无审核记录</td>
        </tr>
    <% End If %>
    </tbody>
</table>
