<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' ============================================
' 通用成分分割函数 - 支持所有类型的分隔符
' 包括：逗号、空格、换行符、NBSP、全角空格等
' ============================================
Function SplitIngredients(rawStr)
    Dim result, arr, item, i
    Set result = CreateObject("Scripting.Dictionary")
    
    If rawStr = "" Then
        Set SplitIngredients = result
        Exit Function
    End If
    
    ' 统一将所有分隔符转换为英文逗号
    rawStr = Replace(rawStr, "，", ",")      ' 中文逗号
    rawStr = Replace(rawStr, vbCrLf, ",")   ' 回车换行
    rawStr = Replace(rawStr, vbLf, ",")     ' 换行符
    rawStr = Replace(rawStr, vbCr, ",")     ' 回车符
    rawStr = Replace(rawStr, Chr(160), ",") ' NBSP
    rawStr = Replace(rawStr, "　", ",")     ' 全角空格
    
    ' 清理连续逗号
    Do While InStr(rawStr, ",,") > 0
        rawStr = Replace(rawStr, ",,", ",")
    Loop
    
    ' 用逗号分割
    arr = Split(rawStr, ",")
    For i = 0 To UBound(arr)
        item = Trim(arr(i))
        If item <> "" And Not result.Exists(item) Then
            result.Add item, True
        End If
    Next
    
    Set SplitIngredients = result
End Function
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

Dim orderId
orderId = Request.QueryString("order_id")

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Redirect "orders.asp"
    Response.End
End If

' 获取订单基本信息
Dim rsOrder
Set rsOrder = ExecuteQuery("SELECT o.*, u.Username, u.Email FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID WHERE o.OrderID = " & CLng(orderId))

If rsOrder Is Nothing Or rsOrder.EOF Then
    Response.Redirect "orders.asp"
    Response.End
End If

' 存储订单信息到变量
Dim orderNo, totalAmount, paymentMethod, orderStatus, createdAt, updatedAt, notes
Dim shippingName, shippingPhone, shippingAddress, customerName, customerEmail
orderNo = rsOrder("OrderNo")
totalAmount = rsOrder("TotalAmount")
paymentMethod = rsOrder("PaymentMethod")
orderStatus = rsOrder("Status")
createdAt = rsOrder("CreatedAt")
updatedAt = rsOrder("UpdatedAt")
notes = rsOrder("Notes")
shippingName = rsOrder("ShippingName")
shippingPhone = rsOrder("ShippingPhone")
shippingAddress = rsOrder("ShippingAddress")
customerName = rsOrder("Username")
customerEmail = rsOrder("Email")

rsOrder.Close
Set rsOrder = Nothing

' 获取订单商品列表（包含产品类型信息）
Dim rsDetails
Set rsDetails = ExecuteQuery("SELECT od.*, p.ProductType FROM OrderDetails od LEFT JOIN Products p ON od.ProductID = p.ProductID WHERE od.OrderID = " & CLng(orderId))
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>订单详情 #<%= orderNo %> - 管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .order-info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 30px; }
        .info-group { background: #f9f9f9; padding: 15px; border-radius: 8px; border: 1px solid #eee; }
        .info-group h3 { margin-top: 0; border-bottom: 1px solid #ddd; padding-bottom: 10px; margin-bottom: 15px; font-size: 16px; color: #333; }
        .info-row { display: flex; margin-bottom: 8px; font-size: 14px; }
        .info-row .label { width: 100px; color: #666; font-weight: bold; }
        .info-row .value { flex: 1; color: #333; }
        .status-Pending { color: #f39c12; }
        .status-Paid { color: #27ae60; font-weight: bold; }
        .status-Failed { color: #e74c3c; }
        
        .product-card { border: 1px solid #eee; border-radius: 8px; padding: 15px; margin-bottom: 15px; background: #fff; }
        .product-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px dashed #eee; padding-bottom: 10px; margin-bottom: 10px; }
        .product-name { font-weight: bold; font-size: 16px; color: #007bff; }
        
        .custom-detail { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; }
        .note-section { font-size: 13px; }
        .note-section h4 { font-size: 14px; margin-bottom: 5px; color: #555; }
        .note-item { margin-bottom: 3px; display: flex; justify-content: space-between; padding: 2px 5px; background: #f5f5f5; border-radius: 3px; }
        
        .ingredients-box { margin-top: 30px; background: #fff; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; }
        .ingredients-header { background: #f1f8ff; padding: 12px 15px; border-bottom: 1px solid #e0e0e0; font-weight: bold; color: #0366d6; }
        .ingredients-list { padding: 15px; display: flex; flex-wrap: wrap; gap: 8px; }
        .ingredient-tag { background: #e1ecf4; color: #39739d; padding: 4px 10px; border-radius: 15px; font-size: 12px; border: 1px solid #cedee7; }
        
        .action-bar { margin-bottom: 20px; display: flex; gap: 10px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="admin-card">
            <div class="admin-card-header">
                <h2 class="admin-card-title">订单详情 #<%= orderNo %></h2>
                <div class="action-bar">
                    <a href="orders.asp" class="admin-btn admin-btn-outline"><i class="fas fa-arrow-left"></i> 返回列表</a>
                    <a href="../production/order_production.asp?order_id=<%= orderId %>" class="btn-print" target="_blank"><i class="fas fa-print"></i> 打印生产工单</a>
                </div>
            </div>
            
            <div class="admin-card-body">
                <!-- 基本信息 -->
                <div class="order-info-grid">
                    <div class="info-group">
                        <h3><i class="fas fa-info-circle"></i> 订单基本信息</h3>
                        <div class="info-row">
                            <span class="label">订单号:</span>
                            <span class="value"><%= orderNo %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">下单时间:</span>
                            <span class="value"><%= createdAt %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">订单金额:</span>
                            <span class="value"><%= FormatMoney(totalAmount) %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">当前状态:</span>
                            <span class="value status-<%= orderStatus %>">
                                <% Select Case orderStatus
                                    Case "Pending": Response.Write "待支付"
                                    Case "Paid": Response.Write "已支付"
                                    Case "Failed": Response.Write "支付失败"
                                    Case "Refunded": Response.Write "已退款"
                                    Case Else: Response.Write orderStatus
                                End Select %>
                            </span>
                        </div>
                        <div class="info-row">
                            <span class="label">支付方式:</span>
                            <span class="value">
                                <% Select Case paymentMethod & ""
                                    Case "1", "wechat": Response.Write "微信支付"
                                    Case "2", "alipay": Response.Write "支付宝"
                                    Case "3", "paypal": Response.Write "PayPal"
                                    Case "4", "cod": Response.Write "货到付款"
                                    Case Else: Response.Write Server.HTMLEncode(paymentMethod & "")
                                End Select %>
                            </span>
                        </div>
                    </div>
                    
                    <div class="info-group">
                        <h3><i class="fas fa-user"></i> 客户与收货信息</h3>
                        <div class="info-row">
                            <span class="label">客户账号:</span>
                            <span class="value"><%= Server.HTMLEncode(customerName & "") %> (<%= Server.HTMLEncode(customerEmail & "") %>)</span>
                        </div>
                        <div class="info-row">
                            <span class="label">收货人:</span>
                            <span class="value"><%= Server.HTMLEncode(shippingName & "") %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">联系电话:</span>
                            <span class="value"><%= Server.HTMLEncode(shippingPhone & "") %></span>
                        </div>
                        <div class="info-row">
                            <span class="label">收货地址:</span>
                            <span class="value"><%= Server.HTMLEncode(shippingAddress & "") %></span>
                        </div>
                    </div>
                </div>

                <!-- 订单商品 (生产快照) -->
                <h3><i class="fas fa-flask"></i> 生产明细 (商品快照)</h3>
                <%
                If Not rsDetails Is Nothing Then
                    If Not rsDetails.EOF Then
                        Do While Not rsDetails.EOF
                            Dim detailId
                            detailId = rsDetails("DetailID")
                %>
                <div class="product-card">
                    <div class="product-header">
                        <span class="product-name"><%= Server.HTMLEncode(rsDetails("ProductName") & "") %></span>
                        <span>数量: <strong><%= rsDetails("Quantity") %></strong> | 单价: <%= FormatMoney(rsDetails("UnitPrice")) %></span>
                    </div>
                    
                    <div class="custom-detail">
                        <!-- 香调配比 -->
                        <div class="note-section">
                            <h4>香调配比</h4>
                            <%
                            ' 获取前中后调配比
                            Dim rsNoteSels, noteType, noteName, percentage
                            Dim hasNotes
                            hasNotes = False
                            Set rsNoteSels = ExecuteQuery("SELECT s.*, n.NoteName FROM OrderDetailNoteSelections s LEFT JOIN FragranceNotes n ON s.NoteID = n.NoteID WHERE s.DetailID = " & detailId & " ORDER BY s.NoteType")
                            
                            If Not rsNoteSels Is Nothing Then
                                Do While Not rsNoteSels.EOF
                                    hasNotes = True
                                    noteType = rsNoteSels("NoteType")
                                    noteName = rsNoteSels("NoteName")
                                    percentage = rsNoteSels("Percentage")
                            %>
                            <div class="note-item">
                                <span>[<%= noteType %>] <%= noteName %></span>
                                <strong><%= percentage %>%</strong>
                            </div>
                            <%
                                    rsNoteSels.MoveNext
                                Loop
                                rsNoteSels.Close
                                Set rsNoteSels = Nothing
                            End If
                            
                            ' 如果没有香调配比数据，尝试显示成分信息
                            If Not hasNotes Then
                                Dim rsDetailIngredients
                                Set rsDetailIngredients = ExecuteQuery("SELECT IngredientName FROM OrderIngredients WHERE DetailID = " & detailId & " ORDER BY IngredientID")
                                If Not rsDetailIngredients Is Nothing Then
                                    If Not rsDetailIngredients.EOF Then
                                        Do While Not rsDetailIngredients.EOF
                            %>
                            <div class="note-item">
                                <span>• <%= Server.HTMLEncode(rsDetailIngredients("IngredientName") & "") %></span>
                                <strong>原料</strong>
                            </div>
                            <%
                                            rsDetailIngredients.MoveNext
                                        Loop
                                    Else
                                        Response.Write "<div class='text-muted'>无配比数据</div>"
                                    End If
                                    rsDetailIngredients.Close
                                    Set rsDetailIngredients = Nothing
                                Else
                                    Response.Write "<div class='text-muted'>无法加载配比</div>"
                                End If
                            End If
                            %>
                        </div>
                        
                        <!-- 规格参数 -->
                        <div class="note-section">
                            <h4>规格参数</h4>
                            <div class="info-row">
                                <span class="label">容量:</span>
                                <span class="value"><%= Server.HTMLEncode(rsDetails("VolumeML") & "") %>ml (<%= Server.HTMLEncode(rsDetails("VolumeName") & "") %>)</span>
                            </div>
                            <div class="info-row">
                                <span class="label">瓶身:</span>
                                <span class="value"><%= Server.HTMLEncode(rsDetails("BottleName") & "") %></span>
                            </div>
                        </div>
                        
                        <!-- 定制信息 -->
                        <div class="note-section">
                            <h4>定制信息</h4>
                            
                            <!-- 产品类型标识 -->
                            <div class="info-row">
                                <span class="label">产品类型:</span>
                                <span class="value">
                                    <%
                                    Dim productTypeAdmin
                                    productTypeAdmin = rsDetails("ProductType") & ""
                                    Select Case productTypeAdmin
                                        Case "Fixed"
                                            Response.Write "<span style='background:#2196f3;color:white;padding:2px 8px;border-radius:3px;font-size:12px;'>品牌定香</span>"
                                        Case "Custom"
                                            Response.Write "<span style='background:#4caf50;color:white;padding:2px 8px;border-radius:3px;font-size:12px;'>用户定制</span>"
                                        Case "KOL"
                                            Response.Write "<span style='background:#9c27b0;color:white;padding:2px 8px;border-radius:3px;font-size:12px;'>KOL推荐</span>"
                                        Case Else
                                            Response.Write "<span style='background:#95a5a6;color:white;padding:2px 8px;border-radius:3px;font-size:12px;'>未知类型</span>"
                                    End Select
                                    %>
                                </span>
                            </div>
                            
                            <div class="info-row">
                                <span class="label">瓶身刻字:</span>
                                <span class="value">
                                    <% If rsDetails("CustomLabel") <> "" Then %>
                                        <strong style="color: #e67e22; border: 1px solid #e67e22; padding: 2px 5px; border-radius: 3px;"><%= Server.HTMLEncode(rsDetails("CustomLabel") & "") %></strong>
                                    <% Else %>
                                        <span class="text-muted">无刻字</span>
                                    <% End If %>
                                </span>
                            </div>
                        </div>
                        
                        <!-- 成分信息（仅对定制和KOL产品显示，品牌定香产品随包装附成分说明书） -->
                        <%
                        Dim productTypeLC_ingredient
                        productTypeLC_ingredient = LCase(productTypeAdmin & "")
                        If productTypeLC_ingredient = "custom" Or productTypeLC_ingredient = "kol" Then
                        %>
                        <div class="note-section">
                            <h4>成分列表</h4>
                            <%
                            ' 获取该产品的具体成分（通过DetailID关联）并处理去重
                            Dim detailUniqueIngr, detailRawIngr
                            Dim detailIngrKeys, detailIngrKey
                            Dim detailI, detailJ, detailTempKey
                            Dim detailSplitResult, detailSplitKey
                            Set detailUniqueIngr = CreateObject("Scripting.Dictionary")
                            
                            Set rsDetailIngredients = ExecuteQuery("SELECT IngredientName FROM OrderIngredients WHERE DetailID = " & detailId)
                            
                            If Not rsDetailIngredients Is Nothing Then
                                Do While Not rsDetailIngredients.EOF
                                    detailRawIngr = rsDetailIngredients("IngredientName") & ""
                                    
                                    ' 使用通用分割函数处理成分
                                    Set detailSplitResult = SplitIngredients(detailRawIngr)
                                    For Each detailSplitKey In detailSplitResult.Keys
                                        If Not detailUniqueIngr.Exists(detailSplitKey) Then
                                            detailUniqueIngr.Add detailSplitKey, True
                                        End If
                                    Next
                                    Set detailSplitResult = Nothing
                                    
                                    rsDetailIngredients.MoveNext
                                Loop
                                rsDetailIngredients.Close
                                Set rsDetailIngredients = Nothing
                            End If
                            
                            ' 显示去重后的成分
                            If detailUniqueIngr.Count > 0 Then
                                ' 转换为数组并排序
                                ReDim detailIngrKeys(detailUniqueIngr.Count - 1)
                                Dim detailIdx
                                detailIdx = 0
                                For Each detailIngrKey In detailUniqueIngr.Keys
                                    detailIngrKeys(detailIdx) = detailIngrKey
                                    detailIdx = detailIdx + 1
                                Next
                                
                                ' 冒泡排序
                                For detailI = 0 To UBound(detailIngrKeys) - 1
                                    For detailJ = detailI + 1 To UBound(detailIngrKeys)
                                        If detailIngrKeys(detailI) > detailIngrKeys(detailJ) Then
                                            detailTempKey = detailIngrKeys(detailI)
                                            detailIngrKeys(detailI) = detailIngrKeys(detailJ)
                                            detailIngrKeys(detailJ) = detailTempKey
                                        End If
                                    Next
                                Next
                                
                                ' 显示成分
                                For detailI = 0 To UBound(detailIngrKeys)
                            %>
                            <div class="note-item">
                                <span>• <%= Server.HTMLEncode(detailIngrKeys(detailI)) %></span>
                                <strong>成分</strong>
                            </div>
                            <%
                                Next
                            Else
                                Response.Write "<div class='text-muted'>暂无成分数据</div>"
                            End If
                            
                            Set detailUniqueIngr = Nothing
                            %>
                        </div>
                        <% End If %>
                    </div>
                </div>
                <%
                            rsDetails.MoveNext
                        Loop
                        rsDetails.Close
                    Else
                        Response.Write "<div class='alert alert-warning'>此订单无详细商品数据</div>"
                    End If
                    Set rsDetails = Nothing
                End If
                %>

                <!-- 汇总成分 (用于质控) -->
                <div class="ingredients-box">
                    <div class="ingredients-header"><i class="fas fa-microscope"></i> 订单成分汇总 (质控与过敏原管理)</div>
                    <div class="ingredients-list">
                        <%
                        ' 重新查询整个订单的所有成分并处理去重
                        Dim rsAllIngredients, allUniqueIngr, allRawIngr
                        Dim allIngrKeys, allIngrKey
                        Dim allI, allJ, allTempKey
                        Dim allSplitResult, allSplitKey
                        Set allUniqueIngr = CreateObject("Scripting.Dictionary")
                        
                        Set rsAllIngredients = ExecuteQuery("SELECT IngredientName FROM OrderIngredients WHERE OrderID = " & CLng(orderId))
                        
                        If Not rsAllIngredients Is Nothing Then
                            Do While Not rsAllIngredients.EOF
                                allRawIngr = rsAllIngredients("IngredientName") & ""
                                
                                ' 使用通用分割函数处理成分
                                Set allSplitResult = SplitIngredients(allRawIngr)
                                For Each allSplitKey In allSplitResult.Keys
                                    If Not allUniqueIngr.Exists(allSplitKey) Then
                                        allUniqueIngr.Add allSplitKey, True
                                    End If
                                Next
                                Set allSplitResult = Nothing
                                
                                rsAllIngredients.MoveNext
                            Loop
                            rsAllIngredients.Close
                            Set rsAllIngredients = Nothing
                        End If
                        
                        ' 显示去重后的成分
                        If allUniqueIngr.Count > 0 Then
                            ' 转换为数组并排序
                            ReDim allIngrKeys(allUniqueIngr.Count - 1)
                            Dim allIdx
                            allIdx = 0
                            For Each allIngrKey In allUniqueIngr.Keys
                                allIngrKeys(allIdx) = allIngrKey
                                allIdx = allIdx + 1
                            Next
                            
                            ' 冒泡排序
                            For allI = 0 To UBound(allIngrKeys) - 1
                                For allJ = allI + 1 To UBound(allIngrKeys)
                                    If allIngrKeys(allI) > allIngrKeys(allJ) Then
                                        allTempKey = allIngrKeys(allI)
                                        allIngrKeys(allI) = allIngrKeys(allJ)
                                        allIngrKeys(allJ) = allTempKey
                                    End If
                                Next
                            Next
                            
                            ' 显示成分
                            For allI = 0 To UBound(allIngrKeys)
                        %>
                        <span class="ingredient-tag"><%= Server.HTMLEncode(allIngrKeys(allI)) %></span>
                        <%
                            Next
                        Else
                            Response.Write "<div class='text-muted' style='padding:15px;'>此订单暂无详细成分记录</div>"
                        End If
                        
                        Set allUniqueIngr = Nothing
                        %>
                    </div>
                </div>
                
                <% If notes <> "" Then %>
                <div class="info-group" style="margin-top: 20px; background: #fffbe6; border-color: #ffe58f;">
                    <h3><i class="fas fa-sticky-note"></i> 交易说明 / 备注</h3>
                    <div style="font-size: 14px; white-space: pre-wrap;"><%= notes %></div>
                </div>
                <% End If %>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
