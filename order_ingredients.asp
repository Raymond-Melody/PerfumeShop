<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' ============================================
' 通用成分分割函数 - 支持所有类型的分隔符
' 包括：逗号、空格、换行符、NBSP、全角空格等
' ============================================
Function SplitIngredients(rawStr)
    Dim result, arr, item, idx
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
    For idx = 0 To UBound(arr)
        item = Trim(arr(idx))
        If item <> "" And Not result.Exists(item) Then
            result.Add item, True
        End If
    Next
    
    Set SplitIngredients = result
End Function
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()

' 获取订单ID
Dim orderId
orderId = Request.QueryString("order_id")

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Write "无效的订单ID"
    Response.End
End If

' 获取订单信息
Dim rsOrder
Set rsOrder = ExecuteQuery("SELECT o.*, u.Username FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID WHERE o.OrderID = " & CLng(orderId))

If rsOrder Is Nothing Or rsOrder.EOF Then
    Response.Write "订单不存在"
    Response.End
End If

Dim orderNo, orderDate, customerName
orderNo = rsOrder("OrderNo")
orderDate = rsOrder("CreatedAt")
customerName = rsOrder("ShippingName")

rsOrder.Close
Set rsOrder = Nothing

' 获取订单商品列表
Dim rsDetails
Set rsDetails = ExecuteQuery("SELECT * FROM OrderDetails WHERE OrderID = " & CLng(orderId))

' 获取订单成分列表并处理去重
Dim rsIngredients
Set rsIngredients = ExecuteQuery("SELECT IngredientName FROM OrderIngredients WHERE OrderID = " & CLng(orderId))

' 使用Dictionary进行成分去重
Dim uniqueIngredients, rawIngredient, splitResult, splitKey
Set uniqueIngredients = CreateObject("Scripting.Dictionary")

If Not rsIngredients Is Nothing Then
    Do While Not rsIngredients.EOF
        rawIngredient = rsIngredients("IngredientName") & ""
        
        ' 使用通用分割函数处理成分
        Set splitResult = SplitIngredients(rawIngredient)
        For Each splitKey In splitResult.Keys
            If Not uniqueIngredients.Exists(splitKey) Then
                uniqueIngredients.Add splitKey, True
            End If
        Next
        Set splitResult = Nothing
        
        rsIngredients.MoveNext
    Loop
    rsIngredients.Close
    Set rsIngredients = Nothing
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>订单成分表 - <%= orderNo %></title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: "Microsoft YaHei", Arial, sans-serif; padding: 40px; background: white; }
        .container { max-width: 800px; margin: 0 auto; }
        
        .header { text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 3px solid #333; }
        .header h1 { font-size: 28px; color: #333; margin-bottom: 10px; }
        .header .subtitle { font-size: 14px; color: #666; }
        
        .info-section { margin-bottom: 30px; }
        .info-row { display: flex; margin-bottom: 10px; font-size: 14px; }
        .info-row .label { width: 120px; font-weight: bold; color: #333; }
        .info-row .value { flex: 1; color: #666; }
        
        .section-title { font-size: 18px; font-weight: bold; color: #333; margin: 30px 0 15px 0; padding-bottom: 10px; border-bottom: 2px solid #333; }
        
        .products-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
        .products-table th, .products-table td { padding: 12px; text-align: left; border: 1px solid #ddd; font-size: 14px; }
        .products-table th { background: #f5f5f5; font-weight: bold; color: #333; }
        .products-table tr:nth-child(even) { background: #fafafa; }
        
        .ingredients-list { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 15px; }
        .ingredient-item { padding: 10px 15px; background: #f8f9fa; border-left: 3px solid #007bff; font-size: 14px; color: #333; }
        
        .notes-section { margin-top: 30px; padding: 15px; background: #fff9e6; border: 1px solid #ffd700; }
        .notes-section .title { font-weight: bold; color: #856404; margin-bottom: 10px; }
        .notes-section ul { margin-left: 20px; }
        .notes-section li { margin-bottom: 5px; color: #666; font-size: 13px; }
        
        .footer { margin-top: 50px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; font-size: 12px; color: #999; }
        
        .btn-print { display: inline-block; padding: 12px 30px; background: #007bff; color: white; text-decoration: none; border-radius: 4px; border: none; cursor: pointer; font-size: 14px; margin-bottom: 20px; }
        .btn-print:hover { background: #0056b3; }
        
        @media print {
            .btn-print, .no-print { display: none; }
            body { padding: 20px; }
            .header { page-break-after: avoid; }
            .section-title { page-break-after: avoid; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="no-print" style="margin-bottom: 20px;">
            <button class="btn-print" onclick="window.print()">🖨️ 打印成分表</button>
            <button class="btn-print btn--neutral" onclick="window.close()">← 关闭</button>
        </div>
        
        <div class="header">
            <h1>产品成分表 / Ingredients List</h1>
            <div class="subtitle">Product Ingredients List</div>
        </div>
        
        <div class="info-section">
            <div class="info-row">
                <div class="label">订单编号：</div>
                <div class="value"><%= orderNo %></div>
            </div>
            <div class="info-row">
                <div class="label">订单日期：</div>
                <div class="value"><%= SafeFormatDateTime(orderDate, 2) %></div>
            </div>
            <div class="info-row">
                <div class="label">客户姓名：</div>
                <div class="value"><%= HTMLEncode(customerName) %></div>
            </div>
            <div class="info-row">
                <div class="label">打印时间：</div>
                <div class="value"><%= Now() %></div>
            </div>
        </div>
        
        <div class="section-title">📦 订单商品明细</div>
        <table class="products-table">
            <thead>
                <tr>
                    <th>商品名称</th>
                    <th>容量规格</th>
                    <th>数量</th>
                    <th>定制备注</th>
                </tr>
            </thead>
            <tbody>
                <%
                If Not rsDetails Is Nothing Then
                    If Not rsDetails.EOF Then
                        Do While Not rsDetails.EOF
                %>
                <tr>
                    <td><strong><%= HTMLEncode(rsDetails("ProductName")) %></strong></td>
                    <td><%= HTMLEncode(rsDetails("VolumeName") & "") %></td>
                    <td><%= rsDetails("Quantity") %></td>
                    <td>
                        <% 
                        Dim customInfo
                        customInfo = ""
                        If Not IsNull(rsDetails("CustomLabel")) And rsDetails("CustomLabel") <> "" Then
                            customInfo = "刻字: " & HTMLEncode(rsDetails("CustomLabel"))
                        End If
                        If customInfo <> "" Then
                            Response.Write customInfo
                        Else
                            Response.Write "-"
                        End If
                        %>
                    </td>
                </tr>
                <%
                            rsDetails.MoveNext
                        Loop
                    Else
                        Response.Write "<tr><td colspan='4' style='text-align:center;color:#999;'>无商品数据</td></tr>"
                    End If
                    rsDetails.Close
                Else
                    Response.Write "<tr><td colspan='4' style='text-align:center;color:#999;'>无法读取商品数据</td></tr>"
                End If
                Set rsDetails = Nothing
                %>
            </tbody>
        </table>
        
        <div class="section-title">🧪 产品成分清单 / Ingredients</div>
        <div style="font-size: 13px; color: #666; margin-bottom: 15px;">
            本产品由以下成分配制而成（按字母顺序排列）。如您对某些成分过敏，请在使用前咨询专业人士。
        </div>
        
        <%
        Dim ingredientCount, ingrKey
        ingredientCount = 0
        If uniqueIngredients.Count > 0 Then
        %>
        <div class="ingredients-list">
            <%
                ' 将Dictionary的Keys转换为数组并排序
                Dim ingrKeys(), i, j, tempKey
                ReDim ingrKeys(uniqueIngredients.Count - 1)
                i = 0
                For Each ingrKey In uniqueIngredients.Keys
                    ingrKeys(i) = ingrKey
                    i = i + 1
                Next
                
                ' 冒泡排序（按字母顺序）
                For i = 0 To UBound(ingrKeys) - 1
                    For j = i + 1 To UBound(ingrKeys)
                        If ingrKeys(i) > ingrKeys(j) Then
                            tempKey = ingrKeys(i)
                            ingrKeys(i) = ingrKeys(j)
                            ingrKeys(j) = tempKey
                        End If
                    Next
                Next
                
                ' 显示排序后的成分
                For i = 0 To UBound(ingrKeys)
                    ingredientCount = ingredientCount + 1
            %>
            <div class="ingredient-item">
                <%= ingredientCount %>. <%= HTMLEncode(ingrKeys(i)) %>
            </div>
            <%
                Next
            %>
        </div>
        <%
            Else
        %>
        <div style="padding: 20px; background: #fff3cd; border: 1px solid #ffc107; color: #856404; text-align: center;">
            ⚠️ 此订单暂无成分信息记录
        </div>
        <%
            End If
        %>
        
        <div class="notes-section">
            <div class="title">📋 重要说明</div>
            <ul>
                <li>本成分表仅供识别产品成分，用于过敏原检查</li>
                <li>成分列表不包含具体数量或比例信息</li>
                <li>所有成分均符合相关国家标准和行业规范</li>
                <li>产品制作详情（包括香调配比、瓶身规格等）请参考商家生产工单</li>
                <li>如有疑问请联系客服部门</li>
            </ul>
        </div>
        
        <div class="footer">
            <p>本文档由系统自动生成，供客户参考</p>
            <p>打印时间: <%= Now() %> | 订单编号: <%= orderNo %></p>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
