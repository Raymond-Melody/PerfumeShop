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

' V14: 会员登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If

' 获取订单ID
Dim orderId
orderId = Request.QueryString("order_id")

If orderId = "" Or Not IsNumeric(orderId) Then
    Response.Write T("order_invalid_id", Empty)
    Response.End
End If

' 获取订单信息
Dim rsOrder
Set rsOrder = ExecuteQuery("SELECT o.*, u.Username FROM Orders o LEFT JOIN Users u ON o.UserID = u.UserID WHERE o.OrderID = " & CLng(orderId))

If rsOrder Is Nothing Or rsOrder.EOF Then
    Response.Write T("order_not_found", Empty)
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

Dim hasOrderIngredients : hasOrderIngredients = False

If Not rsIngredients Is Nothing Then
    Do While Not rsIngredients.EOF
        rawIngredient = rsIngredients("IngredientName") & ""
        hasOrderIngredients = True
        
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

' ========== 回退机制：当OrderIngredients无数据时，从香调链路重新计算成分 ==========
If Not hasOrderIngredients Then
    Dim rsDetailItems, fbDetailId, fbProductId, fbProductType, fbRecipeId
    Set rsDetailItems = ExecuteQuery("SELECT od.DetailID, od.ProductID, p.ProductType, p.RecipeID, p.BaseIngredients FROM OrderDetails od LEFT JOIN Products p ON od.ProductID=p.ProductID WHERE od.OrderID=" & CLng(orderId))
    If Not rsDetailItems Is Nothing Then
        Do While Not rsDetailItems.EOF
            fbDetailId = rsDetailItems("DetailID")
            fbProductId = rsDetailItems("ProductID")
            fbProductType = LCase(rsDetailItems("ProductType") & "")
            fbRecipeId = 0
            On Error Resume Next
            If Not IsNull(rsDetailItems("RecipeID")) Then fbRecipeId = CLng(rsDetailItems("RecipeID"))
            On Error GoTo 0
            
            ' 路径1: 从配方获取成分
            If fbRecipeId > 0 Then
                Dim rsRecipeIngr
                Set rsRecipeIngr = ExecuteQuery("SELECT IngredientName FROM RecipeIngredients WHERE RecipeID=" & fbRecipeId)
                If Not rsRecipeIngr Is Nothing Then
                    Do While Not rsRecipeIngr.EOF
                        rawIngredient = Trim(rsRecipeIngr("IngredientName") & "")
                        If rawIngredient <> "" And Not uniqueIngredients.Exists(rawIngredient) Then
                            uniqueIngredients.Add rawIngredient, True
                        End If
                        rsRecipeIngr.MoveNext
                    Loop
                    rsRecipeIngr.Close
                End If
                Set rsRecipeIngr = Nothing
            End If
            
            ' 路径2: 从香调选择→基香→成分
            Dim rsNoteSels, fbNoteId
            Set rsNoteSels = ExecuteQuery("SELECT NoteID FROM OrderDetailNoteSelections WHERE DetailID=" & fbDetailId)
            If Not rsNoteSels Is Nothing Then
                Do While Not rsNoteSels.EOF
                    fbNoteId = rsNoteSels("NoteID")
                    Dim rsBaseIngr
                    Set rsBaseIngr = ExecuteQuery("SELECT b.Ingredients FROM NoteIngredients ni LEFT JOIN BaseNotes b ON ni.BaseNoteID=b.BaseNoteID WHERE ni.NoteID=" & fbNoteId & " AND b.Ingredients IS NOT NULL AND b.Ingredients <> ''")
                    If Not rsBaseIngr Is Nothing Then
                        Do While Not rsBaseIngr.EOF
                            Set splitResult = SplitIngredients(rsBaseIngr("Ingredients") & "")
                            For Each splitKey In splitResult.Keys
                                If Not uniqueIngredients.Exists(splitKey) Then
                                    uniqueIngredients.Add splitKey, True
                                End If
                            Next
                            Set splitResult = Nothing
                            rsBaseIngr.MoveNext
                        Loop
                        rsBaseIngr.Close
                    End If
                    Set rsBaseIngr = Nothing
                    rsNoteSels.MoveNext
                Loop
                rsNoteSels.Close
            End If
            Set rsNoteSels = Nothing
            
            ' 路径3: 品牌定香产品的BaseIngredients
            Dim fbBaseIngr
            fbBaseIngr = ""
            On Error Resume Next
            fbBaseIngr = Trim(rsDetailItems("BaseIngredients") & "")
            On Error GoTo 0
            If fbBaseIngr <> "" Then
                Set splitResult = SplitIngredients(fbBaseIngr)
                For Each splitKey In splitResult.Keys
                    If Not uniqueIngredients.Exists(splitKey) Then
                        uniqueIngredients.Add splitKey, True
                    End If
                Next
                Set splitResult = Nothing
            End If
            
            rsDetailItems.MoveNext
        Loop
        rsDetailItems.Close
    End If
    Set rsDetailItems = Nothing
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title><% If FEATURE_I18N Then %><%= T("order_ingredients_title", Empty) %><% Else %>订单成分表<% End If %> - <%= orderNo %></title>
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
            <button class="btn-print" onclick="window.print()">🖨️ <% If FEATURE_I18N Then %><%= T("order_ingredients_print", Empty) %><% Else %>打印成分表<% End If %></button>
            <button class="btn-print btn--neutral" onclick="window.close()">← <% If FEATURE_I18N Then %><%= T("order_ingredients_close", Empty) %><% Else %>关闭<% End If %></button>
        </div>
        
        <div class="header">
            <h1><% If FEATURE_I18N Then %><%= T("order_ingredients_main_title", Empty) %><% Else %>产品成分表 / Ingredients List<% End If %></h1>
            <div class="subtitle">Product Ingredients List</div>
        </div>
        
        <div class="info-section">
            <div class="info-row">
                <div class="label"><% If FEATURE_I18N Then %><%= T("order_ingredients_label_order_no", Empty) %><% Else %>订单编号<% End If %>：</div>
                <div class="value"><%= orderNo %></div>
            </div>
            <div class="info-row">
                <div class="label"><% If FEATURE_I18N Then %><%= T("order_ingredients_label_date", Empty) %><% Else %>订单日期<% End If %>：</div>
                <div class="value"><%= SafeFormatDateTime(orderDate, 2) %></div>
            </div>
            <div class="info-row">
                <div class="label"><% If FEATURE_I18N Then %><%= T("order_ingredients_label_customer", Empty) %><% Else %>客户姓名<% End If %>：</div>
                <div class="value"><%= HTMLEncode(customerName) %></div>
            </div>
            <div class="info-row">
                <div class="label"><% If FEATURE_I18N Then %><%= T("order_ingredients_label_print_time", Empty) %><% Else %>打印时间<% End If %>：</div>
                <div class="value"><%= Now() %></div>
            </div>
        </div>
        
        <div class="section-title">📦 <% If FEATURE_I18N Then %><%= T("order_ingredients_section_products", Empty) %><% Else %>订单商品明细<% End If %></div>
        <table class="products-table">
            <thead>
                <tr>
                    <th><% If FEATURE_I18N Then %><%= T("product_name", Empty) %><% Else %>商品名称<% End If %></th>
                    <th><% If FEATURE_I18N Then %><%= T("product_option_volume", Empty) %><% Else %>容量规格<% End If %></th>
                    <th><% If FEATURE_I18N Then %><%= T("quantity", Empty) %><% Else %>数量<% End If %></th>
                    <th><% If FEATURE_I18N Then %><%= T("product_option_label", Empty) %><% Else %>定制备注<% End If %></th>
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
                            If FEATURE_I18N Then
                                customInfo = T("order_ingredients_engraving", Empty) & ": " & HTMLEncode(rsDetails("CustomLabel"))
                            Else
                                customInfo = "刻字: " & HTMLEncode(rsDetails("CustomLabel"))
                            End If
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
                        If FEATURE_I18N Then
                            Response.Write "<tr><td colspan='4' style='text-align:center;color:#999;'>" & T("order_ingredients_no_data", Empty) & "</td></tr>"
                        Else
                            Response.Write "<tr><td colspan='4' style='text-align:center;color:#999;'>无商品数据</td></tr>"
                        End If
                    End If
                    rsDetails.Close
                Else
                    If FEATURE_I18N Then
                        Response.Write "<tr><td colspan='4' style='text-align:center;color:#999;'>" & T("order_ingredients_cannot_read", Empty) & "</td></tr>"
                    Else
                        Response.Write "<tr><td colspan='4' style='text-align:center;color:#999;'>无法读取商品数据</td></tr>"
                    End If
                End If
                Set rsDetails = Nothing
                %>
            </tbody>
        </table>
        
        <div class="section-title">🧪 <% If FEATURE_I18N Then %><%= T("order_ingredients_section_ingredients", Empty) %><% Else %>产品成分清单 / Ingredients<% End If %></div>
        <div style="font-size: 13px; color: #666; margin-bottom: 15px;">
            <% If FEATURE_I18N Then %><%= T("order_ingredients_ing_desc", Empty) %><% Else %>本产品由以下成分配制而成（按字母顺序排列）。如您对某些成分过敏，请在使用前咨询专业人士。<% End If %>
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
        ' 显示数据来源提示
        If Not hasOrderIngredients And uniqueIngredients.Count > 0 Then
        %>
        <div style="font-size: 12px; color: #999; margin-top: 8px; font-style: italic;">
            <% If FEATURE_I18N Then %><%= T("order_ingredients_fallback_note", Empty) %><% Else %>* 成分数据从基香配置链路实时计算生成<% End If %>
        </div>
        <%
        End If
        %>
        <%
            Else
        %>
        <div style="padding: 20px; background: #fff3cd; border: 1px solid #ffc107; color: #856404; text-align: center;">
            ⚠️ <% If FEATURE_I18N Then %><%= T("order_ingredients_empty_warn", Empty) %><% Else %>此订单暂无成分信息记录。可能原因：基香成分尚未配置，或产品未关联配方。请联系技术人员完善基香数据。<% End If %>
        </div>
        <%
            End If
        %>
        
        <div class="notes-section">
            <div class="title">📋 <% If FEATURE_I18N Then %><%= T("order_ingredients_notes_title", Empty) %><% Else %>重要说明<% End If %></div>
            <ul>
                <li><% If FEATURE_I18N Then %><%= T("order_ingredients_notes_1", Empty) %><% Else %>本成分表仅供识别产品成分，用于过敏原检查<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("order_ingredients_notes_2", Empty) %><% Else %>成分列表不包含具体数量或比例信息<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("order_ingredients_notes_3", Empty) %><% Else %>所有成分均符合相关国家标准和行业规范<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("order_ingredients_notes_4", Empty) %><% Else %>产品制作详情（包括香调配比、瓶身规格等）请参考商家生产工单<% End If %></li>
                <li><% If FEATURE_I18N Then %><%= T("order_ingredients_notes_5", Empty) %><% Else %>如有疑问请联系客服部门<% End If %></li>
            </ul>
        </div>
        
        <div class="footer">
            <p><% If FEATURE_I18N Then %><%= T("order_ingredients_footer", Empty) %><% Else %>本文档由系统自动生成，供客户参考<% End If %></p>
            <p><% If FEATURE_I18N Then %><%= T("order_ingredients_label_print_time", Empty) %><% Else %>打印时间<% End If %>: <%= Now() %> | <% If FEATURE_I18N Then %><%= T("order_ingredients_label_order_no", Empty) %><% Else %>订单编号<% End If %>: <%= orderNo %></p>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
