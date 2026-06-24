<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

' CSRF验证
If Not ValidateCSRFToken() Then
    Response.Clear
    Response.ContentType = "text/html"
    Response.Write "<html><body><script>alert('安全验证失败，请刷新页面重试'); window.history.back();</script></body></html>"
    Response.End
End If

Dim productId, topNote, middleNote, baseNote, volumeId, bottleId, customLabel, quantity
Dim sessionId, userId
Dim unitPrice, basePrice, notesPrice, bottlePrice, multiplier
Dim totalPercent
Dim topPercent, middlePercent, basePercent  ' 每种调性的累计比例

' 初始化比例变量
totalPercent = 0
topPercent = 0
middlePercent = 0
basePercent = 0

' 检查是否有POST数据
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' 获取表单数据
    productId = Request.Form("productId")
    ' 注意：现在 topNote, middleNote, baseNote 可能是逗号分隔的列表
    topNote = Request.Form("topNote")
    middleNote = Request.Form("middleNote")
    baseNote = Request.Form("baseNote")
    volumeId = Request.Form("volume")
    bottleId = Request.Form("bottle")
    customLabel = SafeSQL(Request.Form("customLabel"))
    quantity = Request.Form("quantity")
    
    ' 如果Request.Form无法获取数据，尝试使用Request.BinaryRead
    If Request.Form.Count = 0 Then
        Response.Clear
        Response.ContentType = "text/html"
        Response.Write "<html><body><script>alert('表单数据为空，请刷新页面重试'); window.history.back();</script></body></html>"
        Response.End
    End If
    
    ' 验证productId（安全处理，不输出详细调试信息）
    If productId = "" Or Not IsNumeric(productId) Then
        Response.Clear
        Response.ContentType = "text/html"
        Response.Write "<html><body><script>alert('无效的产品参数，请返回重试'); window.history.back();</script></body></html>"
        Response.End
    End If
Else
    ' 非POST请求
    Response.Clear
    Response.ContentType = "text/html"
    Response.Write "<html><body><script>alert('请求方法错误'); window.history.back();</script></body></html>"
    Response.End
End If

' 从SiteSettings获取香调最小比例配置
Dim minTopPercent, minMiddlePercent, minBasePercent
minTopPercent = 10
minMiddlePercent = 10
minBasePercent = 10

Dim rsMinPercent
Set rsMinPercent = ExecuteQuery("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey IN ('MinTopPercent', 'MinMiddlePercent', 'MinBasePercent')")
If Not rsMinPercent Is Nothing Then
    Do While Not rsMinPercent.EOF
        Select Case rsMinPercent("SettingKey")
            Case "MinTopPercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minTopPercent = CInt(rsMinPercent("SettingValue"))
            Case "MinMiddlePercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minMiddlePercent = CInt(rsMinPercent("SettingValue"))
            Case "MinBasePercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minBasePercent = CInt(rsMinPercent("SettingValue"))
        End Select
        rsMinPercent.MoveNext
    Loop
    rsMinPercent.Close
End If
Set rsMinPercent = Nothing

sessionId = Session.SessionID
userId = Session("UserID")

If quantity = "" Or Not IsNumeric(quantity) Then
    quantity = 1
End If
quantity = CInt(quantity)
If quantity < 1 Then quantity = 1

' 获取基础信息和产品类型
Dim rsP
Set rsP = ExecuteQuery("SELECT BasePrice, ProductType, Engravable, EngravingPrice FROM Products WHERE ProductID = " & CInt(productId))
If rsP Is Nothing Or rsP.EOF Then
    Response.Write "<html><body><script>alert('产品不存在'); window.history.back();</script></body></html>"
    Response.End
End If
basePrice = CDbl(rsP("BasePrice"))
productType = LCase(rsP("ProductType") & "")
If productType = "" Then productType = "custom"
Dim productEngravable, productEngravingPrice
productEngravable = False
productEngravingPrice = 0
On Error Resume Next
productEngravable = (rsP("Engravable") = True)
productEngravingPrice = CDbl(rsP("EngravingPrice"))
If Err.Number <> 0 Then productEngravingPrice = 0
On Error GoTo 0
rsP.Close: Set rsP = Nothing

' 计算香调附加费和验证百分比
notesPrice = 0
totalPercent = 0

Function ProcessNotes(notesList, noteTypeLabel)
    Dim currentNotesPrice, nId, nPrice, nPercent, prefix
    Dim typePercent  ' 当前调性的累计比例
    If noteTypeLabel = "前调" Then
        prefix = "top"
    ElseIf noteTypeLabel = "中调" Then
        prefix = "mid"
    Else
        prefix = "base"
    End If
    
    currentNotesPrice = 0
    typePercent = 0
    If notesList <> "" Then
        Dim arr, i
        arr = Split(notesList, ",")
        For i = 0 To UBound(arr)
            nId = Trim(arr(i))
            If IsNumeric(nId) Then
                nPrice = GetScalar("SELECT PriceAddition FROM FragranceNotes WHERE NoteID = " & CInt(nId))
                nPercent = Request.Form("percent_" & prefix & "_" & nId)
                ' 添加空值保护 — DECIMAL类型不能用IsNumeric判断，统一CDbl转换
                If IsNull(nPrice) Or Trim(nPrice & "") = "" Then nPrice = 0 Else nPrice = CDbl(nPrice)
                If IsNumeric(nPercent) Then
                    currentNotesPrice = currentNotesPrice + (CDbl(nPrice) * CDbl(nPercent) / 100)
                    totalPercent = totalPercent + CDbl(nPercent)
                    typePercent = typePercent + CDbl(nPercent)  ' 累加当前调性比例
                End If
            End If
        Next
    End If
    
    ' 根据调性保存到对应变量
    If noteTypeLabel = "前调" Then
        topPercent = typePercent
    ElseIf noteTypeLabel = "中调" Then
        middlePercent = typePercent
    ElseIf noteTypeLabel = "后调" Then
        basePercent = typePercent
    End If
    
    ProcessNotes = currentNotesPrice
End Function

' 获取容量系数（提前获取，standard产品需要用到）
multiplier = 1
If volumeId <> "" And IsNumeric(volumeId) Then
    Dim volMulti
    volMulti = GetScalar("SELECT PriceMultiplier FROM Volumes WHERE VolumeID = " & CInt(volumeId))
    If Not IsNull(volMulti) And volMulti <> "" Then multiplier = CDbl(volMulti)
End If

If productType = "standard" Then
    ' 固定规格价格 - 优先从ProductVolumePrices获取，回退到BasePrice×VolumeMultiplier
    unitPrice = GetScalar("SELECT Price FROM ProductVolumePrices WHERE ProductID = " & CInt(productId) & " AND VolumeID = " & CInt(volumeId))
    If IsNull(unitPrice) Or unitPrice = "" Or CDbl(unitPrice) <= 0 Then
        ' 回退：使用基础价格×容量系数
        unitPrice = basePrice * multiplier
    Else
        unitPrice = CDbl(unitPrice)
    End If
    ' 刻字费用在购物车和结算时单独计算，不加入unitPrice
    totalPercent = 100
    ' 品牌定香产品 - 配比已在后台预设，从数据库读取
    ' 重置表单值，完全从数据库获取
    topNote = ""
    middleNote = ""
    baseNote = ""
    Dim rsSTDNotes
    Set rsSTDNotes = ExecuteQuery("SELECT pn.NoteID, n.NoteType FROM ProductNotes pn LEFT JOIN FragranceNotes n ON pn.NoteID = n.NoteID WHERE pn.ProductID = " & CInt(productId))
    If Not rsSTDNotes Is Nothing Then
        Do While Not rsSTDNotes.EOF
            Dim stdNoteId, stdNoteType
            stdNoteId = rsSTDNotes("NoteID")
            stdNoteType = rsSTDNotes("NoteType") & ""
            If stdNoteType = "前调" Then
                If topNote = "" Then topNote = stdNoteId Else topNote = topNote & "," & stdNoteId
            ElseIf stdNoteType = "中调" Then
                If middleNote = "" Then middleNote = stdNoteId Else middleNote = middleNote & "," & stdNoteId
            ElseIf stdNoteType = "后调" Then
                If baseNote = "" Then baseNote = stdNoteId Else baseNote = baseNote & "," & stdNoteId
            End If
            rsSTDNotes.MoveNext
        Loop
        rsSTDNotes.Close
    End If
    Set rsSTDNotes = Nothing
ElseIf productType = "kol" Then
    ' KOL商品 - 配比已在后台预设，从数据库读取
    totalPercent = 100
    ' 重置表单值，完全从数据库获取（防止表单提交值与DB查询结果重复）
    topNote = ""
    middleNote = ""
    baseNote = ""
    ' 从数据库获取KOL商品的香调配比（通过FragranceNotes获取NoteType）
    Dim rsKOLNotes
    Set rsKOLNotes = ExecuteQuery("SELECT pn.NoteID, n.NoteType FROM ProductNotes pn LEFT JOIN FragranceNotes n ON pn.NoteID = n.NoteID WHERE pn.ProductID = " & CInt(productId))
    If Not rsKOLNotes Is Nothing Then
        Do While Not rsKOLNotes.EOF
            Dim kolNoteId, kolNoteType
            kolNoteId = rsKOLNotes("NoteID")
            kolNoteType = rsKOLNotes("NoteType") & ""
            ' 根据香调类型分类
            If kolNoteType = "前调" Then
                If topNote = "" Then topNote = kolNoteId Else topNote = topNote & "," & kolNoteId
            ElseIf kolNoteType = "中调" Then
                If middleNote = "" Then middleNote = kolNoteId Else middleNote = middleNote & "," & kolNoteId
            ElseIf kolNoteType = "后调" Then
                If baseNote = "" Then baseNote = kolNoteId Else baseNote = baseNote & "," & kolNoteId
            End If
            rsKOLNotes.MoveNext
        Loop
        rsKOLNotes.Close
    End If
    Set rsKOLNotes = Nothing
Else
    ' 定制香水 - 计算香调附加费
    notesPrice = notesPrice + ProcessNotes(topNote, "前调")
    notesPrice = notesPrice + ProcessNotes(middleNote, "中调")
    notesPrice = notesPrice + ProcessNotes(baseNote, "后调")
    
    ' 后端再次验证总百分比（使用容差0.01避免浮点数精度问题）
    If Abs(totalPercent - 100) > 0.01 Then
        Response.Write "<html><body><script>alert('配比总和必须等于100% (当前: " & FormatNumber(totalPercent, 1) & "%)'); window.history.back();</script></body></html>"
        Response.End
    End If
    
    ' 验证每种调性最小比例（使用容差0.01避免浮点数精度问题）
    If topPercent < (minTopPercent - 0.01) Then
        Response.Write "<html><body><script>alert('前调比例不能低于" & minTopPercent & "%，当前为" & FormatNumber(topPercent, 1) & "%'); window.history.back();</script></body></html>"
        Response.End
    End If
    If middlePercent < (minMiddlePercent - 0.01) Then
        Response.Write "<html><body><script>alert('中调比例不能低于" & minMiddlePercent & "%，当前为" & FormatNumber(middlePercent, 1) & "%'); window.history.back();</script></body></html>"
        Response.End
    End If
    If basePercent < (minBasePercent - 0.01) Then
        Response.Write "<html><body><script>alert('后调比例不能低于" & minBasePercent & "%，当前为" & FormatNumber(basePercent, 1) & "%'); window.history.back();</script></body></html>"
        Response.End
    End If
    
    ' ==================== 库存校验 ====================
    ' 检查是否启用库存检查
    Dim enableInventoryCheck
    enableInventoryCheck = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableInventoryCheck'")
    If IsNull(enableInventoryCheck) Then enableInventoryCheck = "1"
    
    If enableInventoryCheck = "1" Then
        ' 检查所选香调的库存
        Dim stockCheckSql, rsStock, insufficientNotes
        insufficientNotes = ""
        
        ' 检查前调库存
        If topNote <> "" Then
            Dim topArr, topId
            topArr = Split(topNote, ",")
            For i = 0 To UBound(topArr)
                topId = Trim(topArr(i))
                If IsNumeric(topId) Then
                    Set rsStock = ExecuteQuery("SELECT fn.NoteName, ni.StockQuantity FROM FragranceNotes fn INNER JOIN NoteInventory ni ON fn.NoteID = ni.NoteID WHERE fn.NoteID = " & CInt(topId))
                    If Not rsStock Is Nothing And Not rsStock.EOF Then
                        If CInt(rsStock("StockQuantity")) < quantity Then
                            If insufficientNotes <> "" Then insufficientNotes = insufficientNotes & ", "
                            insufficientNotes = insufficientNotes & rsStock("NoteName") & "(当前库存:" & rsStock("StockQuantity") & "ml)"
                        End If
                        rsStock.Close
                    End If
                    Set rsStock = Nothing
                End If
            Next
        End If
        
        ' 检查中调库存
        If middleNote <> "" Then
            Dim midArr, midId
            midArr = Split(middleNote, ",")
            For i = 0 To UBound(midArr)
                midId = Trim(midArr(i))
                If IsNumeric(midId) Then
                    Set rsStock = ExecuteQuery("SELECT fn.NoteName, ni.StockQuantity FROM FragranceNotes fn INNER JOIN NoteInventory ni ON fn.NoteID = ni.NoteID WHERE fn.NoteID = " & CInt(midId))
                    If Not rsStock Is Nothing And Not rsStock.EOF Then
                        If CInt(rsStock("StockQuantity")) < quantity Then
                            If insufficientNotes <> "" Then insufficientNotes = insufficientNotes & ", "
                            insufficientNotes = insufficientNotes & rsStock("NoteName") & "(当前库存:" & rsStock("StockQuantity") & "ml)"
                        End If
                        rsStock.Close
                    End If
                    Set rsStock = Nothing
                End If
            Next
        End If
        
        ' 检查后调库存
        If baseNote <> "" Then
            Dim baseArr2, baseId2
            baseArr2 = Split(baseNote, ",")
            For i = 0 To UBound(baseArr2)
                baseId2 = Trim(baseArr2(i))
                If IsNumeric(baseId2) Then
                    Set rsStock = ExecuteQuery("SELECT fn.NoteName, ni.StockQuantity FROM FragranceNotes fn INNER JOIN NoteInventory ni ON fn.NoteID = ni.NoteID WHERE fn.NoteID = " & CInt(baseId2))
                    If Not rsStock Is Nothing And Not rsStock.EOF Then
                        If CInt(rsStock("StockQuantity")) < quantity Then
                            If insufficientNotes <> "" Then insufficientNotes = insufficientNotes & ", "
                            insufficientNotes = insufficientNotes & rsStock("NoteName") & "(当前库存:" & rsStock("StockQuantity") & "ml)"
                        End If
                        rsStock.Close
                    End If
                    Set rsStock = Nothing
                End If
            Next
        End If
        
        ' 如果有库存不足的香调，显示错误
        If insufficientNotes <> "" Then
            Response.Write "<html><body><script>alert('以下香调库存不足，无法添加到购物车：\n" & insufficientNotes & "\n\n请调整购买数量或联系客服咨询。'); window.history.back();</script></body></html>"
            Response.End
        End If
    End If
    ' ==================== 库存校验结束 ====================
End If

' 获取瓶身附加费
bottlePrice = 0
If bottleId <> "" And IsNumeric(bottleId) Then
    Dim btlPrice
    btlPrice = GetScalar("SELECT PriceAddition FROM BottleStyles WHERE BottleID = " & CInt(bottleId))
    If Not IsNull(btlPrice) And btlPrice <> "" Then bottlePrice = CDbl(btlPrice)
End If

' 计算刻字费用（单独记录，不加入unitPrice）
Dim engravingPrice
engravingPrice = 0
If productEngravable And customLabel <> "" Then
    engravingPrice = productEngravingPrice
End If

' 计算单价（不包含刻字费用，刻字费用在购物车和结算时单独计算）
' 品牌定香价格已在上方从ProductVolumePrices获取，不要覆盖
If productType <> "standard" Then
    unitPrice = (basePrice + notesPrice + bottlePrice) * multiplier
End If

' 构建INSERT语句
Dim sql, userIdValue, topNoteValue, middleNoteValue, baseNoteValue, volumeIdValue, bottleIdValue

If userId <> "" And IsNumeric(userId) Then
    userIdValue = userId
Else
    userIdValue = "NULL"
End If

If topNote <> "" And IsNumeric(topNote) Then topNoteValue = topNote Else topNoteValue = "NULL"
If middleNote <> "" And IsNumeric(middleNote) Then middleNoteValue = middleNote Else middleNoteValue = "NULL"
If baseNote <> "" And IsNumeric(baseNote) Then baseNoteValue = baseNote Else baseNoteValue = "NULL"
If volumeId <> "" And IsNumeric(volumeId) Then volumeIdValue = volumeId Else volumeIdValue = "NULL"
If bottleId <> "" And IsNumeric(bottleId) Then bottleIdValue = bottleId Else bottleIdValue = "NULL"

' 为了兼容性，存储每类香调的第一个 ID
Dim firstTop, firstMiddle, firstBase
firstTop = "NULL": firstMiddle = "NULL": firstBase = "NULL"
If topNote <> "" Then firstTop = Split(topNote, ",")(0)
If middleNote <> "" Then firstMiddle = Split(middleNote, ",")(0)
If baseNote <> "" Then firstBase = Split(baseNote, ",")(0)

sql = "INSERT INTO Cart (UserID, SessionID, ProductID, TopNoteID, MiddleNoteID, BaseNoteID, VolumeID, BottleID, CustomLabel, Quantity, UnitPrice, CreatedAt) VALUES (" & _
    userIdValue & ", '" & SafeSQL(sessionId) & "', " & CInt(productId) & ", " & firstTop & ", " & firstMiddle & ", " & firstBase & ", " & _
    volumeIdValue & ", " & bottleIdValue & ", '" & customLabel & "', " & quantity & ", " & CDbl(unitPrice) & ", GETDATE())"

    If ExecuteNonQuery(sql) Then
        ' 获取新插入的 CartID
        Dim cartId, rawCartId
        cartId = 0
        rawCartId = GetScalar("SELECT SCOPE_IDENTITY()")
        If Not IsNull(rawCartId) And rawCartId <> "" Then
            On Error Resume Next
            cartId = CLng(rawCartId)
            If Err.Number <> 0 Then
                Err.Clear
                cartId = 0
            End If
            On Error GoTo 0
        End If
        
        ' 如果 SCOPE_IDENTITY() 失败，尝试通过 SessionID 和 ProductID 获取最新的 CartID
        If cartId = 0 Then
            rawCartId = GetScalar("SELECT MAX(CartID) FROM Cart WHERE SessionID = '" & SafeSQL(sessionId) & "' AND ProductID = " & CInt(productId))
            If Not IsNull(rawCartId) And rawCartId <> "" Then
                On Error Resume Next
                cartId = CLng(rawCartId)
                If Err.Number <> 0 Then
                    Err.Clear
                    cartId = 0
                End If
                On Error GoTo 0
            End If
        End If
        
        If cartId > 0 Then
            Call SaveNoteSelections(cartId, topNote, "前调")
            Call SaveNoteSelections(cartId, middleNote, "中调")
            Call SaveNoteSelections(cartId, baseNote, "后调")
        End If

    ' 检查是否为直接购买
    If Request.Form("buyNow") <> "" Then
        Response.Write "<html><body><script>alert('已添加到购物车！'); window.location.href = '/checkout.asp';</script></body></html>"
    Else
        Response.Write "<html><body><script>alert('已添加到购物车！'); window.history.back();</script></body></html>"
    End If
Else
    Response.Write "<html><body><script>alert('添加失败，请重试'); window.history.back();</script></body></html>"
End If

' 保存详细的香调配比到 CartNoteSelections
Sub SaveNoteSelections(cId, notesList, noteType)
    If notesList <> "" Then
        Dim arr, i, nId, nPercent, prefix
        If noteType = "前调" Then
            prefix = "top"
        ElseIf noteType = "中调" Then
            prefix = "mid"
        Else
            prefix = "base"
        End If
        
        arr = Split(notesList, ",")
        For i = 0 To UBound(arr)
            nId = Trim(arr(i))
            ' 尝试从表单获取百分比，如果没有则从数据库读取（KOL商品）
            nPercent = Request.Form("percent_" & prefix & "_" & nId)
            If nPercent = "" Or Not IsNumeric(nPercent) Then
                ' 尝试从ProductNoteRatios获取KOL商品的预设比例（返回DECIMAL类型需CDbl转换）
                nPercent = GetScalar("SELECT Percentage FROM ProductNoteRatios WHERE ProductID = " & CInt(productId) & " AND NoteID = " & CInt(nId))
                If IsNull(nPercent) Or Trim(nPercent & "") = "" Then nPercent = 0 Else nPercent = CDbl(nPercent)
            End If
            If IsNumeric(nId) Then
                ExecuteNonQuery("INSERT INTO CartNoteSelections (CartID, NoteID, NoteType, Percentage) VALUES (" & _
                    cId & ", " & CInt(nId) & ", '" & noteType & "', " & CInt(CDbl(nPercent)) & ")")
            End If
        Next
    End If
End Sub

Call CloseConnection()
%>
