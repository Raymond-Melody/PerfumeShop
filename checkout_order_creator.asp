<%
' ============================================
' V14.6 结算页 - 订单创建器
' 从 checkout.asp 提取: SyncOrderDetailsAndIngredients, DeductNoteInventory
' 注意: SafeCreatePaymentOrder 已在 includes/payment_handler.asp 中定义
' ============================================

' 同步订单详情和成分数据到OrderDetails和OrderIngredients表
Sub SyncOrderDetailsAndIngredients(orderId, userId, whereClause)
    On Error Resume Next
    
    DebugLog "===== SyncOrderDetailsAndIngredients 开始 ====="
    DebugLog "orderId=" & orderId & ", userId=" & userId
    DebugLog "whereClause=" & whereClause
    
    ' 查询购物车数据
    Dim cartSql, rsCart
    cartSql = "SELECT c.*, p.ProductName, p.ImageURL, " & _
        "tn.NoteName AS TopNoteName, mn.NoteName AS MiddleNoteName, bn.NoteName AS BaseNoteName, " & _
        "v.VolumeName, v.VolumeML, b.BottleName, " & _
        "c.Quantity * c.UnitPrice AS SubTotal " & _
        "FROM ((((((Cart c " & _
        "LEFT JOIN Products p ON c.ProductID = p.ProductID) " & _
        "LEFT JOIN FragranceNotes tn ON c.TopNoteID = tn.NoteID) " & _
        "LEFT JOIN FragranceNotes mn ON c.MiddleNoteID = mn.NoteID) " & _
        "LEFT JOIN FragranceNotes bn ON c.BaseNoteID = bn.NoteID) " & _
        "LEFT JOIN Volumes v ON c.VolumeID = v.VolumeID) " & _
        "LEFT JOIN BottleStyles b ON c.BottleID = b.BottleID) " & _
        "WHERE " & whereClause & " ORDER BY c.CreatedAt DESC"
    
    Set rsCart = ExecuteQuery(cartSql)
    
    If rsCart Is Nothing Then
        Exit Sub
    End If
    
    Dim detailId, insertDetailSql, insertNoteSql, insertIngredientSql
    Dim productId, productName, quantity, unitPrice, subtotal
    Dim topNoteId, middleNoteId, baseNoteId, topNoteName, middleNoteName, baseNoteName
    Dim volumeId, volumeName, volumeML, bottleId, bottleName, customLabel
    Dim cartId, rsCartNotes, rsCartNotes2, processedIngredients
    Dim cNoteId, cNoteType, cPercentage, cNoteId2
    Dim rsBaseIngredients, ingredientsStr, ingredientArr, singleIngredient
    Dim rsDetailId
    Dim productRecipeId, rsRecipeCheck, rsRecipeIngredients, ingName
    
    Do While Not rsCart.EOF
        productId = rsCart("ProductID")
        productName = rsCart("ProductName") & ""
        quantity = CLng(rsCart("Quantity"))
        unitPrice = CDbl(rsCart("UnitPrice"))
        subtotal = CDbl(rsCart("SubTotal"))
        topNoteId = rsCart("TopNoteID") & ""
        middleNoteId = rsCart("MiddleNoteID") & ""
        baseNoteId = rsCart("BaseNoteID") & ""
        topNoteName = rsCart("TopNoteName") & ""
        middleNoteName = rsCart("MiddleNoteName") & ""
        baseNoteName = rsCart("BaseNoteName") & ""
        volumeId = rsCart("VolumeID") & ""
        volumeName = rsCart("VolumeName") & ""
        volumeML = rsCart("VolumeML") & ""
        bottleId = rsCart("BottleID") & ""
        bottleName = rsCart("BottleName") & ""
        customLabel = rsCart("CustomLabel") & ""
        
        ' 插入订单详情记录
        Dim volumeMLValue
        On Error Resume Next
        volumeMLValue = CDbl(volumeML)
        If Err.Number <> 0 Then volumeMLValue = 0 : Err.Clear
        On Error GoTo 0
        insertDetailSql = "INSERT INTO OrderDetails (OrderID, ProductID, ProductName, Quantity, UnitPrice, Subtotal, VolumeName, VolumeML, BottleName, CustomLabel) VALUES (" & _
            orderId & ", " & productId & ", '" & SafeSQL(productName) & "', " & quantity & ", " & unitPrice & ", " & subtotal & ", '" & _
            SafeSQL(volumeName) & "', " & volumeMLValue & ", '" & SafeSQL(bottleName) & "', '" & SafeSQL(customLabel) & "')"
        
        ExecuteNonQuery(insertDetailSql)
        
        ' 获取刚插入的DetailID
        Set rsDetailId = ExecuteQuery("SELECT MAX(DetailID) as maxId FROM OrderDetails WHERE OrderID = " & orderId)
        If Not rsDetailId.EOF Then
            detailId = rsDetailId("maxId")
        Else
            detailId = 0
        End If
        rsDetailId.Close
        Set rsDetailId = Nothing
        
        ' 从CartNoteSelections表读取实际配比数据
        cartId = rsCart("CartID")
        DebugLog "  获取CartID: " & cartId
        Set rsCartNotes = ExecuteQuery("SELECT * FROM CartNoteSelections WHERE CartID = " & cartId)
        If Not rsCartNotes Is Nothing Then
            Do While Not rsCartNotes.EOF
                cNoteId = rsCartNotes("NoteID")
                cNoteType = rsCartNotes("NoteType") & ""
                cPercentage = CDbl(rsCartNotes("Percentage"))
                
                insertNoteSql = "INSERT INTO OrderDetailNoteSelections (DetailID, NoteID, NoteType, Percentage) VALUES (" & _
                    detailId & ", " & cNoteId & ", '" & cNoteType & "', " & CLng(cPercentage) & ")"
                ExecuteNonQuery(insertNoteSql)
                
                rsCartNotes.MoveNext
            Loop
            rsCartNotes.Close
            Set rsCartNotes = Nothing
        End If
        
        ' 从CartNoteSelections表读取实际香调信息并插入到OrderIngredients表
        Set processedIngredients = CreateObject("Scripting.Dictionary")
        
        DebugLog "处理购物车 CartID=" & cartId & ", DetailID=" & detailId
        
        ' 检查产品是否有关联配方
        productRecipeId = 0
        Set rsRecipeCheck = ExecuteQuery("SELECT RecipeID FROM Products WHERE ProductID = " & productId)
        If Not rsRecipeCheck Is Nothing Then
            If Not rsRecipeCheck.EOF Then
                If Not IsNull(rsRecipeCheck("RecipeID")) Then
                    productRecipeId = rsRecipeCheck("RecipeID")
                End If
            End If
            rsRecipeCheck.Close
            Set rsRecipeCheck = Nothing
        End If
        
        If productRecipeId > 0 Then
            ' 从 RecipeIngredients 直接读取成分（高效路径）
            Set rsRecipeIngredients = ExecuteQuery("SELECT IngredientName FROM RecipeIngredients WHERE RecipeID = " & CInt(productRecipeId))
            If Not rsRecipeIngredients Is Nothing Then
                Do While Not rsRecipeIngredients.EOF
                    ingName = Trim(rsRecipeIngredients("IngredientName") & "")
                    If ingName <> "" And Not processedIngredients.Exists(ingName) Then
                        processedIngredients.Add ingName, True
                        insertIngredientSql = "INSERT INTO OrderIngredients (OrderID, DetailID, IngredientName, CreatedAt) VALUES (" & _
                            orderId & ", " & detailId & ", '" & SafeSQL(ingName) & "', GETDATE())"
                        DebugLog "    保存配方成分: " & ingName
                        ExecuteNonQuery(insertIngredientSql)
                    End If
                    rsRecipeIngredients.MoveNext
                Loop
                rsRecipeIngredients.Close
                Set rsRecipeIngredients = Nothing
            End If
        Else
            ' 保持现有逻辑：从 NoteIngredients→BaseNotes 逐层查询
            Set rsCartNotes2 = ExecuteQuery("SELECT c.NoteID, c.NoteType, c.Percentage, n.NoteName FROM CartNoteSelections c LEFT JOIN FragranceNotes n ON c.NoteID = n.NoteID WHERE c.CartID = " & cartId)
            If Not rsCartNotes2 Is Nothing Then
                DebugLog "  找到 " & rsCartNotes2.RecordCount & " 个香调选择"
                
                Do While Not rsCartNotes2.EOF
                    cNoteId2 = rsCartNotes2("NoteID")
                    
                    DebugLog "  处理香调 NoteID=" & cNoteId2
                    
                    Set rsBaseIngredients = ExecuteQuery("SELECT b.Ingredients FROM NoteIngredients ni LEFT JOIN BaseNotes b ON ni.BaseNoteID = b.BaseNoteID WHERE ni.NoteID = " & cNoteId2 & " AND b.Ingredients IS NOT NULL AND b.Ingredients <> ''")
                    
                    If Not rsBaseIngredients Is Nothing Then
                        If Not rsBaseIngredients.EOF Then
                            DebugLog "    找到 " & rsBaseIngredients.RecordCount & " 个基香"
                            
                            Do While Not rsBaseIngredients.EOF
                                ingredientsStr = rsBaseIngredients("Ingredients") & ""
                                DebugLog "    基香成分: " & ingredientsStr
                                
                                If ingredientsStr <> "" Then
                                    Dim ingredientDict, ingKey
                                    Set ingredientDict = SplitIngredientsUniversal(ingredientsStr)
                                    
                                    For Each ingKey In ingredientDict.Keys
                                        If Not processedIngredients.Exists(ingKey) Then
                                            processedIngredients.Add ingKey, True
                                            insertIngredientSql = "INSERT INTO OrderIngredients (OrderID, DetailID, IngredientName, CreatedAt) VALUES (" & _
                                                orderId & ", " & detailId & ", '" & SafeSQL(ingKey) & "', GETDATE())"
                                            DebugLog "    执行SQL: " & insertIngredientSql
                                            ExecuteNonQuery(insertIngredientSql)
                                        End If
                                    Next
                                    Set ingredientDict = Nothing
                                End If
                                rsBaseIngredients.MoveNext
                            Loop
                        Else
                            DebugLog "    该香调没有基香成分"
                        End If
                        rsBaseIngredients.Close
                        Set rsBaseIngredients = Nothing
                    Else
                        DebugLog "    查询基香失败"
                    End If
                    
                    rsCartNotes2.MoveNext
                Loop
                
                DebugLog "  完成处理，共保存 " & processedIngredients.Count & " 个成分"
                
                rsCartNotes2.Close
                Set rsCartNotes2 = Nothing
            Else
                DebugLog "  没有找到香调选择"
            End If
        End If
        
        ' ===== 品牌定香商品的BaseIngredients处理 =====
        Dim rsProductIngr, productBaseIngr, fixedIngDict, fixedIngKey, productIngrType
        Set rsProductIngr = ExecuteQuery("SELECT BaseIngredients, ProductType FROM Products WHERE ProductID = " & productId)
        If Not rsProductIngr Is Nothing Then
            If Not rsProductIngr.EOF Then
                productBaseIngr = Trim(rsProductIngr("BaseIngredients") & "")
                productIngrType = LCase(Trim(rsProductIngr("ProductType") & ""))
                If productBaseIngr <> "" And (productIngrType = "custom" Or productIngrType = "kol") Then
                    DebugLog "  品牌定香商品BaseIngredients: " & productBaseIngr
                    
                    If processedIngredients Is Nothing Then
                        Set processedIngredients = CreateObject("Scripting.Dictionary")
                    End If
                    
                    Set fixedIngDict = SplitIngredientsUniversal(productBaseIngr)
                    For Each fixedIngKey In fixedIngDict.Keys
                        If Not processedIngredients.Exists(fixedIngKey) Then
                            processedIngredients.Add fixedIngKey, True
                            insertIngredientSql = "INSERT INTO OrderIngredients (OrderID, DetailID, IngredientName, CreatedAt) VALUES (" & _
                                orderId & ", " & detailId & ", '" & SafeSQL(fixedIngKey) & "', GETDATE())"
                            DebugLog "    保存品牌定香成分: " & fixedIngKey
                            ExecuteNonQuery(insertIngredientSql)
                        End If
                    Next
                    Set fixedIngDict = Nothing
                    DebugLog "  品牌定香成分处理完成"
                End If
            End If
            rsProductIngr.Close
            Set rsProductIngr = Nothing
        End If
        
        If Err.Number <> 0 Then
            DebugLog "  错误: " & Err.Description & " (" & Err.Number & ")"
            Err.Clear
        End If
        
        ' ==================== 库存扣减 ====================
        Dim enableInventoryCheck
        enableInventoryCheck = GetScalar("SELECT SettingValue FROM SiteSettings WHERE SettingKey = 'EnableInventoryCheck'")
        If IsNull(enableInventoryCheck) Then enableInventoryCheck = "1"
        
        If enableInventoryCheck = "1" Then
            Call DeductNoteInventory(topNoteId, quantity, orderId, productName)
            Call DeductNoteInventory(middleNoteId, quantity, orderId, productName)
            Call DeductNoteInventory(baseNoteId, quantity, orderId, productName)
        End If
        ' ==================== 库存扣减结束 ====================
        
        rsCart.MoveNext
    Loop
    
    rsCart.Close
    Set rsCart = Nothing
    
    DebugLog "===== SyncOrderDetailsAndIngredients 结束 ====="
    
    On Error GoTo 0
End Sub

' 扣减香调库存的辅助函数
Sub DeductNoteInventory(noteIdList, quantity, orderId, productName)
    On Error Resume Next
    
    If noteIdList = "" Then Exit Sub
    
    Dim arr, i, nId
    arr = Split(noteIdList & "", ",")
    
    For i = 0 To UBound(arr)
        nId = Trim(arr(i))
        If IsNumeric(nId) Then
            Dim deductSql
            deductSql = "UPDATE NoteInventory SET " & _
                "StockQuantity = StockQuantity - " & CInt(quantity) & ", " & _
                "UpdatedAt = GETDATE() " & _
                "WHERE NoteID = " & CInt(nId) & " AND StockQuantity >= " & CInt(quantity)
            ExecuteNonQuery(deductSql)
            
            Dim transSql
            transSql = "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, ReferenceOrderID, ReferenceType, Notes, CreatedAt) VALUES (" & _
                CInt(nId) & ", -" & CInt(quantity) & ", '订单消耗', " & orderId & ", '订单" & orderId & "-" & SafeSQL(productName) & "', '', GETDATE())"
            ExecuteNonQuery(transSql)
        End If
    Next
    
    On Error GoTo 0
End Sub
%>