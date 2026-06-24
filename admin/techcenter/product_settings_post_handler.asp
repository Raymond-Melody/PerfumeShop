<%
' ============================================
' V14.6 产品设置 - POST请求处理器
' 从 product_settings.asp 提取
' ============================================

    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim postAction
    postAction = Request.Form("action")
    
    ' ========== 标准香氛(standard)产品拦截：禁止在技术中心新增/编辑 ==========
        If postAction = "add_product" And LCase(SafeSQL(Request.Form("productType"))) = "standard" Then
            Response.Write "<script>alert('标准香氛产品请前往【采购管理 → 品牌定香】模块进行创建和管理');history.back();</script>"
            Response.End
        End If
        If postAction = "edit_product" Then
            Dim editPid, editPType
            editPid = Request.Form("productId")
            If IsNumeric(editPid) Then
                Dim rsEditCheck
                Set rsEditCheck = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = " & CLng(editPid))
                If Not rsEditCheck Is Nothing And Not rsEditCheck.EOF Then
                    editPType = LCase(rsEditCheck("ProductType") & "")
                    If editPType = "standard" Then
                        rsEditCheck.Close
                        Set rsEditCheck = Nothing
                        Response.Write "<script>alert('标准香氛产品请前往【采购管理 → 品牌定香】模块进行编辑管理');history.back();</script>"
                        Response.End
                    End If
                End If
                If Not rsEditCheck Is Nothing Then
                    rsEditCheck.Close
                    Set rsEditCheck = Nothing
                End If
            End If
        End If
        
        If postAction = "add_product" Or postAction = "edit_product" Then
        Dim productName, description, basePrice, productType, baseIngredients, reviewStatus, isActive, imageURL
        Dim engravable, engravingPrice, kolId, recipeId
        Dim i  ' 循环计数器变量
        
        productName = SafeSQL(Request.Form("productName"))
        description = SafeSQL(Request.Form("description"))
        basePrice = Request.Form("basePrice")
        If basePrice = "" Or Not IsNumeric(basePrice) Then basePrice = 0
        productType = LCase(SafeSQL(Request.Form("productType")))
        
        ' ========== 价格验证 ==========
        Dim priceErrorMsg
        priceErrorMsg = ""
        
        ' standard类型必须BasePrice > 0，custom类型可为0
        If productType = "standard" Then
            If SafeNum(basePrice) <= 0 Then
                priceErrorMsg = "Fixed类型产品的基础价格必须大于0"
            End If
        End If
        
        ' 容量价格非负检查
        If priceErrorMsg = "" Then
            Dim checkVolumes, checkVolArr, checkVolItem, checkVolPrice
            checkVolumes = Request.Form("selectedVolumes")
            If checkVolumes <> "" Then
                checkVolArr = Split(checkVolumes, ",")
                For i = 0 To UBound(checkVolArr)
                    checkVolItem = Trim(checkVolArr(i))
                    If checkVolItem <> "" And IsNumeric(checkVolItem) Then
                        checkVolPrice = Request.Form("volumePrice_" & checkVolItem)
                        If checkVolPrice <> "" And IsNumeric(checkVolPrice) Then
                            If SafeNum(checkVolPrice) < 0 Then
                                priceErrorMsg = "容量价格不能为负数"
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
        End If
        
        ' 如果价格验证失败，显示错误并停止处理
        If priceErrorMsg <> "" Then
            Response.Write "<script>alert('价格验证失败：" & Replace(priceErrorMsg, "'", "\'") & "');</script>"
        Else
        baseIngredients = SafeSQL(Request.Form("baseIngredients"))
        reviewStatus = SafeSQL(Request.Form("reviewStatus"))
        If reviewStatus = "" Then reviewStatus = "Pending"
        isActive = Request.Form("isActive")
        If isActive = "" Then isActive = 1
        imageURL = SafeSQL(Request.Form("imageURL"))
        If imageURL = "" Then imageURL = "/images/default-product.svg"
        
        ' 刻字配置
        If Len(Request.Form("engravable")) > 0 Then
            engravable = 1
        Else
            engravable = 0
        End If
        engravingPrice = Request.Form("engravingPrice")
        If engravingPrice = "" Or Not IsNumeric(engravingPrice) Then engravingPrice = 0
        
        ' KOL ID
        kolId = Request.Form("kolId")
        If kolId = "" Or Not IsNumeric(kolId) Then kolId = 0
        
        ' 关联配方ID
        recipeId = Request.Form("recipeId")
        If recipeId = "" Or Not IsNumeric(recipeId) Then recipeId = 0
        
        ' KOL和Custom类型不关联配方，强制为NULL
        If productType = "kol" Or productType = "custom" Then recipeId = 0
        
        Dim productId, isNewProduct
        isNewProduct = (postAction = "add_product")
        
        If isNewProduct Then
            ' 添加新产品
            Dim addProductSql
            addProductSql = "INSERT INTO Products (ProductName, Description, BasePrice, ProductType, BaseIngredients, ReviewStatus, IsActive, ImageURL, KOLID, Engravable, EngravingPrice, RecipeID, CreatedAt) VALUES ('" & _
                            productName & "', '" & description & "', " & SafeNum(basePrice) & ", '" & productType & "', '" & baseIngredients & "', '" & reviewStatus & "', " & CInt(isActive) & ", '" & imageURL & "', " & CLng(kolId) & ", " & CInt(engravable) & ", " & SafeNum(engravingPrice) & ", " & IIf(CLng(recipeId) > 0, CLng(recipeId), "NULL") & ", GETDATE())"
            If ExecuteNonQuery(addProductSql) Then
                ' 获取新插入的产品ID
                Dim rsNewId
                Set rsNewId = ExecuteQuery("SELECT MAX(ProductID) AS NewID FROM Products WHERE ProductName = '" & productName & "'")
                If Not rsNewId Is Nothing Then
                    If Not rsNewId.EOF Then
                        productId = rsNewId("NewID")
                    End If
                    rsNewId.Close
                End If
                Set rsNewId = Nothing
            Else
                Response.Write "<script>alert('添加失败：" & Replace(Session("LastDBError"), "'", "\'") & "');</script>"
            End If
        Else
            ' 编辑产品
            productId = Request.Form("productId")
            If productId = "" Or Not IsNumeric(productId) Then productId = 0
            Dim editProductSql
            editProductSql = "UPDATE Products SET ProductName = '" & productName & "', Description = '" & description & "', " & _
                             "BasePrice = " & SafeNum(basePrice) & ", ProductType = '" & productType & "', " & _
                             "BaseIngredients = '" & baseIngredients & "', ReviewStatus = '" & reviewStatus & "', " & _
                             "IsActive = " & CInt(isActive) & ", ImageURL = '" & imageURL & "', " & _
                             "KOLID = " & CLng(kolId) & ", Engravable = " & CInt(engravable) & ", " & _
                             "EngravingPrice = " & SafeNum(engravingPrice) & ", RecipeID = " & IIf(CLng(recipeId) > 0, CLng(recipeId), "NULL") & " WHERE ProductID = " & CLng(productId)
            If ExecuteNonQuery(editProductSql) Then
                ' 继续处理关联数据
            Else
                Response.Write "<script>alert('更新失败');</script>"
                productId = 0
            End If
        End If
        
        ' 处理关联数据（新增和编辑都需要）
        If productId > 0 Then
            Dim noteId, volumeId, bottleId
            
            ' ========== KOL类型配比校验 ==========
            Dim ratioErrorMsg
            ratioErrorMsg = ""
            
            If productType = "kol" Then
                ' 从SiteSettings表读取最小比例配置
                Dim minTopRatio, minMiddleRatio, minBaseRatio
                Dim rsMinRatio
                Set rsMinRatio = ExecuteQuery("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey IN ('MinTopPercent', 'MinMiddlePercent', 'MinBasePercent')")
                If Not rsMinRatio Is Nothing Then
                    Do While Not rsMinRatio.EOF
                        If rsMinRatio("SettingKey") = "MinTopPercent" Then
                            minTopRatio = SafeNum(rsMinRatio("SettingValue"))
                        ElseIf rsMinRatio("SettingKey") = "MinMiddlePercent" Then
                            minMiddleRatio = SafeNum(rsMinRatio("SettingValue"))
                        ElseIf rsMinRatio("SettingKey") = "MinBasePercent" Then
                            minBaseRatio = SafeNum(rsMinRatio("SettingValue"))
                        End If
                        rsMinRatio.MoveNext
                    Loop
                    rsMinRatio.Close
                End If
                Set rsMinRatio = Nothing
                
                ' 设置默认值（如果配置不存在）
                If minTopRatio = 0 Then minTopRatio = 10
                If minMiddleRatio = 0 Then minMiddleRatio = 10
                If minBaseRatio = 0 Then minBaseRatio = 10
                
                ' 计算各调性比例
                Dim totalTopPercent, totalMiddlePercent, totalBasePercent, totalPercent
                totalTopPercent = 0
                totalMiddlePercent = 0
                totalBasePercent = 0
                
                Dim selectedNotes, notePercent
                ' KOL类型不再关联配方，直接使用表单提交的配比进行校验
                selectedNotes = Request.Form("selectedNotes")
                If selectedNotes <> "" Then
                    Dim noteArr, noteItem
                    noteArr = Split(selectedNotes, ",")
                    For i = 0 To UBound(noteArr)
                        noteItem = Trim(noteArr(i))
                        If noteItem <> "" And IsNumeric(noteItem) Then
                            notePercent = Request.Form("notePercent_" & noteItem)
                            If notePercent = "" Or Not IsNumeric(notePercent) Then notePercent = 0
                            notePercent = SafeNum(notePercent)
                            
                            ' 获取该香调的调性类型
                            Dim rsNoteType
                            Set rsNoteType = ExecuteQuery("SELECT NoteType FROM FragranceNotes WHERE NoteID = " & CLng(noteItem))
                            If Not rsNoteType Is Nothing Then
                                If Not rsNoteType.EOF Then
                                    If rsNoteType("NoteType") = "前调" Then
                                        totalTopPercent = totalTopPercent + notePercent
                                    ElseIf rsNoteType("NoteType") = "中调" Then
                                        totalMiddlePercent = totalMiddlePercent + notePercent
                                    ElseIf rsNoteType("NoteType") = "后调" Then
                                        totalBasePercent = totalBasePercent + notePercent
                                    End If
                                End If
                                rsNoteType.Close
                            End If
                            Set rsNoteType = Nothing
                        End If
                    Next
                End If
                
                totalPercent = totalTopPercent + totalMiddlePercent + totalBasePercent
                
                ' 校验规则（使用容差0.01避免浮点数精度问题）
                If totalTopPercent < (minTopRatio - 0.01) Then
                    ratioErrorMsg = "前调比例不能低于" & minTopRatio & "%，当前为" & FormatNumber(totalTopPercent, 1) & "%"
                ElseIf totalMiddlePercent < (minMiddleRatio - 0.01) Then
                    ratioErrorMsg = "中调比例不能低于" & minMiddleRatio & "%，当前为" & FormatNumber(totalMiddlePercent, 1) & "%"
                ElseIf totalBasePercent < (minBaseRatio - 0.01) Then
                    ratioErrorMsg = "后调比例不能低于" & minBaseRatio & "%，当前为" & FormatNumber(totalBasePercent, 1) & "%"
                ElseIf Abs(totalPercent - 100) > 0.01 Then
                    ratioErrorMsg = "香调配比总和必须等于100%，当前为" & FormatNumber(totalPercent, 1) & "%"
                End If
            End If
            
            ' 如果配比校验失败，回滚并显示错误
            If ratioErrorMsg <> "" Then
                ' 删除已创建的产品
                Call ExecuteNonQuery("DELETE FROM Products WHERE ProductID = " & CLng(productId))
                Response.Write "<script>alert('配比校验失败：" & Replace(ratioErrorMsg, "'", "\'") & "');</script>"
            Else
                ' 校验通过，保存关联数据
                ' 删除旧的关联数据
                Call ExecuteNonQuery("DELETE FROM ProductNotes WHERE ProductID = " & CLng(productId))
                Call ExecuteNonQuery("DELETE FROM ProductVolumePrices WHERE ProductID = " & CLng(productId))
                Call ExecuteNonQuery("DELETE FROM ProductNoteRatios WHERE ProductID = " & CLng(productId))
                
                ' 保存香调配置（仅Custom和KOL类型）
                If productType = "custom" Or productType = "kol" Then
                    selectedNotes = Request.Form("selectedNotes")
                    If selectedNotes <> "" Then
                        noteArr = Split(selectedNotes, ",")
                        For i = 0 To UBound(noteArr)
                            noteItem = Trim(noteArr(i))
                            If noteItem <> "" Then
                                If IsNumeric(noteItem) Then
                                    Call ExecuteNonQuery("INSERT INTO ProductNotes (ProductID, NoteID) VALUES (" & CLng(productId) & ", " & CLng(noteItem) & ")")
                                End If
                            End If
                        Next
                    End If
                    
                    ' 对于KOL类型，保存预设比例
                    If productType = "kol" Then
                        ' KOL类型不再关联配方，直接使用表单提交的配比
                        selectedNotes = Request.Form("selectedNotes")
                        If selectedNotes <> "" Then
                            noteArr = Split(selectedNotes, ",")
                            For i = 0 To UBound(noteArr)
                                noteItem = Trim(noteArr(i))
                                If noteItem <> "" And IsNumeric(noteItem) Then
                                    notePercent = Request.Form("notePercent_" & noteItem)
                                    If notePercent = "" Or Not IsNumeric(notePercent) Then notePercent = 0
                                    If CInt(notePercent) > 0 Then
                                        Call ExecuteNonQuery("INSERT INTO ProductNoteRatios (ProductID, NoteID, Percentage) VALUES (" & CLng(productId) & ", " & CLng(noteItem) & ", " & CInt(notePercent) & ")")
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            
                ' 保存容量配置
                Dim selectedVolumes
                selectedVolumes = Request.Form("selectedVolumes")
                If selectedVolumes <> "" Then
                    Dim volArr, volItem, volPrice
                    volArr = Split(selectedVolumes, ",")
                    For i = 0 To UBound(volArr)
                        volItem = Trim(volArr(i))
                        If volItem <> "" Then
                            If IsNumeric(volItem) Then
                                ' 获取容量价格（Fixed类型有自定义价格）
                                volPrice = Request.Form("volumePrice_" & volItem)
                                If volPrice = "" Or Not IsNumeric(volPrice) Then
                                    ' 使用默认价格计算
                                    Dim rsVol
                                    Set rsVol = ExecuteQuery("SELECT PriceMultiplier FROM Volumes WHERE VolumeID = " & CLng(volItem))
                                    If Not rsVol Is Nothing Then
                                        If Not rsVol.EOF Then
                                            volPrice = SafeNum(basePrice) * SafeNum(rsVol("PriceMultiplier"))
                                        End If
                                        rsVol.Close
                                    End If
                                    Set rsVol = Nothing
                                End If
                                Call ExecuteNonQuery("INSERT INTO ProductVolumePrices (ProductID, VolumeID, Price) VALUES (" & CLng(productId) & ", " & CLng(volItem) & ", " & SafeNum(volPrice) & ")")
                            End If
                        End If
                    Next
                End If
                
                ' 保存瓶型配置（仅Custom和KOL类型）
                If productType = "custom" Or productType = "kol" Then
                    Dim selectedBottles
                    selectedBottles = Request.Form("selectedBottles")
                    ' 先删除旧的瓶型关联
                    Call ExecuteNonQuery("DELETE FROM ProductBottleStyles WHERE ProductID = " & CLng(productId))
                    
                    If selectedBottles <> "" Then
                        Dim bottleArr, bottleItem
                        bottleArr = Split(selectedBottles, ",")
                        For i = 0 To UBound(bottleArr)
                            bottleItem = Trim(bottleArr(i))
                            If bottleItem <> "" And IsNumeric(bottleItem) Then
                                ' CustomPrice 设为 NULL，前端将使用 BottleStyles.PriceAddition
                                Call ExecuteNonQuery("INSERT INTO ProductBottleStyles (ProductID, BottleID, CustomPrice) VALUES (" & CLng(productId) & ", " & CLng(bottleItem) & ", NULL)")
                            End If
                        Next
                    End If
                End If
                
                If isNewProduct Then
                    Response.Redirect "product_settings.asp?tab=products&msg=" & Server.URLEncode("产品添加成功")
                Else
                    Response.Redirect "product_settings.asp?tab=products&msg=" & Server.URLEncode("产品更新成功")
                End If
            End If
        End If
        End If ' 关闭 priceErrorMsg Else 块
    ElseIf postAction = "delete_product" Then
        ' 软删除产品 - 仅TECH_MANAGER可操作
        If isManager Then
            Dim deleteProductId, deleteProductSql
            deleteProductId = Request.Form("productId")
            If deleteProductId = "" Or Not IsNumeric(deleteProductId) Then deleteProductId = 0
            deleteProductSql = "UPDATE Products SET IsActive = 0 WHERE ProductID = " & CLng(deleteProductId)
            If ExecuteNonQuery(deleteProductSql) Then
                Response.Redirect "product_settings.asp?tab=products&msg=" & Server.URLEncode("产品已禁用")
            Else
                Response.Write "<script>alert('操作失败');</script>"
            End If
        Else
            Response.Write "<script>alert('权限不足');</script>"
        End If
    ElseIf postAction = "restore_product" Then
        ' 恢复产品 - 仅TECH_MANAGER可操作
        If isManager Then
            Dim restoreProductId, restoreProductSql
            restoreProductId = Request.Form("productId")
            If restoreProductId = "" Or Not IsNumeric(restoreProductId) Then restoreProductId = 0
            restoreProductSql = "UPDATE Products SET IsActive = 1 WHERE ProductID = " & CLng(restoreProductId)
            If ExecuteNonQuery(restoreProductSql) Then
                Response.Redirect "product_settings.asp?tab=products&msg=" & Server.URLEncode("产品已恢复")
            Else
                Response.Write "<script>alert('操作失败');</script>"
            End If
        Else
            Response.Write "<script>alert('权限不足');</script>"
        End If
    ElseIf postAction = "edit_type" Then
        ' 编辑产品类型 - 仅TECH_MANAGER可操作
        If isManager Then
            Dim configId, displayName, navName, typeDescription, icon, requiresReview, requiresRatio, displayOrder, typeIsActive
            
            configId = Request.Form("configId")
            If configId = "" Or Not IsNumeric(configId) Then configId = 0
            displayName = SafeSQL(Request.Form("displayName"))
            navName = SafeSQL(Request.Form("navName"))
            typeDescription = SafeSQL(Request.Form("description"))
            icon = SafeSQL(Request.Form("icon"))
            
            ' 复选框处理
            If Len(Request.Form("requiresReview")) > 0 Then
                requiresReview = 1
            Else
                requiresReview = 0
            End If
            
            If Len(Request.Form("requiresRatio")) > 0 Then
                requiresRatio = 1
            Else
                requiresRatio = 0
            End If
            
            If Len(Request.Form("isActive")) > 0 Then
                typeIsActive = 1
            Else
                typeIsActive = 0
            End If
            
            displayOrder = Request.Form("displayOrder")
            If displayOrder = "" Or Not IsNumeric(displayOrder) Then displayOrder = 0
            
            Dim editTypeSql
            editTypeSql = "UPDATE ProductTypeConfig SET DisplayName = '" & displayName & "', NavName = '" & navName & "', " & _
                          "Description = '" & typeDescription & "', Icon = '" & icon & "', " & _
                          "RequiresReview = " & requiresReview & ", RequiresRatio = " & requiresRatio & ", " & _
                          "DisplayOrder = " & CInt(displayOrder) & ", IsActive = " & typeIsActive & " WHERE ConfigID = " & CLng(configId)
            If ExecuteNonQuery(editTypeSql) Then
                Response.Redirect "product_settings.asp?tab=types&msg=" & Server.URLEncode("类型更新成功")
            Else
                Response.Write "<script>alert('更新失败');</script>"
            End If
        Else
            Response.Write "<script>alert('权限不足');</script>"
        End If
    ElseIf postAction = "save_ratio_settings" Then
        ' 保存香调配比参数设置 - 仅TECH_MANAGER可操作
        If isManager Then
            Dim saveMinTop, saveMinMiddle, saveMinBase
            
            saveMinTop = Trim(Request.Form("minTopPercent"))
            saveMinMiddle = Trim(Request.Form("minMiddlePercent"))
            saveMinBase = Trim(Request.Form("minBasePercent"))
            
            ' 验证输入值 - 分步验证避免VBScript Or不短路问题
            If Not IsNumeric(saveMinTop) Then
                saveMinTop = "10"
            ElseIf CInt(saveMinTop) < 0 Or CInt(saveMinTop) > 100 Then
                saveMinTop = "10"
            End If

            If Not IsNumeric(saveMinMiddle) Then
                saveMinMiddle = "10"
            ElseIf CInt(saveMinMiddle) < 0 Or CInt(saveMinMiddle) > 100 Then
                saveMinMiddle = "10"
            End If

            If Not IsNumeric(saveMinBase) Then
                saveMinBase = "10"
            ElseIf CInt(saveMinBase) < 0 Or CInt(saveMinBase) > 100 Then
                saveMinBase = "10"
            End If
            
            ' 更新SiteSettings表（使用Upsert逻辑：先检查是否存在，不存在则INSERT，存在则UPDATE）
            Dim resultMinTop, resultMinMiddle, resultMinBase
            Dim existCount
            
            ' 处理 MinTopPercent
            existCount = GetScalar("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = 'MinTopPercent'")
            If CInt(existCount) > 0 Then
                resultMinTop = ExecuteNonQuery("UPDATE SiteSettings SET SettingValue = '" & saveMinTop & "', UpdatedAt = GETDATE() WHERE SettingKey = 'MinTopPercent'")
            Else
                resultMinTop = ExecuteNonQuery("INSERT INTO SiteSettings (SettingKey, SettingValue, UpdatedAt) VALUES ('MinTopPercent', '" & saveMinTop & "', GETDATE())")
            End If
            
            ' 处理 MinMiddlePercent
            existCount = GetScalar("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = 'MinMiddlePercent'")
            If CInt(existCount) > 0 Then
                resultMinMiddle = ExecuteNonQuery("UPDATE SiteSettings SET SettingValue = '" & saveMinMiddle & "', UpdatedAt = GETDATE() WHERE SettingKey = 'MinMiddlePercent'")
            Else
                resultMinMiddle = ExecuteNonQuery("INSERT INTO SiteSettings (SettingKey, SettingValue, UpdatedAt) VALUES ('MinMiddlePercent', '" & saveMinMiddle & "', GETDATE())")
            End If
            
            ' 处理 MinBasePercent
            existCount = GetScalar("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = 'MinBasePercent'")
            If CInt(existCount) > 0 Then
                resultMinBase = ExecuteNonQuery("UPDATE SiteSettings SET SettingValue = '" & saveMinBase & "', UpdatedAt = GETDATE() WHERE SettingKey = 'MinBasePercent'")
            Else
                resultMinBase = ExecuteNonQuery("INSERT INTO SiteSettings (SettingKey, SettingValue, UpdatedAt) VALUES ('MinBasePercent', '" & saveMinBase & "', GETDATE())")
            End If
            
            If resultMinTop And resultMinMiddle And resultMinBase Then
                Response.Redirect "product_settings.asp?tab=ratio&msg=" & Server.URLEncode("香调配比参数已保存")
            Else
                Response.Write "<script>alert('保存失败：" & Replace(Session("LastDBError"), "'", "\'") & "');</script>"
            End If
        Else
            Response.Write "<script>alert('权限不足');</script>"
        End If
    End If
End If
%>
