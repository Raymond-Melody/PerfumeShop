<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<!--#include file="../../includes/product_type_utils.asp"-->
<%
Call OpenConnection()

' ========== SafeNum/SafeDiv 函数 ==========
Function SafeNum(val)
    On Error Resume Next
    If IsNull(val) Or IsEmpty(val) Or val = "" Then
        SafeNum = 0
    Else
        SafeNum = CDbl(val)
        If Err.Number <> 0 Then
            SafeNum = 0
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

Function SafeDiv(numerator, denominator)
    On Error Resume Next
    If IsNull(denominator) Or denominator = "" Then
        SafeDiv = 0
    ElseIf Not IsNumeric(denominator) Then
        SafeDiv = 0
    ElseIf CDbl(denominator) = 0 Then
        SafeDiv = 0
    Else
        SafeDiv = CDbl(numerator) / CDbl(denominator)
    End If
    On Error GoTo 0
End Function

' ========== 预加载产品类型数据 ==========
Dim allProductTypes
allProductTypes = GetAllProductTypes()

' ========== 获取当前Tab ==========
Dim currentTab
currentTab = Request.QueryString("tab")
If currentTab = "" Then currentTab = "products"

' ========== 处理产品表单提交 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim postAction
    postAction = Request.Form("action")
    
    ' ========== 品牌定香(Fixed)产品拦截：禁止在技术中心新增/编辑 ==========
        If postAction = "add_product" And SafeSQL(Request.Form("productType")) = "Fixed" Then
            Response.Write "<script>alert('品牌定香产品请前往【采购管理 → 品牌定香】模块进行创建和管理');history.back();</script>"
            Response.End
        End If
        If postAction = "edit_product" Then
            Dim editPid, editPType
            editPid = Request.Form("productId")
            If IsNumeric(editPid) Then
                Dim rsEditCheck
                Set rsEditCheck = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = " & CLng(editPid))
                If Not rsEditCheck Is Nothing And Not rsEditCheck.EOF Then
                    editPType = rsEditCheck("ProductType") & ""
                    If editPType = "Fixed" Then
                        rsEditCheck.Close
                        Set rsEditCheck = Nothing
                        Response.Write "<script>alert('品牌定香产品请前往【采购管理 → 品牌定香】模块进行编辑管理');history.back();</script>"
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
        productType = SafeSQL(Request.Form("productType"))
        
        ' ========== 价格验证 ==========
        Dim priceErrorMsg
        priceErrorMsg = ""
        
        ' Fixed类型必须BasePrice > 0，Custom类型可为0
        If productType = "Fixed" Then
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
        If productType = "KOL" Or productType = "Custom" Then recipeId = 0
        
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
            
            If productType = "KOL" Then
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
                
                ' 校验规则
                If totalTopPercent < minTopRatio Then
                    ratioErrorMsg = "前调比例不能低于" & minTopRatio & "%，当前为" & totalTopPercent & "%"
                ElseIf totalMiddlePercent < minMiddleRatio Then
                    ratioErrorMsg = "中调比例不能低于" & minMiddleRatio & "%，当前为" & totalMiddlePercent & "%"
                ElseIf totalBasePercent < minBaseRatio Then
                    ratioErrorMsg = "后调比例不能低于" & minBaseRatio & "%，当前为" & totalBasePercent & "%"
                ElseIf totalPercent <> 100 Then
                    ratioErrorMsg = "香调配比总和必须等于100%，当前为" & totalPercent & "%"
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
                If productType = "Custom" Or productType = "KOL" Then
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
                    If productType = "KOL" Then
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
                If productType = "Custom" Or productType = "KOL" Then
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
        End If  ' 关闭 If priceErrorMsg <> "" Then
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

' ========== 获取产品筛选参数 ==========
Dim filterProductType, productSearch
filterProductType = Request.QueryString("product_type")
productSearch = SafeSQL(Request.QueryString("product_search"))

' ========== 构建产品查询条件 ==========
Dim productWhereClause
productWhereClause = ""

If filterProductType <> "" Then
    If productWhereClause <> "" Then productWhereClause = productWhereClause & " AND "
    productWhereClause = productWhereClause & "ProductType = '" & filterProductType & "'"
End If

If productSearch <> "" Then
    If productWhereClause <> "" Then productWhereClause = productWhereClause & " AND "
    productWhereClause = productWhereClause & "ProductName LIKE '%" & productSearch & "%'"
End If

If productWhereClause <> "" Then
    productWhereClause = "WHERE " & productWhereClause
End If

' ========== 获取产品列表 ==========
Dim rsProducts, productSql
productSql = "SELECT * FROM Products " & productWhereClause & " ORDER BY ProductID DESC"
Set rsProducts = ExecuteQuery(productSql)

' ========== 获取产品类型列表 ==========
Dim rsTypeConfig
Set rsTypeConfig = ExecuteQuery("SELECT * FROM ProductTypeConfig ORDER BY DisplayOrder ASC")

' ========== 预加载香调数据（用于产品表单） ==========
Dim rsFragranceNotes, noteListTop, noteListMiddle, noteListBase
Set rsFragranceNotes = ExecuteQuery("SELECT * FROM FragranceNotes WHERE IsActive <> 0 ORDER BY NoteType, NoteName")

' ========== 预加载容量数据 ==========
Dim rsVolumes
Set rsVolumes = ExecuteQuery("SELECT * FROM Volumes WHERE IsActive <> 0 ORDER BY VolumeML")

' ========== 预加载瓶型数据 ==========
Dim rsBottleStyles
Set rsBottleStyles = ExecuteQuery("SELECT * FROM BottleStyles WHERE IsActive <> 0 ORDER BY BottleName")

' ========== 预加载配方数据（用于产品表单配方导入） ==========
Dim rsFormulas, formulaDataJson
formulaDataJson = ""
Set rsFormulas = ExecuteQuery("SELECT RecipeID AS FormulaID, RecipeName AS FormulaName, RecipeCode, ProductType FROM Recipes WHERE IsActive <> 0 AND ProductType = 'KOL' AND (ReviewStatus = 'Approved' OR ReviewStatus IS NULL) ORDER BY RecipeName")
If Not rsFormulas Is Nothing Then
    formulaDataJson = "var formulaData = {};"
    Dim currentFormulaId, formulaNotesArr
    Do While Not rsFormulas.EOF
        currentFormulaId = rsFormulas("FormulaID")
        formulaDataJson = formulaDataJson & "formulaData[" & currentFormulaId & "] = ["
        
        ' 查询该配方的香调配比
        Dim rsFormulaNotes
        Set rsFormulaNotes = ExecuteQuery("SELECT NoteID, Percentage FROM RecipeNotes WHERE RecipeID = " & CLng(currentFormulaId))
        If Not rsFormulaNotes Is Nothing Then
            Dim noteIdx
            noteIdx = 0
            Do While Not rsFormulaNotes.EOF
                If noteIdx > 0 Then formulaDataJson = formulaDataJson & ","
                formulaDataJson = formulaDataJson & "{noteId:" & rsFormulaNotes("NoteID") & ",percentage:" & rsFormulaNotes("Percentage") & "}"
                noteIdx = noteIdx + 1
                rsFormulaNotes.MoveNext
            Loop
            rsFormulaNotes.Close
            Set rsFormulaNotes = Nothing
        End If
        
        formulaDataJson = formulaDataJson & "];"
        rsFormulas.MoveNext
    Loop
End If

' ========== 预加载配方数据（用于产品表单关联配方选择） ==========
Dim rsRecipes, recipeDataJson
recipeDataJson = ""
Set rsRecipes = ExecuteQuery("SELECT RecipeID, RecipeName, RecipeCode, ProductType FROM Recipes WHERE IsActive <> 0 AND (ReviewStatus = 'Approved' OR ReviewStatus IS NULL) ORDER BY RecipeCode")
If Not rsRecipes Is Nothing Then
    recipeDataJson = "var recipeData = ["
    Dim recipeIdx
    recipeIdx = 0
    Do While Not rsRecipes.EOF
        If recipeIdx > 0 Then recipeDataJson = recipeDataJson & ","
        Dim rName, rCode, rType
        rName = rsRecipes("RecipeName") & ""
        rCode = rsRecipes("RecipeCode") & ""
        rType = rsRecipes("ProductType") & ""
        rName = Replace(Replace(rName, "\", "\\"), "'", "\'")
        rCode = Replace(Replace(rCode, "\", "\\"), "'", "\'")
        rType = Replace(rType, "'", "\'")
        recipeDataJson = recipeDataJson & "{id:" & rsRecipes("RecipeID") & ",name:'" & rName & "',code:'" & rCode & "',type:'" & rType & "'}"
        recipeIdx = recipeIdx + 1
        rsRecipes.MoveNext
    Loop
    recipeDataJson = recipeDataJson & "];"
End If

' ========== 获取香调配比最小比例设置 ==========
' 显式声明页面级变量，确保在整个页面中可用
Dim minTopPercent, minMiddlePercent, minBasePercent

' 设置默认值
minTopPercent = 10
minMiddlePercent = 10
minBasePercent = 10

' 从数据库读取配置值（使用单一查询和字符串比较）
Dim rsRatioSettings
Set rsRatioSettings = ExecuteQuery("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey IN ('MinTopPercent', 'MinMiddlePercent', 'MinBasePercent')")
If Not rsRatioSettings Is Nothing Then
    Do While Not rsRatioSettings.EOF
        Dim settingKey, settingValue
        settingKey = UCase(Trim(rsRatioSettings("SettingKey") & ""))
        settingValue = Trim(rsRatioSettings("SettingValue") & "")
        
        If settingKey = "MINTOPPERCENT" And IsNumeric(settingValue) Then
            minTopPercent = CInt(settingValue)
        ElseIf settingKey = "MINMIDDLEPERCENT" And IsNumeric(settingValue) Then
            minMiddlePercent = CInt(settingValue)
        ElseIf settingKey = "MINBASEPERCENT" And IsNumeric(settingValue) Then
            minBasePercent = CInt(settingValue)
        End If
        rsRatioSettings.MoveNext
    Loop
    rsRatioSettings.Close
End If
Set rsRatioSettings = Nothing

' ========== 获取类型统计数据 ==========
Dim productStats
Set productStats = CreateObject("Scripting.Dictionary")

Dim rsStats
Set rsStats = ExecuteQuery("SELECT ProductType, COUNT(*) AS Total, SUM(IIF(IsActive<>0, 1, 0)) AS ActiveCount FROM Products GROUP BY ProductType")
If Not rsStats Is Nothing Then
    Do While Not rsStats.EOF
        Dim statKey, statArray
        statKey = CStr(rsStats("ProductType").Value)
        ReDim statArray(1)
        statArray(0) = CLng("0" & rsStats("Total").Value)
        statArray(1) = CLng("0" & rsStats("ActiveCount").Value)
        productStats.Add statKey, statArray
        rsStats.MoveNext
    Loop
    rsStats.Close
End If
Set rsStats = Nothing

' ========== 计算总体统计 ==========
Dim totalProductCount, activeProductCount, inactiveProductCount
Dim fixedProductCount, customProductCount, kolProductCount
totalProductCount = 0
activeProductCount = 0
inactiveProductCount = 0
fixedProductCount = 0
customProductCount = 0
kolProductCount = 0

If productStats.Exists("Fixed") Then
    fixedProductCount = productStats("Fixed")(0)
End If
If productStats.Exists("Custom") Then
    customProductCount = productStats("Custom")(0)
End If
If productStats.Exists("KOL") Then
    kolProductCount = productStats("KOL")(0)
End If

Dim statKeyAll
For Each statKeyAll In productStats.Keys
    totalProductCount = totalProductCount + productStats(statKeyAll)(0)
    activeProductCount = activeProductCount + productStats(statKeyAll)(1)
Next
inactiveProductCount = totalProductCount - activeProductCount
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>产品设置 - 产品技术管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 页面头部 */
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .page-title {
            font-size: 24px;
            color: #fff;
            margin: 0;
        }
        .page-title i {
            color: #00bcd4;
            margin-right: 10px;
        }
        .breadcrumb {
            font-size: 14px;
            color: #888;
        }
        .breadcrumb a {
            color: #00bcd4;
            text-decoration: none;
        }
        .breadcrumb a:hover {
            text-decoration: underline;
        }
        
        /* Tab导航 */
        .tab-nav {
            display: flex;
            gap: 5px;
            margin-bottom: 25px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            padding-bottom: 0;
        }
        .tab-link {
            padding: 12px 24px;
            background: transparent;
            color: #888;
            text-decoration: none;
            border-radius: 6px 6px 0 0;
            transition: all 0.2s ease;
            border-bottom: 2px solid transparent;
            margin-bottom: -1px;
        }
        .tab-link:hover {
            color: #fff;
            background: rgba(255,255,255,0.05);
        }
        .tab-link.active {
            color: #00bcd4;
            border-bottom-color: #00bcd4;
            background: rgba(0,188,212,0.05);
        }
        
        /* 筛选栏 */
        .filter-bar {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 25px;
            border: 1px solid rgba(255,255,255,0.05);
            display: flex;
            gap: 20px;
            align-items: center;
            flex-wrap: wrap;
        }
        .filter-group {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .filter-label {
            font-size: 14px;
            color: #888;
        }
        .search-box {
            display: flex;
            gap: 10px;
        }
        .search-input, .filter-select {
            padding: 10px 15px;
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 6px;
            background: rgba(255,255,255,0.05);
            color: #fff;
            font-size: 14px;
        }
        .search-input {
            width: 200px;
        }
        .search-input:focus, .filter-select:focus {
            outline: none;
            border-color: #00bcd4;
        }
        .search-input::placeholder {
            color: #999;
        }
        select.filter-select option {
            background: #2d2d44;
            color: #fff;
        }
        
        /* 卡片样式 */
        .admin-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            border: 1px solid rgba(255,255,255,0.05);
            margin-bottom: 25px;
        }
        .admin-card-header {
            padding: 20px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .admin-card-title {
            font-size: 18px;
            color: #fff;
            margin: 0;
        }
        .admin-card-body {
            padding: 20px;
        }
        
        /* .admin-btn 样式已由 /css/buttons.css Section 2 & 4 & 5 统一管理 */
        
        /* 表格样式 */
        .admin-table {
            width: 100%;
            border-collapse: collapse;
        }
        .admin-table th {
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 500;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .admin-table td {
            padding: 15px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            color: #e0e0e0;
        }
        .admin-table tr:hover td {
            background: rgba(255,255,255,0.02);
        }
        .admin-table tr:last-child td {
            border-bottom: none;
        }
        
        /* 状态标签 */
        .status-badge {
            display: inline-block;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 500;
        }
        .status-active {
            background: rgba(76,175,80,0.2);
            color: #4caf50;
        }
        .status-inactive {
            background: rgba(244,67,54,0.2);
            color: #f44336;
        }
        .status-pending {
            background: rgba(255,193,7,0.2);
            color: #ffc107;
        }
        .status-approved {
            background: rgba(33,150,243,0.2);
            color: #2196f3;
        }
        .status-fixed {
            background: rgba(156,39,176,0.2);
            color: #9c27b0;
        }
        .status-custom {
            background: rgba(0,150,136,0.2);
            color: #009688;
        }
        .status-kol {
            background: rgba(255,87,34,0.2);
            color: #ff5722;
        }
        
        /* 操作按钮组 */
        .action-btns {
            display: flex;
            gap: 8px;
        }
        
        /* 提示消息 */
        .alert {
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .alert-success {
            background: rgba(76,175,80,0.1);
            color: #4caf50;
            border-left: 4px solid #4caf50;
        }
        
        /* 模态框样式 */
        .admin-modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.7);
            backdrop-filter: blur(4px);
        }
        .admin-modal-content {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            margin: 5% auto;
            border-radius: 12px;
            width: 90%;
            max-width: 700px;
            border: 1px solid rgba(255,255,255,0.1);
            box-shadow: 0 20px 60px rgba(0,0,0,0.5);
            max-height: 90vh;
            overflow-y: auto;
        }
        .admin-modal-header {
            padding: 20px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .admin-modal-title {
            font-size: 18px;
            color: #fff;
            margin: 0;
        }
        .admin-modal-close {
            background: none;
            border: none;
            color: #bbb;
            font-size: 24px;
            cursor: pointer;
            transition: color 0.2s;
        }
        .admin-modal-close:hover {
            color: #fff;
        }
        .admin-modal-body {
            padding: 20px;
        }
        .admin-modal-footer {
            padding: 20px;
            border-top: 1px solid rgba(255,255,255,0.1);
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        
        /* 表单样式 */
        .admin-form-group {
            margin-bottom: 20px;
        }
        .admin-form-label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #e0e0e0;
            font-size: 14px;
        }
        .admin-modal-content .admin-form-label {
            color: #e0e0e0 !important;
        }
        .admin-form-control {
            width: 100%;
            padding: 12px 15px;
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 8px;
            font-size: 14px;
            box-sizing: border-box;
            background: rgba(255,255,255,0.05);
            color: #fff;
        }
        .admin-form-control:focus {
            border-color: #00bcd4;
            outline: none;
        }
        .admin-form-row {
            display: flex;
            gap: 20px;
        }
        .admin-form-col {
            flex: 1;
        }
        select.admin-form-control option {
            background: #2d2d44;
            color: #fff;
        }
        .admin-form-control::placeholder { color: #999; }
        textarea.admin-form-control {
            resize: vertical;
            min-height: 80px;
        }
        .checkbox-group {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .checkbox-group input[type="checkbox"] {
            width: 18px;
            height: 18px;
            cursor: pointer;
        }
        
        /* 类型卡片 */
        .type-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
        }
        .type-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
        }
        .type-card:hover {
            border-color: rgba(0,188,212,0.3);
            transform: translateY(-2px);
        }
        .type-card-header {
            display: flex;
            align-items: center;
            gap: 15px;
            margin-bottom: 15px;
        }
        .type-icon {
            width: 50px;
            height: 50px;
            border-radius: 10px;
            background: linear-gradient(135deg, #00bcd4 0%, #00838f 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            color: white;
        }
        .type-info h4 {
            margin: 0 0 5px 0;
            color: #fff;
            font-size: 16px;
        }
        .type-code {
            font-size: 12px;
            color: #888;
            font-family: monospace;
        }
        .type-stats {
            display: flex;
            gap: 15px;
            margin: 15px 0;
            padding: 15px 0;
            border-top: 1px solid rgba(255,255,255,0.05);
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }
        .type-stat {
            text-align: center;
        }
        .type-stat-value {
            font-size: 20px;
            font-weight: 600;
            color: #00bcd4;
        }
        .type-stat-label {
            font-size: 12px;
            color: #888;
        }
        .type-features {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            margin-bottom: 15px;
        }
        .type-feature {
            font-size: 12px;
            padding: 4px 10px;
            border-radius: 12px;
            background: rgba(255,255,255,0.05);
            color: #888;
        }
        .type-feature.active {
            background: rgba(0,188,212,0.2);
            color: #00bcd4;
        }
        
        /* 空状态 */
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #999;
        }
        .empty-state i {
            font-size: 64px;
            margin-bottom: 20px;
            color: #888;
        }
        
        /* 权限提示 */
        .readonly-notice {
            background: rgba(255,193,7,0.1);
            padding: 12px 15px;
            border-radius: 6px;
            color: #ffc107;
            font-size: 13px;
            margin-bottom: 20px;
            border-left: 4px solid #ffc107;
        }
        
        /* 复选框网格 */
        .checkbox-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
            gap: 10px;
        }
        .checkbox-item {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 8px 12px;
            background: rgba(255,255,255,0.05);
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 13px;
        }
        .checkbox-item:hover {
            background: rgba(255,255,255,0.1);
        }
        .checkbox-item input[type="checkbox"] {
            width: 16px;
            height: 16px;
            cursor: pointer;
        }
        .checkbox-item span {
            flex: 1;
            color: #e0e0e0;
        }
        .checkbox-item.selected {
            background: rgba(0,188,212,0.2);
            border: 1px solid rgba(0,188,212,0.5);
        }
        
        /* KOL比例输入 */
        .note-percent-input {
            background: rgba(255,255,255,0.1);
            border: 1px solid rgba(255,255,255,0.2);
            border-radius: 4px;
            color: #fff;
            padding: 2px 5px;
            font-size: 12px;
        }
        .note-percent-input:focus {
            border-color: #00bcd4;
            outline: none;
        }
        
        /* 产品卡片网格 */
        .product-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        }
        .product-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
        }
        .product-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            border-color: rgba(0,188,212,0.2);
        }
        
        /* 产品卡片图片 */
        .product-card-image {
            position: relative;
            width: 100%;
            padding-top: 100%;
            border-radius: 8px;
            overflow: hidden;
            background: rgba(0,0,0,0.25);
            margin-bottom: 15px;
        }
        .product-card-image img {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            object-fit: cover;
            transition: transform 0.3s ease;
        }
        .product-card:hover .product-card-image img {
            transform: scale(1.05);
        }
        .product-card-image .img-placeholder {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-direction: column;
            gap: 8px;
            color: #555;
            background: rgba(0,0,0,0.25);
        }
        .product-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 12px;
        }
        .product-title {
            font-size: 16px;
            font-weight: 600;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .product-title i { color: #00bcd4; }
        .product-id {
            font-size: 11px;
            color: #999;
            background: rgba(0,0,0,0.3);
            padding: 2px 8px;
            border-radius: 4px;
        }
        .product-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 12px;
            align-items: center;
        }
        .product-price {
            font-size: 18px;
            font-weight: 700;
            color: #00bcd4;
            margin-bottom: 12px;
        }
        .product-info-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
            font-size: 13px;
        }
        .product-info-label { color: #888; }
        .product-info-value { color: #e0e0e0; }
        .product-footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 15px;
            border-top: 1px solid rgba(255,255,255,0.05);
            margin-top: 10px;
        }
        .product-type-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 10px;
            font-weight: 600;
            text-transform: uppercase;
        }
        .product-type-badge.fixed { background: rgba(33, 150, 243, 0.2); color: #2196f3; }
        .product-type-badge.custom { background: rgba(76, 175, 80, 0.2); color: #4caf50; }
        .product-type-badge.kol { background: rgba(156,39,176,0.2); color: #9c27b0; }

        /* 统计卡片 */
        .stats-section {
            margin-bottom: 25px;
        }
        .stats-cards {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 15px;
        }
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            cursor: pointer;
        }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,188,212,0.2); }
        .stat-card.active { border-color: #00bcd4; box-shadow: 0 0 15px rgba(0,188,212,0.3); }
        .stat-value { font-size: 28px; font-weight: 700; color: #fff; }
        .stat-label { font-size: 11px; color: #888; margin-top: 5px; text-transform: uppercase; }
        .stat-card.total .stat-value { color: #00bcd4; }
        .stat-card.active-stat .stat-value { color: #4caf50; }
        .stat-card.inactive-stat .stat-value { color: #f44336; }
        .stat-card.fixed-stat .stat-value { color: #2196f3; }
        .stat-card.custom-stat .stat-value { color: #4caf50; }
        .stat-card.kol-stat .stat-value { color: #ff9800; }

        /* 响应式 */
        @media (max-width: 1200px) {
            .stats-cards { grid-template-columns: repeat(3, 1fr); }
            .product-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width: 768px) {
            .tab-nav {
                flex-wrap: wrap;
            }
            .filter-bar {
                flex-direction: column;
                align-items: stretch;
            }
            .search-box {
                width: 100%;
            }
            .search-input, .filter-select {
                flex: 1;
            }
            .admin-form-row {
                flex-direction: column;
                gap: 0;
            }
            .type-grid {
                grid-template-columns: 1fr;
            }
            .product-grid { grid-template-columns: 1fr; }
            .stats-cards { grid-template-columns: repeat(2, 1fr); }
        }

        /* ====== 弹窗颜色覆盖（防 admin.css 污染）====== */
        .admin-modal-content .admin-modal-title,
        .admin-modal-content h1,
        .admin-modal-content h2,
        .admin-modal-content h3,
        .admin-modal-content h4,
        .admin-modal-content h5,
        .admin-modal-content h6 { color: #ffffff !important; }
        .admin-modal-content .admin-modal-close { color: #bbb !important; }
        .admin-modal-content .admin-modal-close:hover { color: #fff !important; }
        .admin-modal-content .form-label,
        .admin-modal-content .admin-form-label { color: #e0e0e0 !important; }
        .admin-modal-content .form-control,
        .admin-modal-content .admin-form-control { color: #fff !important; background: rgba(255,255,255,0.05) !important; border-color: rgba(255,255,255,0.1) !important; }
        .admin-modal-content .form-control::placeholder,
        .admin-modal-content .admin-form-control::placeholder { color: #999 !important; }
        .admin-modal-content small,
        .admin-modal-content .form-text,
        .admin-modal-content .text-muted { color: #aaa !important; }
        .admin-modal-footer { background: rgba(0,0,0,0.2) !important; }
        /* 图片上传组件 */
        .image-upload-wrapper {
            border: 2px dashed rgba(255,255,255,0.15);
            border-radius: 12px;
            padding: 16px;
            text-align: center;
            transition: border-color 0.3s;
        }
        .image-upload-wrapper:hover,
        .image-upload-wrapper.dragover {
            border-color: rgba(0,188,212,0.5);
        }
        .image-preview {
            width: 100%;
            max-width: 280px;
            height: 200px;
            margin: 0 auto 12px;
            border-radius: 8px;
            overflow: hidden;
            background: rgba(0,0,0,0.2);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
        }
        .image-preview img {
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
        }
        .image-placeholder {
            color: #666;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
        }
        .image-placeholder i {
            font-size: 36px;
            color: #555;
        }
        .image-placeholder span {
            font-size: 13px;
        }
        .image-upload-actions {
            display: flex;
            gap: 8px;
            justify-content: center;
            margin-top: 8px;
        }
        .image-upload-actions .btn-sm {
            padding: 6px 14px;
            font-size: 13px;
            border-radius: 6px;
            border: none;
            cursor: pointer;
        }
        .upload-progress {
            margin-top: 10px;
            background: rgba(0,0,0,0.3);
            border-radius: 6px;
            height: 24px;
            position: relative;
            overflow: hidden;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #00bcd4, #00e5ff);
            border-radius: 6px;
            width: 0%;
            transition: width 0.3s;
        }
        .progress-text {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 12px;
            color: #fff;
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2 class="page-title"><i class="fas fa-box-open"></i> 产品设置</h2>
                <div class="breadcrumb">
                    <a href="index.asp">技术中心</a> / <span>产品设置</span>
                </div>
            </div>
        </div>
        
        <% If Request.QueryString("msg") <> "" Then %>
        <div class="alert alert-success">
            <i class="fas fa-check-circle"></i>
            <%= Server.HTMLEncode(Request.QueryString("msg")) %>
        </div>
        <% End If %>
        
        <% If Not isManager Then %>
        <div class="readonly-notice">
            <i class="fas fa-info-circle"></i> 您当前为技术人员，部分管理功能（删除、创建类型）需要技术经理权限
        </div>
        <% End If %>
        
        <!-- Tab导航 -->
        <div class="tab-nav">
            <a href="?tab=products" class="tab-link <%= IIf(currentTab = "products", "active", "") %>">
                <i class="fas fa-box"></i> 产品管理
            </a>
            <a href="?tab=types" class="tab-link <%= IIf(currentTab = "types", "active", "") %>">
                <i class="fas fa-tags"></i> 类型配置
            </a>
            <a href="?tab=ratio" class="tab-link <%= IIf(currentTab = "ratio", "active", "") %>">
                <i class="fas fa-percentage"></i> 香调配比参数
            </a>
        </div>
        
        <% If currentTab = "products" Then %>
        <!-- 产品管理Tab -->
        
        <!-- 统计区域 -->
        <div class="stats-section">
            <div class="stats-cards">
                <div class="stat-card total <%= IIf(filterProductType="" And productSearch="", "active", "") %>" onclick="location.href='?tab=products'">
                    <div class="stat-value"><%= totalProductCount %></div>
                    <div class="stat-label">总产品</div>
                </div>
                <div class="stat-card active-stat" onclick="location.href='?tab=products'">
                    <div class="stat-value"><%= activeProductCount %></div>
                    <div class="stat-label">上架中</div>
                </div>
                <div class="stat-card inactive-stat" onclick="location.href='?tab=products'">
                    <div class="stat-value"><%= inactiveProductCount %></div>
                    <div class="stat-label">已下架</div>
                </div>
                <div class="stat-card fixed-stat <%= IIf(filterProductType="Fixed", "active", "") %>" onclick="location.href='?tab=products&product_type=Fixed'">
                    <div class="stat-value"><%= fixedProductCount %></div>
                    <div class="stat-label">品牌定香</div>
                </div>
                <div class="stat-card custom-stat <%= IIf(filterProductType="Custom", "active", "") %>" onclick="location.href='?tab=products&product_type=Custom'">
                    <div class="stat-value"><%= customProductCount %></div>
                    <div class="stat-label">用户定制</div>
                </div>
                <div class="stat-card kol-stat <%= IIf(filterProductType="KOL", "active", "") %>" onclick="location.href='?tab=products&product_type=KOL'">
                    <div class="stat-value"><%= kolProductCount %></div>
                    <div class="stat-label">KOL推荐</div>
                </div>
            </div>
        </div>
        
        <div class="filter-bar">
            <div class="filter-group">
                <span class="filter-label"><i class="fas fa-filter"></i> 类型筛选：</span>
                <select class="filter-select" onchange="location.href='?tab=products&product_type='+this.value+'<%= IIf(productSearch <> "", "&product_search=" & Server.URLEncode(productSearch), "") %>'">
                    <option value="">全部类型</option>
                    <% 
                    If IsArray(allProductTypes) Then
                        Dim ptIdx, ptCode, ptName
                        For ptIdx = 0 To UBound(allProductTypes, 1)
                            ptCode = allProductTypes(ptIdx, 0)
                            ptName = allProductTypes(ptIdx, 1)
                    %>
                    <option value="<%= ptCode %>" <%= IIf(filterProductType = ptCode, "selected", "") %>><%= HTMLEncode(ptName) %></option>
                    <% 
                        Next
                    End If
                    %>
                </select>
            </div>
            <div class="filter-group" style="margin-left: auto;">
                <form method="get" class="search-box">
                    <input type="hidden" name="tab" value="products">
                    <% If filterProductType <> "" Then %>
                    <input type="hidden" name="product_type" value="<%= filterProductType %>">
                    <% End If %>
                    <input type="text" name="product_search" class="search-input" placeholder="搜索产品名称..." value="<%= HTMLEncode(productSearch) %>">
                    <button type="submit" class="admin-btn admin-btn-primary admin-btn-sm">
                        <i class="fas fa-search"></i>
                    </button>
                    <% If productSearch <> "" Then %>
                    <a href="?tab=products<%= IIf(filterProductType <> "", "&product_type=" & filterProductType, "") %>" class="admin-btn admin-btn-outline admin-btn-sm">
                        <i class="fas fa-times"></i> 清除
                    </a>
                    <% End If %>
                </form>
            </div>
            <button class="admin-btn admin-btn-primary" onclick="showAddProductForm()">
                <i class="fas fa-plus"></i> 新增产品
            </button>
            <a href="../purchase/fixed_brand/product_management.asp" class="admin-btn admin-btn-outline" style="margin-left: 10px;">
                <i class="fas fa-box"></i> 品牌定香采购管理
            </a>
        </div>
        
        <!-- 产品卡片列表 -->
        <div class="product-grid">
            <% 
            Dim hasProducts
            hasProducts = False
            If Not rsProducts Is Nothing Then 
                If Not rsProducts.EOF Then
                    hasProducts = True
                End If
            End If
            
            If hasProducts Then 
                Do While Not rsProducts.EOF 
                    Dim pType, pTypeDisplay, pTypeClass, pBadgeClass
                    pType = rsProducts("ProductType")
                    pTypeDisplay = ""
                    pTypeClass = ""
                    pBadgeClass = ""
                    
                    ' 获取类型显示名称
                    If IsArray(allProductTypes) Then
                        For ptIdx = 0 To UBound(allProductTypes, 1)
                            If allProductTypes(ptIdx, 0) = pType Then
                                pTypeDisplay = allProductTypes(ptIdx, 1)
                                Exit For
                            End If
                        Next
                    End If
                    If pTypeDisplay = "" Then pTypeDisplay = pType
                    
                    ' 设置类型样式
                    Select Case pType
                        Case "Fixed": pTypeClass = "status-fixed": pBadgeClass = "fixed"
                        Case "Custom": pTypeClass = "status-custom": pBadgeClass = "custom"
                        Case "KOL": pTypeClass = "status-kol": pBadgeClass = "kol"
                        Case Else: pTypeClass = "": pBadgeClass = ""
                    End Select
                    
                    ' 计算实际显示价格
                    Dim pDisplayPrice
                    pDisplayPrice = SafeNum(rsProducts("BasePrice"))
                    If pType = "Fixed" Then
                        Dim rsFixedDispPrice
                        Set rsFixedDispPrice = ExecuteQuery("SELECT MIN(Price) AS MinPrice FROM ProductVolumePrices WHERE ProductID = " & rsProducts("ProductID"))
                        If Not rsFixedDispPrice Is Nothing Then
                            If Not rsFixedDispPrice.EOF Then
                                If Not IsNull(rsFixedDispPrice("MinPrice")) And rsFixedDispPrice("MinPrice") & "" <> "" Then
                                    pDisplayPrice = CDbl(rsFixedDispPrice("MinPrice"))
                                End If
                            End If
                            rsFixedDispPrice.Close
                        End If
                        Set rsFixedDispPrice = Nothing
                    End If
            %>
            <div class="product-card">
                <div class="product-card-image">
                    <%
                    Dim pImgUrl
                    pImgUrl = Trim(rsProducts("ImageURL") & "")
                    If pImgUrl = "" Then
                        Response.Write "<div class='img-placeholder'><i class='fas fa-box-open'></i><span>暂无图片</span></div>"
                    Else
                    %>
                    <img src="<%= HTMLEncode(pImgUrl) %>" alt="<%= HTMLEncode(rsProducts("ProductName")) %>" onerror="this.onerror=null;this.parentElement.innerHTML='<div class=\\'img-placeholder\\'><i class=\\'fas fa-box-open\\'></i><span>图片加载失败</span></div>'">
                    <% End If %>
                </div>
                <div class="product-header">
                    <div class="product-title">
                        <i class="fas fa-box-open"></i>
                        <%= HTMLEncode(rsProducts("ProductName")) %>
                    </div>
                    <span class="product-id">#<%= rsProducts("ProductID") %></span>
                </div>
                
                <div class="product-meta">
                    <span class="product-type-badge <%= pBadgeClass %>">
                        <% Select Case pType
                            Case "Fixed" %><i class="fas fa-box"></i> 品牌定香
                        <%  Case "Custom" %><i class="fas fa-paint-brush"></i> 用户定制
                        <%  Case "KOL" %><i class="fas fa-star"></i> KOL推荐
                        <%  Case Else %><%= HTMLEncode(pType) %>
                        <% End Select %>
                    </span>
                    
                    <% 
                    Dim rStatus, rClass
                    rStatus = ""
                    On Error Resume Next
                    rStatus = rsProducts("ReviewStatus")
                    If Err.Number <> 0 Then rStatus = "Pending"
                    On Error GoTo 0
                    
                    Select Case rStatus
                        Case "Pending": rClass = "status-pending"
                        Case "Approved": rClass = "status-approved"
                        Case "Rejected": rClass = "status-inactive"
                        Case Else: rClass = "status-pending"
                    End Select
                    
                    Dim needsReview
                    needsReview = False
                    If IsArray(allProductTypes) Then
                        For ptIdx = 0 To UBound(allProductTypes, 1)
                            If allProductTypes(ptIdx, 0) = pType Then
                                needsReview = allProductTypes(ptIdx, 5)
                                Exit For
                            End If
                        Next
                    End If
                    
                    If needsReview Then
                        Select Case rStatus
                            Case "Pending": Response.Write "<span class='status-badge " & rClass & "'>待审核</span>"
                            Case "Approved": Response.Write "<span class='status-badge " & rClass & "'>已通过</span>"
                            Case "Rejected": Response.Write "<span class='status-badge " & rClass & "'>已驳回</span>"
                            Case Else: Response.Write "<span class='status-badge " & rClass & "'>待审核</span>"
                        End Select
                    End If
                    %>
                    
                    <% 
                    Dim pIsActive
                    pIsActive = 1
                    On Error Resume Next
                    pIsActive = rsProducts("IsActive")
                    If Err.Number <> 0 Then pIsActive = 1
                    On Error GoTo 0
                    
                    If pIsActive <> 0 Then
                        Response.Write "<span class='status-badge status-active'>上架</span>"
                    Else
                        Response.Write "<span class='status-badge status-inactive'>下架</span>"
                    End If
                    %>
                </div>
                
                <div class="product-price">¥<%= FormatNumber(pDisplayPrice, 2) %></div>
                
                <div class="product-info-row">
                    <span class="product-info-label"><i class="fas fa-leaf"></i> 基香成分</span>
                    <span class="product-info-value">
                        <% 
                        Dim baseIng
                        baseIng = ""
                        On Error Resume Next
                        baseIng = Trim(rsProducts("BaseIngredients") & "")
                        On Error GoTo 0
                        
                        If baseIng <> "" Then
                            Response.Write "<span style='color:#00bcd4;'><i class='fas fa-check'></i> 有</span>"
                        Else
                            Response.Write "<span style='color:#999;'>无</span>"
                        End If
                        %>
                    </span>
                </div>
                
                <% 
                ' 获取产品关联数据
                Dim productNotesData, productVolumesData, productRatiosData, productBottlesData
                productNotesData = ""
                productVolumesData = ""
                productRatiosData = ""
                productBottlesData = ""
                
                ' 获取香调和配比
                Dim rsProdNotes
                Set rsProdNotes = ExecuteQuery("SELECT pn.NoteID, IIF(pnr.Percentage IS NULL, 0, pnr.Percentage) AS Percentage FROM ProductNotes pn LEFT JOIN ProductNoteRatios pnr ON pn.ProductID = pnr.ProductID AND pn.NoteID = pnr.NoteID WHERE pn.ProductID = " & rsProducts("ProductID"))
                If Not rsProdNotes Is Nothing Then
                    Do While Not rsProdNotes.EOF
                        If productNotesData <> "" Then productNotesData = productNotesData & ","
                        productNotesData = productNotesData & rsProdNotes("NoteID")
                        ' 配比数据格式: NoteID:Percentage
                        If productRatiosData <> "" Then productRatiosData = productRatiosData & ","
                        productRatiosData = productRatiosData & rsProdNotes("NoteID") & ":" & rsProdNotes("Percentage")
                        rsProdNotes.MoveNext
                    Loop
                    rsProdNotes.Close
                End If
                Set rsProdNotes = Nothing
                
                ' 获取容量
                Dim rsProdVols
                Set rsProdVols = ExecuteQuery("SELECT VolumeID FROM ProductVolumePrices WHERE ProductID = " & rsProducts("ProductID"))
                If Not rsProdVols Is Nothing Then
                    Do While Not rsProdVols.EOF
                        If productVolumesData <> "" Then productVolumesData = productVolumesData & ","
                        productVolumesData = productVolumesData & rsProdVols("VolumeID")
                        rsProdVols.MoveNext
                    Loop
                    rsProdVols.Close
                End If
                Set rsProdVols = Nothing
                
                ' 获取瓶型配置
                Dim rsProdBottles
                Set rsProdBottles = ExecuteQuery("SELECT BottleID, CustomPrice FROM ProductBottleStyles WHERE ProductID = " & rsProducts("ProductID"))
                If Not rsProdBottles Is Nothing Then
                    Do While Not rsProdBottles.EOF
                        If productBottlesData <> "" Then productBottlesData = productBottlesData & ","
                        productBottlesData = productBottlesData & "{'bid':" & rsProdBottles("BottleID") & ",'price':" & SafeNum(rsProdBottles("CustomPrice")) & "}"
                        rsProdBottles.MoveNext
                    Loop
                    rsProdBottles.Close
                End If
                Set rsProdBottles = Nothing
                %>
                <%
                ' 获取刻字配置
                Dim pEngravable, pEngravingPrice
                On Error Resume Next
                pEngravable = rsProducts("Engravable")
                If Err.Number <> 0 Then pEngravable = 0
                pEngravingPrice = rsProducts("EngravingPrice")
                If Err.Number <> 0 Then pEngravingPrice = 0
                On Error GoTo 0
                %>
                
                <div class="product-footer">
                    <div class="action-btns">
                        <% If pType <> "Fixed" Then %>
                        <button class="admin-btn admin-btn-sm admin-btn-outline" onclick="showEditProductForm(this)" 
                            data-id="<%= rsProducts("ProductID") %>" 
                            data-name="<%= SafeOutput(rsProducts("ProductName")) %>" 
                            data-desc="<%= SafeOutput(rsProducts("Description") & "") %>"
                            data-price="<%= pDisplayPrice %>"
                            data-type="<%= pType %>"
                            data-baseing="<%= SafeOutput(baseIng) %>"
                            data-review="<%= rStatus %>"
                            data-active="<%= pIsActive %>"
                            data-image="<%= SafeOutput(rsProducts("ImageURL") & "") %>"
                            data-kolid="<%= SafeNum(rsProducts("KOLID")) %>"
                            data-engravable="<%= pEngravable %>"
                            data-engravingprice="<%= SafeNum(pEngravingPrice) %>"
                            data-recipeid="<%= SafeNum(rsProducts("RecipeID")) %>"
                            data-notes="<%= productNotesData %>"
                            data-ratios="<%= productRatiosData %>"
                            data-volumes="<%= productVolumesData %>"
                            data-bottles="[<%= productBottlesData %>]">
                            <i class="fas fa-edit"></i> 编辑
                        </button>
                        <% If isManager Then %>
                            <% If pIsActive <> 0 Then %>
                            <form method="post" style="display:inline;" onsubmit="return confirm('确定要下架此产品吗？')">
                                <input type="hidden" name="action" value="delete_product">
                                <input type="hidden" name="productId" value="<%= rsProducts("ProductID") %>">
                                <button type="submit" class="admin-btn admin-btn-sm admin-btn-danger">
                                    <i class="fas fa-ban"></i> 下架
                                </button>
                            </form>
                            <% Else %>
                            <form method="post" style="display:inline;" onsubmit="return confirm('确定要恢复此产品吗？')">
                                <input type="hidden" name="action" value="restore_product">
                                <input type="hidden" name="productId" value="<%= rsProducts("ProductID") %>">
                                <button type="submit" class="admin-btn admin-btn-sm admin-btn-success">
                                    <i class="fas fa-undo"></i> 恢复
                                </button>
                            </form>
                            <% End If %>
                        <% End If %>
                        <% Else %>
                        <a href="../purchase/fixed_brand/product_management.asp" class="admin-btn admin-btn-sm admin-btn-outline" title="品牌定香产品请前往采购模块管理">
                            <i class="fas fa-truck"></i> 去采购管理
                        </a>
                        <span style="font-size:11px;color:#999;margin-left:5px;" title="品牌定香产品请前往采购模块管理">
                            <i class="fas fa-info-circle"></i> 采购模块管理
                        </span>
                        <% End If %>
                    </div>
                </div>
            </div>
            <% rsProducts.MoveNext %>
            <% Loop %>
            <% Else %>
            <div class="empty-state" style="grid-column: 1 / -1;">
                <i class="fas fa-box"></i>
                <h3>暂无产品数据</h3>
                <p>点击"新增产品"按钮创建第一个产品</p>
            </div>
            <% End If %>
        </div>
        
        <% ElseIf currentTab = "types" Then %>
        <!-- 类型配置Tab -->
        <div class="admin-card">
            <div class="admin-card-header">
                <h3 class="admin-card-title"><i class="fas fa-tags"></i> 产品类型配置</h3>
            </div>
            <div class="admin-card-body">
                <div class="type-grid">
                    <% 
                    If Not rsTypeConfig Is Nothing Then
                        Do While Not rsTypeConfig.EOF
                            Dim tcId, tcCode, tcDisplay, tcNav, tcDesc, tcIcon, tcReview, tcRatio, tcOrder, tcActive
                            tcId = rsTypeConfig("ConfigID")
                            tcCode = rsTypeConfig("TypeCode")
                            tcDisplay = rsTypeConfig("DisplayName")
                            tcNav = rsTypeConfig("NavName") & ""
                            tcDesc = rsTypeConfig("Description") & ""
                            tcIcon = rsTypeConfig("Icon") & ""
                            tcReview = rsTypeConfig("RequiresReview")
                            tcRatio = rsTypeConfig("RequiresRatio")
                            tcOrder = rsTypeConfig("DisplayOrder")
                            tcActive = rsTypeConfig("IsActive")
                            
                            ' 获取该类型产品数量
                            Dim typeTotal, typeActive
                            typeTotal = 0
                            typeActive = 0
                            If productStats.Exists(tcCode) Then
                                typeTotal = productStats(tcCode)(0)
                                typeActive = productStats(tcCode)(1)
                            End If
                            
                            ' 设置类型样式
                            Dim tcClass
                            Select Case tcCode
                                Case "Fixed": tcClass = "status-fixed"
                                Case "Custom": tcClass = "status-custom"
                                Case "KOL": tcClass = "status-kol"
                                Case Else: tcClass = ""
                            End Select
                    %>
                    <div class="type-card">
                        <div class="type-card-header">
                            <div class="type-icon">
                                <% If tcIcon <> "" Then %>
                                <i class="<%= tcIcon %>"></i>
                                <% Else %>
                                <i class="fas fa-box"></i>
                                <% End If %>
                            </div>
                            <div class="type-info">
                                <h4><%= HTMLEncode(tcDisplay) %></h4>
                                <span class="type-code"><%= tcCode %></span>
                            </div>
                        </div>
                        
                        <div class="type-stats">
                            <div class="type-stat">
                                <div class="type-stat-value"><%= typeActive %></div>
                                <div class="type-stat-label">上架产品</div>
                            </div>
                            <div class="type-stat">
                                <div class="type-stat-value"><%= typeTotal %></div>
                                <div class="type-stat-label">总产品</div>
                            </div>
                            <div class="type-stat">
                                <div class="type-stat-value"><%= tcOrder %></div>
                                <div class="type-stat-label">排序</div>
                            </div>
                        </div>
                        
                        <div class="type-features">
                            <span class="type-feature <%= IIf(tcReview, "active", "") %>">
                                <i class="fas <%= IIf(tcReview, "fa-check", "fa-times") %>"></i> 需要审核
                            </span>
                            <span class="type-feature <%= IIf(tcRatio, "active", "") %>">
                                <i class="fas <%= IIf(tcRatio, "fa-check", "fa-times") %>"></i> 需要配比
                            </span>
                            <span class="type-feature <%= IIf(tcActive, "active", "") %>">
                                <i class="fas <%= IIf(tcActive, "fa-check", "fa-times") %>"></i> 已启用
                            </span>
                        </div>
                        
                        <% If tcDesc <> "" Then %>
                        <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                            <%= HTMLEncode(Left(tcDesc, 50)) %><%= IIf(Len(tcDesc) > 50, "...", "") %>
                        </p>
                        <% End If %>
                        
                        <div class="action-btns">
                            <button class="admin-btn admin-btn-sm admin-btn-outline" onclick="showEditTypeForm(this)"
                                data-id="<%= tcId %>"
                                data-code="<%= SafeOutput(tcCode) %>"
                                data-display="<%= SafeOutput(tcDisplay) %>"
                                data-nav="<%= SafeOutput(tcNav) %>"
                                data-desc="<%= SafeOutput(tcDesc) %>"
                                data-icon="<%= SafeOutput(tcIcon) %>"
                                data-review="<%= tcReview %>"
                                data-ratio="<%= tcRatio %>"
                                data-order="<%= tcOrder %>"
                                data-active="<%= tcActive %>">
                                <i class="fas fa-edit"></i> 编辑
                            </button>
                        </div>
                    </div>
                    <% 
                            rsTypeConfig.MoveNext
                        Loop
                        rsTypeConfig.Close
                        Set rsTypeConfig = Nothing
                    End If
                    %>
                </div>
            </div>
        </div>
        
        <% ElseIf currentTab = "ratio" Then %>
        <!-- 香调配比参数Tab -->
        <div class="admin-card">
            <div class="admin-card-header">
                <h3 class="admin-card-title"><i class="fas fa-percentage"></i> 香调配比参数设置</h3>
                <p style="color: #888; font-size: 14px; margin-top: 10px;">
                    设置定制香水和KOL推荐商品的前、中、后调最小比例限制，确保配方平衡
                </p>
            </div>
            <div class="admin-card-body">
                <% If Request.QueryString("msg") <> "" Then %>
                <div class="alert alert-success" style="margin-bottom: 20px;">
                    <i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %>
                </div>
                <% End If %>
                
                <form method="post" action="product_settings.asp?tab=ratio">
                    <input type="hidden" name="action" value="save_ratio_settings">
                    
                    <div class="type-grid">
                        <!-- 前调最小比例 -->
                        <div class="type-card">
                            <div class="type-card-header">
                                <div class="type-icon" style="background: #e8f5e9; color: #2e7d32;">
                                    <i class="fas fa-wind"></i>
                                </div>
                                <div class="type-info">
                                    <h4>前调最小比例</h4>
                                    <span class="type-code">Top Note</span>
                                </div>
                            </div>
                            <div style="padding: 15px;">
                                <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                                    前调是香水的第一印象，设置最小比例确保香水有足够的首香特征。
                                </p>
                                <div class="admin-form-group">
                                    <label class="admin-form-label">最小比例 (%)</label>
                                    <input type="number" name="minTopPercent" value="<%= minTopPercent %>" 
                                        min="0" max="100" step="1" class="admin-form-control" required>
                                </div>
                                <div style="margin-top: 10px; padding: 8px; background: #e8f5e9; border-radius: 4px; font-size: 13px; color: #2e7d32;">
                                    <i class="fas fa-info-circle"></i> 建议值：10% - 30%
                                </div>
                            </div>
                        </div>
                        
                        <!-- 中调最小比例 -->
                        <div class="type-card">
                            <div class="type-card-header">
                                <div class="type-icon" style="background: #fff3e0; color: #e65100;">
                                    <i class="fas fa-heart"></i>
                                </div>
                                <div class="type-info">
                                    <h4>中调最小比例</h4>
                                    <span class="type-code">Middle Note</span>
                                </div>
                            </div>
                            <div style="padding: 15px;">
                                <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                                    中调是香水的核心灵魂，设置最小比例确保香水有持久的主香特征。
                                </p>
                                <div class="admin-form-group">
                                    <label class="admin-form-label">最小比例 (%)</label>
                                    <input type="number" name="minMiddlePercent" value="<%= minMiddlePercent %>" 
                                        min="0" max="100" step="1" class="admin-form-control" required>
                                </div>
                                <div style="margin-top: 10px; padding: 8px; background: #fff3e0; border-radius: 4px; font-size: 13px; color: #e65100;">
                                    <i class="fas fa-info-circle"></i> 建议值：10% - 40%
                                </div>
                            </div>
                        </div>
                        
                        <!-- 后调最小比例 -->
                        <div class="type-card">
                            <div class="type-card-header">
                                <div class="type-icon" style="background: #f3e5f5; color: #7b1fa2;">
                                    <i class="fas fa-moon"></i>
                                </div>
                                <div class="type-info">
                                    <h4>后调最小比例</h4>
                                    <span class="type-code">Base Note</span>
                                </div>
                            </div>
                            <div style="padding: 15px;">
                                <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                                    后调是香水的持久余韵，设置最小比例确保香水有足够的留香时间。
                                </p>
                                <div class="admin-form-group">
                                    <label class="admin-form-label">最小比例 (%)</label>
                                    <input type="number" name="minBasePercent" value="<%= minBasePercent %>" 
                                        min="0" max="100" step="1" class="admin-form-control" required>
                                </div>
                                <div style="margin-top: 10px; padding: 8px; background: #f3e5f5; border-radius: 4px; font-size: 13px; color: #7b1fa2;">
                                    <i class="fas fa-info-circle"></i> 建议值：10% - 30%
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div style="margin-top: 30px; padding: 20px; background: #e3f2fd; border: 1px solid #2196F3; border-radius: 8px;">
                        <h4 style="margin: 0 0 15px 0; color: #1565c0;">
                            <i class="fas fa-lightbulb"></i> 设置说明
                        </h4>
                        <ul style="margin: 0; padding-left: 20px; color: #1565c0; line-height: 1.8;">
                            <li>该设置适用于<strong>定制香水</strong>和<strong>KOL推荐</strong>两种商品类型</li>
                            <li>用户在前台购买时，系统会验证前、中、后调的比例是否都达到最小值</li>
                            <li>后台管理员新增KOL商品时，也会验证该配比规则</li>
                            <li>建议三种调性的最小比例之和不超过60%，以保留调配灵活性</li>
                            <li>当前设置：前调 <strong><%= minTopPercent %>%</strong> | 中调 <strong><%= minMiddlePercent %>%</strong> | 后调 <strong><%= minBasePercent %>%</strong></li>
                        </ul>
                    </div>
                    
                    <div style="margin-top: 30px; text-align: center;">
                        <button type="submit" class="admin-btn admin-btn-primary" style="padding: 12px 40px; font-size: 16px;" <%= IIf(isManager, "", "disabled") %>>
                            <i class="fas fa-save"></i> 保存设置
                        </button>
                    </div>
                </form>
            </div>
        </div>
        <% End If %>
    </div>
    
    <!-- 添加/编辑产品模态框 -->
    <div id="productModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 id="productModalTitle" class="admin-modal-title">新增产品</h3>
                <button class="admin-modal-close" onclick="closeProductModal()">&times;</button>
            </div>
            <form id="productForm" method="post">
                <div class="admin-modal-body">
                    <input type="hidden" id="productFormAction" name="action" value="add_product">
                    <input type="hidden" id="editProductId" name="productId" value="">
                    <input type="hidden" name="csrf_token" value="<%= Session("CSRFToken") %>">
                    
                    <div class="admin-form-group">
                        <label for="productName" class="admin-form-label">产品名称 *</label>
                        <input type="text" id="productName" name="productName" class="admin-form-control" required placeholder="请输入产品名称">
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="productDescription" class="admin-form-label">产品描述</label>
                        <textarea id="productDescription" name="description" class="admin-form-control" rows="3" placeholder="请输入产品描述"></textarea>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="productType" class="admin-form-label">产品类型 *</label>
                                <select id="productType" name="productType" class="admin-form-control" required onchange="toggleProductFields()">
                                    <% 
                                    If IsArray(allProductTypes) Then
                                        For ptIdx = 0 To UBound(allProductTypes, 1)
                                    %>
                                    <option value="<%= allProductTypes(ptIdx, 0) %>" 
                                            data-review="<%= allProductTypes(ptIdx, 5) %>"
                                            data-ratio="<%= allProductTypes(ptIdx, 6) %>">
                                        <%= HTMLEncode(allProductTypes(ptIdx, 1)) %>
                                    </option>
                                    <% 
                                        Next
                                    End If
                                    %>
                                </select>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="basePrice" class="admin-form-label">基础价格 (¥) *</label>
                                <input type="number" id="basePrice" name="basePrice" step="0.01" min="0" class="admin-form-control" required value="0" placeholder="0.00">
                            </div>
                        </div>
                    </div>
                    
                    <!-- 关联配方选择 -->
                    <div class="admin-form-group" id="recipeFields" style="display:none;">
                        <label for="recipeId" class="admin-form-label" id="recipeLabel">关联配方</label>
                        <select id="recipeId" name="recipeId" class="admin-form-control">
                            <option value="">-- 请选择 --</option>
                            <% 
                            If Not rsRecipes Is Nothing Then
                                rsRecipes.MoveFirst
                                Do While Not rsRecipes.EOF
                            %>
                            <option value="<%= rsRecipes("RecipeID") %>" data-rtype="<%= SafeOutput(rsRecipes("ProductType") & "") %>">[<%= SafeOutput(rsRecipes("RecipeCode") & "") %>] <%= HTMLEncode(rsRecipes("RecipeName")) %></option>
                            <% 
                                    rsRecipes.MoveNext
                                Loop
                            End If
                            %>
                        </select>
                        <small id="recipeHint" style="color:#bbb;"></small>
                    </div>
                    
                    <!-- Fixed类型特有：基香成分 -->
                    <div id="fixedFields" class="admin-form-group" style="display:none;">
                        <label for="baseIngredients" class="admin-form-label">基香成分</label>
                        <textarea id="baseIngredients" name="baseIngredients" class="admin-form-control" rows="2" placeholder="多个成分用逗号分隔"></textarea>
                        <small style="color:#bbb;">品牌定香产品特有的基香成分列表</small>
                    </div>
                    
                    <!-- KOL类型特有：KOL选择 -->
                    <div id="kolFields" class="admin-form-group" style="display:none;">
                        <label for="kolId" class="admin-form-label">推荐KOL ID</label>
                        <input type="number" id="kolId" name="kolId" class="admin-form-control" value="0" min="0" placeholder="输入KOL的ID">
                        <small style="color:#bbb;">输入推荐此产品的KOL ID（0表示无特定KOL）</small>
                    </div>
                    
                    <!-- 需要审核的类型特有：审核状态 -->
                    <!-- KOL类型强制待审核，不显示选择界面 -->
                    <input type="hidden" id="reviewStatus" name="reviewStatus" value="Pending">
                    <div id="reviewFields" class="admin-form-group" style="display:none;">
                        <label for="reviewStatusSelect" class="admin-form-label">审核状态</label>
                        <select id="reviewStatusSelect" class="admin-form-control">
                            <option value="Pending">待审核</option>
                            <option value="Approved">已通过</option>
                            <option value="Rejected">已驳回</option>
                        </select>
                        <small style="color:#bbb;">该产品类型需要运营审核</small>
                    </div>
                    
                    <!-- Custom和KOL类型特有：香调配置 -->
                    <div id="fragranceFields" class="admin-form-group" style="display:none;">
                        <label class="admin-form-label">可选香调配置</label>
                        
                        <!-- 配方选择下拉框（仅KOL类型可见） -->
                        <div id="formulaImportFields" style="margin-bottom:15px;">
                            <label for="formulaSelect" class="admin-form-label" style="font-size:13px;font-weight:normal;color:#aaa;">从配方导入（可选）</label>
                            <select id="formulaSelect" class="admin-form-control" onchange="applyFormula(this.value)">
                                <option value="">-- 手动选择香调 --</option>
                                <% 
                                If Not rsFormulas Is Nothing Then
                                    rsFormulas.MoveFirst
                                    Do While Not rsFormulas.EOF
                                %>
                                <option value="<%= rsFormulas("FormulaID") %>"><%= HTMLEncode(rsFormulas("FormulaName")) %></option>
                                <% 
                                        rsFormulas.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                            <small style="color:#bbb;">选择一个配方可自动填充香调配比，您仍可手动调整</small>
                        </div>
                        
                        <div style="background:rgba(255,255,255,0.03);padding:15px;border-radius:8px;border:1px solid rgba(255,255,255,0.1);">
                            <!-- 前调 -->
                            <div style="margin-bottom:15px;">
                                <div style="color:#00bcd4;font-size:13px;font-weight:500;margin-bottom:8px;"><i class="fas fa-wind"></i> 前调</div>
                                <div class="checkbox-grid" id="topNotesContainer">
                                    <% 
                                    If Not rsFragranceNotes Is Nothing Then
                                        rsFragranceNotes.MoveFirst
                                        Do While Not rsFragranceNotes.EOF
                                            If rsFragranceNotes("NoteType") = "前调" Then
                                    %>
                                    <label class="checkbox-item">
                                        <input type="checkbox" name="noteCheckbox" value="<%= rsFragranceNotes("NoteID") %>" data-type="top" data-name="<%= HTMLEncode(rsFragranceNotes("NoteName")) %>" onchange="toggleNotePercentInput(this)">
                                        <span><%= HTMLEncode(rsFragranceNotes("NoteName")) %></span>
                                        <input type="number" name="notePercent_<%= rsFragranceNotes("NoteID") %>" class="note-percent-input" data-note-type="top" placeholder="%" min="0" max="100" style="width:50px;margin-left:5px;display:none;" oninput="updateRatioSummary()">
                                    </label>
                                    <% 
                                            End If
                                            rsFragranceNotes.MoveNext
                                        Loop
                                    End If
                                    %>
                                </div>
                            </div>
                            <!-- 中调 -->
                            <div style="margin-bottom:15px;">
                                <div style="color:#e91e63;font-size:13px;font-weight:500;margin-bottom:8px;"><i class="fas fa-heart"></i> 中调</div>
                                <div class="checkbox-grid" id="middleNotesContainer">
                                    <% 
                                    If Not rsFragranceNotes Is Nothing Then
                                        rsFragranceNotes.MoveFirst
                                        Do While Not rsFragranceNotes.EOF
                                            If rsFragranceNotes("NoteType") = "中调" Then
                                    %>
                                    <label class="checkbox-item">
                                        <input type="checkbox" name="noteCheckbox" value="<%= rsFragranceNotes("NoteID") %>" data-type="middle" data-name="<%= HTMLEncode(rsFragranceNotes("NoteName")) %>" onchange="toggleNotePercentInput(this)">
                                        <span><%= HTMLEncode(rsFragranceNotes("NoteName")) %></span>
                                        <input type="number" name="notePercent_<%= rsFragranceNotes("NoteID") %>" class="note-percent-input" data-note-type="middle" placeholder="%" min="0" max="100" style="width:50px;margin-left:5px;display:none;" oninput="updateRatioSummary()">
                                    </label>
                                    <% 
                                            End If
                                            rsFragranceNotes.MoveNext
                                        Loop
                                    End If
                                    %>
                                </div>
                            </div>
                            <!-- 后调 -->
                            <div>
                                <div style="color:#9c27b0;font-size:13px;font-weight:500;margin-bottom:8px;"><i class="fas fa-moon"></i> 后调</div>
                                <div class="checkbox-grid" id="baseNotesContainer">
                                    <% 
                                    If Not rsFragranceNotes Is Nothing Then
                                        rsFragranceNotes.MoveFirst
                                        Do While Not rsFragranceNotes.EOF
                                            If rsFragranceNotes("NoteType") = "后调" Then
                                    %>
                                    <label class="checkbox-item">
                                        <input type="checkbox" name="noteCheckbox" value="<%= rsFragranceNotes("NoteID") %>" data-type="base" data-name="<%= HTMLEncode(rsFragranceNotes("NoteName")) %>" onchange="toggleNotePercentInput(this)">
                                        <span><%= HTMLEncode(rsFragranceNotes("NoteName")) %></span>
                                        <input type="number" name="notePercent_<%= rsFragranceNotes("NoteID") %>" class="note-percent-input" data-note-type="base" placeholder="%" min="0" max="100" style="width:50px;margin-left:5px;display:none;" oninput="updateRatioSummary()">
                                    </label>
                                    <% 
                                            End If
                                            rsFragranceNotes.MoveNext
                                        Loop
                                    End If
                                    %>
                                </div>
                            </div>
                        </div>
                        <input type="hidden" id="selectedNotes" name="selectedNotes" value="">
                        <small style="color:#bbb;">选择该产品可用的香调（Custom和KOL类型）</small>
                        
                        <!-- KOL配比提示 -->
                        <div id="ratioSummary" style="margin-top:15px;padding:12px 15px;background:rgba(0,188,212,0.1);border-radius:8px;border:1px solid rgba(0,188,212,0.3);display:none;">
                            <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
                                <div>
                                    <span style="color:#00bcd4;font-weight:500;">配比统计：</span>
                                    <span id="ratioDetail" style="color:#e0e0e0;">前调: 0% | 中调: 0% | 后调: 0%</span>
                                </div>
                                <div id="ratioTotal" style="font-weight:600;color:#ff9800;">总计: 0%</div>
                            </div>
                            <div id="ratioError" style="margin-top:8px;color:#f44336;font-size:13px;display:none;"></div>
                        </div>
                    </div>
                    
                    <!-- 容量配置 -->
                    <div id="volumeFields" class="admin-form-group">
                        <label class="admin-form-label">可选容量配置</label>
                        <div class="checkbox-grid" id="volumesContainer">
                            <% 
                            If Not rsVolumes Is Nothing Then
                                Do While Not rsVolumes.EOF
                            %>
                            <label class="checkbox-item volume-item">
                                <input type="checkbox" name="volumeCheckbox" value="<%= rsVolumes("VolumeID") %>" data-ml="<%= rsVolumes("VolumeML") %>" data-multiplier="<%= rsVolumes("PriceMultiplier") %>">
                                <span><%= rsVolumes("VolumeML") %>ml - <%= HTMLEncode(rsVolumes("VolumeName")) %></span>
                                <span style="color:#b0b0b0;font-size:12px;">(×<%= rsVolumes("PriceMultiplier") %>)</span>
                            </label>
                            <% 
                                    rsVolumes.MoveNext
                                Loop
                            End If
                            %>
                        </div>
                        <input type="hidden" id="selectedVolumes" name="selectedVolumes" value="">
                        <small style="color:#bbb;">选择该产品可用的容量规格</small>
                    </div>
                    
                    <!-- Custom和KOL类型特有：瓶型配置 -->
                    <div id="bottleFields" class="admin-form-group" style="display:none;">
                        <label class="admin-form-label">可选瓶型配置</label>
                        <div class="checkbox-grid" id="bottlesContainer">
                            <% 
                            Dim defaultBottlePrice
                            If Not rsBottleStyles Is Nothing Then
                                Do While Not rsBottleStyles.EOF
                                    defaultBottlePrice = SafeNum(rsBottleStyles("PriceAddition"))
                            %>
                            <label class="checkbox-item bottle-item">
                                <input type="checkbox" name="bottleCheckbox" value="<%= rsBottleStyles("BottleID") %>" data-default-price="<%= defaultBottlePrice %>">
                                <span><%= HTMLEncode(rsBottleStyles("BottleName")) %></span>
                                <span style="color:#b0b0b0;font-size:12px;margin-left:auto;">(+<%= FormatNumber(defaultBottlePrice, 0) %>元)</span>
                            </label>
                            <% 
                                    rsBottleStyles.MoveNext
                                Loop
                            End If
                            %>
                        </div>
                        <input type="hidden" id="selectedBottles" name="selectedBottles" value="">
                        <small style="color:#bbb;">选择该产品可用的瓶型款式（价格统一在瓶型管理页面设置）</small>
                    </div>
                    
                    <!-- Custom和KOL类型特有：刻字配置 -->
                    <div id="engravingFields" class="admin-form-group" style="display:none;">
                        <label class="admin-form-label">刻字配置</label>
                        <div style="background:rgba(255,255,255,0.03);padding:15px;border-radius:8px;border:1px solid rgba(255,255,255,0.1);">
                            <div class="checkbox-group" style="margin-bottom:10px;">
                                <input type="checkbox" id="engravable" name="engravable" value="1">
                                <label for="engravable">支持瓶身刻字</label>
                            </div>
                            <div id="engravingPriceWrapper" style="display:none;">
                                <label style="font-size:13px;color:#b0b0b0;">刻字附加费用：</label>
                                <input type="number" id="engravingPrice" name="engravingPrice" class="admin-form-control" style="width:150px;display:inline-block;" step="0.01" min="0" value="0" placeholder="0.00">
                                <span style="color:#b0b0b0;">元</span>
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label class="admin-form-label">图片</label>
                                <div class="image-upload-wrapper">
                                    <div class="image-preview" id="imagePreview_product">
                                        <img id="previewImg_product" src="" alt="预览" style="display:none;">
                                        <div class="image-placeholder" id="placeholder_product">
                                            <i class="fas fa-cloud-upload-alt"></i>
                                            <span>点击上传或拖拽图片</span>
                                        </div>
                                    </div>
                                    <input type="file" id="fileInput_product" accept="image/jpeg,image/png,image/gif,image/webp,image/svg+xml" style="display:none;">
                                    <div class="image-upload-actions">
                                        <button type="button" class="admin-btn admin-btn-info btn-sm" onclick="document.getElementById('fileInput_product').click();">
                                            <i class="fas fa-upload"></i> 选择图片
                                        </button>
                                        <button type="button" class="admin-btn admin-btn-secondary btn-sm" onclick="toggleUrlInput_product()">
                                            <i class="fas fa-link"></i> 输入URL
                                        </button>
                                    </div>
                                    <div id="urlInputWrapper_product" style="display:none; margin-top:8px;">
                                        <input type="text" id="manualUrl_product" class="admin-form-control" placeholder="输入图片URL地址" style="font-size:13px;">
                                        <button type="button" class="admin-btn admin-btn-secondary btn-sm" onclick="applyManualUrl_product()" style="margin-top:4px;">确认</button>
                                    </div>
                                    <div class="upload-progress" id="uploadProgress_product" style="display:none;">
                                        <div class="progress-bar" id="progressBar_product"></div>
                                        <span class="progress-text" id="progressText_product">上传中...</span>
                                    </div>
                                    <div style="font-size:11px;color:#888;margin-top:6px;">如果原图超过 180KB，将自动压缩后再上传</div>
                                    <input type="hidden" name="imageURL" id="imageURL_product" value="/images/default-product.svg">
                                </div>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="productIsActive" class="admin-form-label">状态</label>
                                <select id="productIsActive" name="isActive" class="admin-form-control">
                                    <option value="1">上架</option>
                                    <option value="0">下架</option>
                                </select>
                            </div>
                        </div>
                    </div>
                </div>
                <div id="fixedTypeWarning" style="display:none; margin:0 24px 16px 24px; padding:12px 16px; background:#fff3e0; border:1px solid #ff9800; border-radius:6px; color:#e65100; font-size:14px;">
                    <i class="fas fa-exclamation-triangle" style="margin-right:6px;"></i>
                    <strong>品牌定香产品</strong>请前往<strong>采购管理 → 品牌定香</strong>模块进行创建和编辑管理，此处仅支持查看。
                    <a href="../purchase/fixed_brand/product_management.asp" style="color:#e65100;font-weight:bold;text-decoration:underline;margin-left:8px;">立即前往 →</a>
                </div>
                <div class="admin-modal-footer">
                    <button type="button" class="admin-btn admin-btn-outline" onclick="closeProductModal()">取消</button>
                    <button type="submit" id="submitProductBtn" class="admin-btn admin-btn-primary">
                        <i class="fas fa-save"></i> 保存
                    </button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 编辑类型模态框 -->
    <div id="typeModal" class="admin-modal">
        <div class="admin-modal-content">
            <div class="admin-modal-header">
                <h3 class="admin-modal-title">编辑产品类型</h3>
                <button class="admin-modal-close" onclick="closeTypeModal()">&times;</button>
            </div>
            <form id="typeForm" method="post">
                <div class="admin-modal-body">
                    <input type="hidden" name="action" value="edit_type">
                    <input type="hidden" id="typeConfigId" name="configId" value="">
                    
                    <div class="admin-form-group">
                        <label class="admin-form-label">类型代码</label>
                        <input type="text" id="typeCodeDisplay" class="admin-form-control" readonly style="background:rgba(255,255,255,0.02);">
                        <small style="color:#bbb;">类型代码不可修改</small>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="typeDisplayName" class="admin-form-label">显示名称 *</label>
                                <input type="text" id="typeDisplayName" name="displayName" class="admin-form-control" required>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="typeNavName" class="admin-form-label">栏目名称</label>
                                <input type="text" id="typeNavName" name="navName" class="admin-form-control" placeholder="为空则不在导航显示">
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <label for="typeDescription" class="admin-form-label">描述</label>
                        <textarea id="typeDescription" name="description" class="admin-form-control" rows="2"></textarea>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="typeIcon" class="admin-form-label">图标</label>
                                <input type="text" id="typeIcon" name="icon" class="admin-form-control" placeholder="如：fas fa-box">
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <label for="typeDisplayOrder" class="admin-form-label">排序号</label>
                                <input type="number" id="typeDisplayOrder" name="displayOrder" class="admin-form-control" value="0" min="0">
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-row">
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <div class="checkbox-group">
                                    <input type="checkbox" id="typeRequiresReview" name="requiresReview" value="1">
                                    <label for="typeRequiresReview">需要审核</label>
                                </div>
                                <small style="color:#bbb;">该类型产品需要运营审核后才能上架</small>
                            </div>
                        </div>
                        <div class="admin-form-col">
                            <div class="admin-form-group">
                                <div class="checkbox-group">
                                    <input type="checkbox" id="typeRequiresRatio" name="requiresRatio" value="1">
                                    <label for="typeRequiresRatio">需要配比</label>
                                </div>
                                <small style="color:#bbb;">该类型产品需要设置香调配比</small>
                            </div>
                        </div>
                    </div>
                    
                    <div class="admin-form-group">
                        <div class="checkbox-group">
                            <input type="checkbox" id="typeIsActive" name="isActive" value="1">
                            <label for="typeIsActive">启用该类型</label>
                        </div>
                        <small style="color:#bbb;">禁用后该类型不会在前台显示</small>
                    </div>
                </div>
                <div class="admin-modal-footer">
                    <button type="button" class="admin-btn admin-btn-outline" onclick="closeTypeModal()">取消</button>
                    <button type="submit" class="admin-btn admin-btn-primary" <%= IIf(isManager, "", "disabled") %>>
                        <i class="fas fa-save"></i> 保存
                    </button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
        // 香调最小比例配置（从后端获取）
        var minTopPercent = <%=minTopPercent%>;
        var minMiddlePercent = <%=minMiddlePercent%>;
        var minBasePercent = <%=minBasePercent%>;
        
        // 配方数据（从后端预加载）
        <%=formulaDataJson%>
        
        // 关联配方数据（从后端预加载）
        <%=recipeDataJson%>
        
        // 应用配方到表单
        function applyFormula(formulaId) {
            if (!formulaId || !formulaData[formulaId]) {
                return;
            }
            
            // 先清空所有选择
            var allCheckboxes = document.querySelectorAll('input[name="noteCheckbox"]');
            for (var i = 0; i < allCheckboxes.length; i++) {
                allCheckboxes[i].checked = false;
                var percentInput = document.querySelector('input[name="notePercent_' + allCheckboxes[i].value + '"]');
                if (percentInput) {
                    percentInput.value = '';
                    percentInput.style.display = 'none';
                }
                allCheckboxes[i].closest('.checkbox-item').classList.remove('selected');
            }
            
            // 获取配方数据
            var notes = formulaData[formulaId];
            var selectedNoteIds = [];
            
            // 应用配方中的香调和百分比
            for (var i = 0; i < notes.length; i++) {
                var noteId = notes[i].noteId;
                var percentage = notes[i].percentage;
                
                var checkbox = document.querySelector('input[name="noteCheckbox"][value="' + noteId + '"]');
                if (checkbox) {
                    checkbox.checked = true;
                    checkbox.closest('.checkbox-item').classList.add('selected');
                    selectedNoteIds.push(noteId);
                    
                    // 设置百分比
                    var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
                    if (percentInput) {
                        percentInput.value = percentage;
                        // 如果是KOL类型，显示百分比输入框
                        var productType = document.getElementById('productType').value;
                        if (productType === 'KOL') {
                            percentInput.style.display = 'inline-block';
                        }
                    }
                }
            }
            
            // 更新选中的香调隐藏字段
            document.getElementById('selectedNotes').value = selectedNoteIds.join(',');
            
            // 更新配比统计
            var productType = document.getElementById('productType').value;
            if (productType === 'KOL') {
                updateRatioSummary();
            }
        }
        
        // closest() polyfill for older browsers
        if (!Element.prototype.closest) {
            Element.prototype.closest = function(s) {
                var el = this;
                do {
                    if (el.matches(s)) return el;
                    el = el.parentElement || el.parentNode;
                } while (el !== null && el.nodeType === 1);
                return null;
            };
        }
        
        // matches() polyfill for older browsers
        if (!Element.prototype.matches) {
            Element.prototype.matches = Element.prototype.msMatchesSelector || Element.prototype.webkitMatchesSelector;
        }
        
        // 产品表单相关
        function showAddProductForm() {
            document.getElementById('productModalTitle').textContent = '新增产品';
            document.getElementById('productFormAction').value = 'add_product';
            document.getElementById('editProductId').value = '';
            document.getElementById('productName').value = '';
            document.getElementById('productDescription').value = '';
            document.getElementById('productType').value = 'Custom';
            document.getElementById('basePrice').value = '0';
            document.getElementById('baseIngredients').value = '';
            document.getElementById('reviewStatus').value = 'Pending';
            document.getElementById('imageURL_product').value = '/images/default-product.svg';
            document.getElementById('previewImg_product').style.display = 'none';
            document.getElementById('placeholder_product').style.display = 'flex';
            document.getElementById('fileInput_product').value = '';
            document.getElementById('productIsActive').value = '1';
            document.getElementById('kolId').value = '0';
            document.getElementById('engravable').checked = false;
            document.getElementById('engravingPrice').value = '0';
            document.getElementById('engravingPriceWrapper').style.display = 'none';
            
            // 清除所有复选框
            var checkboxes = document.querySelectorAll('input[type="checkbox"]');
            for (var i = 0; i < checkboxes.length; i++) {
                checkboxes[i].checked = false;
                var checkboxItem = checkboxes[i].closest('.checkbox-item');
                if (checkboxItem) {
                    checkboxItem.classList.remove('selected');
                }
            }
            // 隐藏所有瓶型价格输入框
            var bottlePriceWrappers = document.querySelectorAll('.bottle-price-wrapper');
            for (var i = 0; i < bottlePriceWrappers.length; i++) {
                bottlePriceWrappers[i].style.display = 'none';
            }
            document.getElementById('selectedNotes').value = '';
            document.getElementById('selectedVolumes').value = '';
            document.getElementById('selectedBottles').value = '';
            
            // 重置配方选择下拉框
            document.getElementById('formulaSelect').value = '';
            document.getElementById('recipeId').value = '';
            
            toggleProductFields();
            document.getElementById('productModal').style.display = 'block';
        }
        
        function showEditProductForm(button) {
            var id = button.getAttribute('data-id');
            var name = button.getAttribute('data-name');
            var desc = button.getAttribute('data-desc');
            var price = button.getAttribute('data-price');
            var type = button.getAttribute('data-type');
            var baseIng = button.getAttribute('data-baseing');
            var review = button.getAttribute('data-review');
            var active = button.getAttribute('data-active');
            var image = button.getAttribute('data-image');
            var kolId = button.getAttribute('data-kolid');
            var engravable = button.getAttribute('data-engravable');
            var engravingPrice = button.getAttribute('data-engravingprice');
            var recipeId = button.getAttribute('data-recipeid');
            var notesData = button.getAttribute('data-notes');
            var ratiosData = button.getAttribute('data-ratios');
            var volumesData = button.getAttribute('data-volumes');
            var bottlesData = button.getAttribute('data-bottles');
            
            document.getElementById('productModalTitle').textContent = '编辑产品';
            document.getElementById('productFormAction').value = 'edit_product';
            document.getElementById('editProductId').value = id;
            document.getElementById('productName').value = name;
            document.getElementById('productDescription').value = desc;
            document.getElementById('productType').value = type;
            document.getElementById('basePrice').value = price;
            document.getElementById('baseIngredients').value = baseIng;
            // KOL类型强制待审核，其他类型使用传入的审核状态
            if (type === 'KOL') {
                document.getElementById('reviewStatus').value = 'Pending';
            } else {
                document.getElementById('reviewStatus').value = review || 'Pending';
                document.getElementById('reviewStatusSelect').value = review || 'Pending';
            }
            var productImageVal = image || '/images/default-product.svg';
            document.getElementById('imageURL_product').value = productImageVal;
            if (productImageVal && productImageVal !== '/images/default-product.svg') {
                document.getElementById('previewImg_product').src = productImageVal;
                document.getElementById('previewImg_product').style.display = 'block';
                document.getElementById('placeholder_product').style.display = 'none';
            } else {
                document.getElementById('previewImg_product').style.display = 'none';
                document.getElementById('placeholder_product').style.display = 'flex';
            }
            document.getElementById('productIsActive').value = active || '1';
            document.getElementById('kolId').value = kolId || '0';
            document.getElementById('recipeId').value = recipeId || '';
            
            // 清除所有复选框
            var checkboxes = document.querySelectorAll('input[type="checkbox"]');
            for (var i = 0; i < checkboxes.length; i++) {
                checkboxes[i].checked = false;
                if (checkboxes[i].closest('.checkbox-item')) {
                    checkboxes[i].closest('.checkbox-item').classList.remove('selected');
                }
            }
            
            // 加载产品关联数据
            loadProductConfig(notesData, volumesData, type, ratiosData, bottlesData);
            
            // 设置刻字配置
            var engravableCheckbox = document.getElementById('engravable');
            var engravingPriceWrapper = document.getElementById('engravingPriceWrapper');
            var engravingPriceInput = document.getElementById('engravingPrice');
            if (engravableCheckbox) {
                engravableCheckbox.checked = (engravable === '1' || engravable === 'True' || engravable === '-1');
                if (engravingPriceWrapper) {
                    engravingPriceWrapper.style.display = engravableCheckbox.checked ? 'block' : 'none';
                }
                if (engravingPriceInput) {
                    engravingPriceInput.value = engravingPrice || '0';
                }
            }
            
            toggleProductFields();
            document.getElementById('productModal').style.display = 'block';
        }
        
        // 加载产品配置数据
        function loadProductConfig(notesData, volumesData, productType, ratiosData, bottlesData) {
            // 重置配方选择下拉框（编辑时不清除已有选择，只是重置下拉框）
            document.getElementById('formulaSelect').value = '';
            
            // 解析配比数据为字典
            var ratiosDict = {};
            if (ratiosData) {
                var ratioPairs = ratiosData.split(',');
                for (var i = 0; i < ratioPairs.length; i++) {
                    var pair = ratioPairs[i].trim();
                    if (pair && pair.indexOf(':') > -1) {
                        var parts = pair.split(':');
                        ratiosDict[parts[0]] = parts[1];
                    }
                }
            }
            
            // 设置香调
            if (notesData) {
                var noteIds = notesData.split(',');
                var selectedNotes = [];
                for (var i = 0; i < noteIds.length; i++) {
                    var noteId = noteIds[i].trim();
                    if (noteId) {
                        var checkbox = document.querySelector('input[name="noteCheckbox"][value="' + noteId + '"]');
                        if (checkbox) {
                            checkbox.checked = true;
                            checkbox.closest('.checkbox-item').classList.add('selected');
                            selectedNotes.push(noteId);
                            // 加载配比值
                            var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
                            if (percentInput && ratiosDict[noteId]) {
                                percentInput.value = ratiosDict[noteId];
                            }
                            // KOL类型时显示百分比输入框
                            if (percentInput && productType === 'KOL') {
                                percentInput.style.display = 'inline-block';
                            }
                        }
                    }
                }
                document.getElementById('selectedNotes').value = selectedNotes.join(',');
            }
            
            // 设置容量
            if (volumesData) {
                var volIds = volumesData.split(',');
                var selectedVolumes = [];
                for (var i = 0; i < volIds.length; i++) {
                    var volId = volIds[i].trim();
                    if (volId) {
                        var checkbox = document.querySelector('input[name="volumeCheckbox"][value="' + volId + '"]');
                        if (checkbox) {
                            checkbox.checked = true;
                            checkbox.closest('.checkbox-item').classList.add('selected');
                            selectedVolumes.push(volId);
                        }
                    }
                }
                document.getElementById('selectedVolumes').value = selectedVolumes.join(',');
            }
            
            // 设置瓶型
            if (bottlesData) {
                try {
                    // 将单引号替换为双引号以兼容JSON.parse
                    var bottlesJson = bottlesData.replace(/'/g, '"');
                    var bottlesArr = JSON.parse(bottlesJson);
                    var selectedBottles = [];
                    for (var i = 0; i < bottlesArr.length; i++) {
                        var bottle = bottlesArr[i];
                        if (bottle && bottle.bid) {
                            var checkbox = document.querySelector('input[name="bottleCheckbox"][value="' + bottle.bid + '"]');
                            if (checkbox) {
                                checkbox.checked = true;
                                checkbox.closest('.checkbox-item').classList.add('selected');
                                selectedBottles.push(bottle.bid);
                            }
                        }
                    }
                    document.getElementById('selectedBottles').value = selectedBottles.join(',');
                } catch (e) {
                    console.error('解析瓶型数据失败:', e);
                }
            }
        }
        
        // 切换香调百分比输入框显示/隐藏
        function toggleNotePercentInput(checkbox) {
            var noteId = checkbox.value;
            var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
            if (percentInput) {
                percentInput.style.display = checkbox.checked ? 'inline-block' : 'none';
                if (!checkbox.checked) {
                    percentInput.value = ''; // 取消勾选时清空值
                }
            }
            updateSelectedNotes();
        }
        
        function toggleProductFields() {
            var typeSelect = document.getElementById('productType');
            var selectedOption = typeSelect.options[typeSelect.selectedIndex];
            var typeCode = selectedOption.value;
            var requiresReview = selectedOption.getAttribute('data-review') === 'True';
            var requiresRatio = selectedOption.getAttribute('data-ratio') === 'True';
            
            // 品牌定香(Fixed)类型提示横幅
            var fixedWarning = document.getElementById('fixedTypeWarning');
            var submitBtn = document.getElementById('submitProductBtn');
            var isFixed = (typeCode === 'Fixed');
            
            if (fixedWarning) {
                fixedWarning.style.display = isFixed ? 'block' : 'none';
            }
            if (submitBtn) {
                submitBtn.style.display = isFixed ? 'none' : 'inline-flex';
            }
            
            // Fixed类型显示基香成分字段
            document.getElementById('fixedFields').style.display = (typeCode === 'Fixed') ? 'block' : 'none';
            
            // KOL类型显示KOL选择字段
            document.getElementById('kolFields').style.display = (typeCode === 'KOL') ? 'block' : 'none';
            
            // KOL类型强制待审核，不显示审核状态选择
            if (typeCode === 'KOL') {
                document.getElementById('reviewStatus').value = 'Pending';
                document.getElementById('reviewFields').style.display = 'none';
            } else {
                // 其他需要审核的类型显示审核状态字段
                document.getElementById('reviewFields').style.display = requiresReview ? 'block' : 'none';
            }
            
            // Custom和KOL类型显示香调配置
            var isCustomOrKOL = (typeCode === 'Custom' || typeCode === 'KOL');
            document.getElementById('fragranceFields').style.display = isCustomOrKOL ? 'block' : 'none';
            
            // 配方导入区域仅KOL类型可见
            var formulaImportFields = document.getElementById('formulaImportFields');
            if (formulaImportFields) {
                formulaImportFields.style.display = (typeCode === 'KOL') ? 'block' : 'none';
            }
            
            // Custom和KOL类型显示瓶型配置
            document.getElementById('bottleFields').style.display = isCustomOrKOL ? 'block' : 'none';
            
            // 所有类型都显示刻字配置
            document.getElementById('engravingFields').style.display = 'block';
            
            // KOL类型：根据复选框状态显示比例输入
            var percentInputs = document.querySelectorAll('.note-percent-input');
            for (var i = 0; i < percentInputs.length; i++) {
                var noteId = percentInputs[i].name.replace('notePercent_', '');
                var checkbox = document.querySelector('input[name="noteCheckbox"][value="' + noteId + '"]');
                if (typeCode === 'KOL') {
                    percentInputs[i].style.display = (checkbox && checkbox.checked) ? 'inline-block' : 'none';
                } else {
                    percentInputs[i].style.display = 'none';
                }
            }
            
            // 显示/隐藏配比统计区域
            document.getElementById('ratioSummary').style.display = (typeCode === 'KOL') ? 'block' : 'none';
            
            // 如果是KOL类型，更新配比统计
            if (typeCode === 'KOL') {
                updateRatioSummary();
            }
            
            // 显示/隐藏关联配方字段
            var recipeFields = document.getElementById('recipeFields');
            var recipeSelect = document.getElementById('recipeId');
            var recipeLabel = document.getElementById('recipeLabel');
            var recipeHint = document.getElementById('recipeHint');
            if (recipeFields) {
                var showRecipe = (typeCode === 'Fixed');
                recipeFields.style.display = showRecipe ? 'block' : 'none';
                
                // 过滤配方选项
                var options = recipeSelect.querySelectorAll('option[data-rtype]');
                for (var i = 0; i < options.length; i++) {
                    var opt = options[i];
                    if (opt.getAttribute('data-rtype') === typeCode) {
                        opt.style.display = '';
                    } else {
                        opt.style.display = 'none';
                        // 如果当前选中的不是该类型的配方，清空选择
                        if (recipeSelect.value === opt.value) {
                            recipeSelect.value = '';
                        }
                    }
                }
                
                if (typeCode === 'Fixed') {
                    recipeLabel.innerHTML = '关联配方';
                    recipeSelect.required = false;
                    recipeHint.textContent = '可选：品牌定香产品不强制关联配方';
                } else if (typeCode === 'Custom') {
                    recipeLabel.innerHTML = '关联配方';
                    recipeSelect.required = false;
                    recipeHint.textContent = '可选：选择一个推荐配方';
                } else {
                    recipeLabel.innerHTML = '关联配方';
                    recipeSelect.required = false;
                    recipeHint.textContent = '';
                }
            }
        }
        
        // 更新选中的香调
        function updateSelectedNotes() {
            var checkboxes = document.querySelectorAll('input[name="noteCheckbox"]:checked');
            var selected = [];
            for (var i = 0; i < checkboxes.length; i++) {
                selected.push(checkboxes[i].value);
            }
            document.getElementById('selectedNotes').value = selected.join(',');
            
            // 更新配比统计
            var productType = document.getElementById('productType').value;
            if (productType === 'KOL') {
                updateRatioSummary();
            }
        }
        
        // 配比统计和校验
        function updateRatioSummary() {
            var topTotal = 0, middleTotal = 0, baseTotal = 0;
            var checkboxes = document.querySelectorAll('input[name="noteCheckbox"]:checked');
            
            for (var i = 0; i < checkboxes.length; i++) {
                var noteType = checkboxes[i].getAttribute('data-type');
                var noteId = checkboxes[i].value;
                var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
                var percent = percentInput ? parseInt(percentInput.value) || 0 : 0;
                
                if (noteType === 'top') {
                    topTotal += percent;
                } else if (noteType === 'middle') {
                    middleTotal += percent;
                } else if (noteType === 'base') {
                    baseTotal += percent;
                }
            }
            
            var total = topTotal + middleTotal + baseTotal;
            
            // 更新显示 - 包含最小比例提示
            var ratioDetailText = '前调: ' + topTotal + '% (最低' + minTopPercent + '%) | ' +
                                  '中调: ' + middleTotal + '% (最低' + minMiddlePercent + '%) | ' +
                                  '后调: ' + baseTotal + '% (最低' + minBasePercent + '%)';
            document.getElementById('ratioDetail').textContent = ratioDetailText;
            var totalEl = document.getElementById('ratioTotal');
            totalEl.textContent = '总计: ' + total + '%';
            
            // 检查各项是否满足最小比例
            var topValid = topTotal >= minTopPercent;
            var middleValid = middleTotal >= minMiddlePercent;
            var baseValid = baseTotal >= minBasePercent;
            var totalValid = total === 100;
            
            // 根据验证结果设置颜色
            if (totalValid && topValid && middleValid && baseValid) {
                totalEl.style.color = '#4caf50'; // 绿色 - 全部通过
            } else {
                totalEl.style.color = '#ff9800'; // 橙色 - 有错误
            }
            
            // 显示/隐藏错误提示
            var errorEl = document.getElementById('ratioError');
            var errorMsgs = [];
            
            if (!topValid) {
                errorMsgs.push('前调比例(' + topTotal + '%)不得低于最小值' + minTopPercent + '%');
            }
            if (!middleValid) {
                errorMsgs.push('中调比例(' + middleTotal + '%)不得低于最小值' + minMiddlePercent + '%');
            }
            if (!baseValid) {
                errorMsgs.push('后调比例(' + baseTotal + '%)不得低于最小值' + minBasePercent + '%');
            }
            if (!totalValid) {
                errorMsgs.push('配比总和必须等于100%，当前为' + total + '%');
            }
            
            if (errorMsgs.length > 0) {
                errorEl.innerHTML = errorMsgs.join('<br>');
                errorEl.style.display = 'block';
            } else {
                errorEl.style.display = 'none';
            }
            
            // 返回验证结果对象
            return {
                valid: totalValid && topValid && middleValid && baseValid,
                total: total,
                top: topTotal,
                middle: middleTotal,
                base: baseTotal,
                topValid: topValid,
                middleValid: middleValid,
                baseValid: baseValid,
                totalValid: totalValid
            };
        }
        
        // 更新选中的容量
        function updateSelectedVolumes() {
            var checkboxes = document.querySelectorAll('input[name="volumeCheckbox"]:checked');
            var selected = [];
            for (var i = 0; i < checkboxes.length; i++) {
                selected.push(checkboxes[i].value);
            }
            document.getElementById('selectedVolumes').value = selected.join(',');
        }
        
        // 更新选中的瓶型
        function updateSelectedBottles() {
            var checkboxes = document.querySelectorAll('input[name="bottleCheckbox"]:checked');
            var selected = [];
            for (var i = 0; i < checkboxes.length; i++) {
                selected.push(checkboxes[i].value);
            }
            document.getElementById('selectedBottles').value = selected.join(',');
        }
        
        // 绑定复选框事件
        document.addEventListener('DOMContentLoaded', function() {
            // 香调复选框
            var noteCheckboxes = document.querySelectorAll('input[name="noteCheckbox"]');
            for (var i = 0; i < noteCheckboxes.length; i++) {
                noteCheckboxes[i].addEventListener('change', function() {
                    updateSelectedNotes();
                    // 切换选中样式
                    if (this.checked) {
                        this.closest('.checkbox-item').classList.add('selected');
                    } else {
                        this.closest('.checkbox-item').classList.remove('selected');
                    }
                });
            }
            
            // 容量复选框
            var volumeCheckboxes = document.querySelectorAll('input[name="volumeCheckbox"]');
            for (var i = 0; i < volumeCheckboxes.length; i++) {
                volumeCheckboxes[i].addEventListener('change', function() {
                    updateSelectedVolumes();
                    // 切换选中样式
                    if (this.checked) {
                        this.closest('.checkbox-item').classList.add('selected');
                    } else {
                        this.closest('.checkbox-item').classList.remove('selected');
                    }
                });
            }
            
            // 瓶型复选框
            var bottleCheckboxes = document.querySelectorAll('input[name="bottleCheckbox"]');
            for (var i = 0; i < bottleCheckboxes.length; i++) {
                bottleCheckboxes[i].addEventListener('change', function() {
                    // 切换选中样式
                    if (this.checked) {
                        this.closest('.checkbox-item').classList.add('selected');
                    } else {
                        this.closest('.checkbox-item').classList.remove('selected');
                    }
                    // 更新选中列表
                    updateSelectedBottles();
                });
            }
            
            // 配比输入框事件监听
            var percentInputs = document.querySelectorAll('.note-percent-input');
            for (var i = 0; i < percentInputs.length; i++) {
                percentInputs[i].addEventListener('input', function() {
                    updateRatioSummary();
                });
            }
            
            // 表单提交前校验
            var productForm = document.getElementById('productForm');
            if (productForm) {
                productForm.addEventListener('submit', function(e) {
                    var productType = document.getElementById('productType').value;
                    if (productType === 'KOL') {
                        var result = updateRatioSummary();
                        if (!result.valid) {
                            e.preventDefault();
                            var errorMsgs = [];
                            if (!result.topValid) {
                                errorMsgs.push('前调比例(' + result.top + '%)不得低于最小值' + minTopPercent + '%');
                            }
                            if (!result.middleValid) {
                                errorMsgs.push('中调比例(' + result.middle + '%)不得低于最小值' + minMiddlePercent + '%');
                            }
                            if (!result.baseValid) {
                                errorMsgs.push('后调比例(' + result.base + '%)不得低于最小值' + minBasePercent + '%');
                            }
                            if (!result.totalValid) {
                                errorMsgs.push('配比总和必须等于100%，当前为' + result.total + '%');
                            }
                            alert('配比校验失败：\n' + errorMsgs.join('\n'));
                            return false;
                        }
                    }
                });
            }
            
            // 刻字开关
            var engravableCheckbox = document.getElementById('engravable');
            if (engravableCheckbox) {
                engravableCheckbox.addEventListener('change', function() {
                    document.getElementById('engravingPriceWrapper').style.display = this.checked ? 'block' : 'none';
                });
            }
        });
        
        function closeProductModal() {
            document.getElementById('productModal').style.display = 'none';
        }
        
        // 类型表单相关
        function showEditTypeForm(button) {
            var id = button.getAttribute('data-id');
            var code = button.getAttribute('data-code');
            var display = button.getAttribute('data-display');
            var nav = button.getAttribute('data-nav');
            var desc = button.getAttribute('data-desc');
            var icon = button.getAttribute('data-icon');
            var review = button.getAttribute('data-review') === 'True';
            var ratio = button.getAttribute('data-ratio') === 'True';
            var order = button.getAttribute('data-order');
            var active = button.getAttribute('data-active') === 'True';
            
            document.getElementById('typeConfigId').value = id;
            document.getElementById('typeCodeDisplay').value = code;
            document.getElementById('typeDisplayName').value = display;
            document.getElementById('typeNavName').value = nav;
            document.getElementById('typeDescription').value = desc;
            document.getElementById('typeIcon').value = icon;
            document.getElementById('typeDisplayOrder').value = order;
            document.getElementById('typeRequiresReview').checked = review;
            document.getElementById('typeRequiresRatio').checked = ratio;
            document.getElementById('typeIsActive').checked = active;
            
            document.getElementById('typeModal').style.display = 'block';
        }
        
        function closeTypeModal() {
            document.getElementById('typeModal').style.display = 'none';
        }
        
        // 点击模态框外部关闭
        window.onclick = function(event) {
            var productModal = document.getElementById('productModal');
            var typeModal = document.getElementById('typeModal');
            if (event.target == productModal) {
                productModal.style.display = 'none';
            }
            if (event.target == typeModal) {
                typeModal.style.display = 'none';
            }
        }

        // 图片压缩函数 - product
        function compressImage_product(file, maxSizeKB, callback) {
            // SVG 不压缩，直接返回
            if (file.type === 'image/svg+xml') {
                callback(file, false);
                return;
            }
            var maxSize = maxSizeKB * 1024;
            // 文件已经足够小，直接返回
            if (file.size <= maxSize) {
                callback(file, false);
                return;
            }
            var reader = new FileReader();
            reader.onload = function(e) {
                var img = new Image();
                img.onload = function() {
                    var canvas = document.createElement('canvas');
                    var ctx = canvas.getContext('2d');
                    var maxDim = 1200;
                    var width = img.width;
                    var height = img.height;
                    if (width > maxDim || height > maxDim) {
                        if (width > height) {
                            height = Math.round(height * maxDim / width);
                            width = maxDim;
                        } else {
                            width = Math.round(width * maxDim / height);
                            height = maxDim;
                        }
                    }
                    canvas.width = width;
                    canvas.height = height;
                    ctx.drawImage(img, 0, 0, width, height);
                    var quality = 0.8;
                    var tryCompress = function() {
                        canvas.toBlob(function(blob) {
                            if (blob.size > maxSize && quality > 0.1) {
                                quality -= 0.1;
                                tryCompress();
                            } else {
                                var compressedFile = new File([blob], file.name.replace(/\.[^.]+$/, '.jpg'), {
                                    type: 'image/jpeg',
                                    lastModified: Date.now()
                                });
                                callback(compressedFile, true);
                            }
                        }, 'image/jpeg', quality);
                    };
                    tryCompress();
                };
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        }

        // 图片上传 - product
        document.getElementById('fileInput_product').addEventListener('change', function(e) {
            var file = e.target.files[0];
            if (!file) return;
            var maxSize = 5 * 1024 * 1024;
            if (file.size > maxSize) { alert('文件大小不能超过5MB'); return; }
            var allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/svg+xml'];
            if (allowedTypes.indexOf(file.type) === -1) { alert('仅支持 JPG/PNG/GIF/WebP/SVG 格式'); return; }
            var reader = new FileReader();
            reader.onload = function(ev) {
                document.getElementById('previewImg_product').src = ev.target.result;
                document.getElementById('previewImg_product').style.display = 'block';
                document.getElementById('placeholder_product').style.display = 'none';
            };
            reader.readAsDataURL(file);
            compressImage_product(file, 180, function(fileToUpload, wasCompressed) {
                uploadImage_product(fileToUpload, 'product', wasCompressed);
            });
        });

        function uploadImage_product(file, uploadType, wasCompressed) {
            var formData = new FormData();
            formData.append('file', file);
            formData.append('type', uploadType);
            var csrfInput = document.querySelector('input[name="csrf_token"]');
            if (csrfInput) formData.append('csrf_token', csrfInput.value);
            var progressDiv = document.getElementById('uploadProgress_product');
            var progressBar = document.getElementById('progressBar_product');
            var progressText = document.getElementById('progressText_product');
            progressDiv.style.display = 'block';
            progressBar.style.width = '0%';
            progressText.textContent = '上传中...';
            var xhr = new XMLHttpRequest();
            xhr.upload.addEventListener('progress', function(e) {
                if (e.lengthComputable) {
                    var pct = Math.round(e.loaded / e.total * 100);
                    progressBar.style.width = pct + '%';
                    progressText.textContent = pct + '%';
                }
            });
            xhr.addEventListener('load', function() {
                try {
                    var resp = JSON.parse(xhr.responseText);
                    if (resp.success) {
                        document.getElementById('imageURL_product').value = resp.url;
                        progressBar.style.width = '100%';
                        progressText.textContent = wasCompressed ? '上传成功（图片已自动压缩）' : '上传成功';
                        setTimeout(function() { progressDiv.style.display = 'none'; }, 2000);
                    } else {
                        alert('上传失败: ' + (resp.error || '未知错误'));
                        progressDiv.style.display = 'none';
                    }
                } catch(ex) {
                    alert('上传响应解析失败');
                    progressDiv.style.display = 'none';
                }
            });
            xhr.addEventListener('error', function() {
                alert('上传请求失败，请检查网络');
                progressDiv.style.display = 'none';
            });
            xhr.open('POST', '/api/upload.asp', true);
            xhr.send(formData);
        }

        (function() {
            var wrapper = document.getElementById('imagePreview_product').parentElement;
            wrapper.addEventListener('dragover', function(e) { e.preventDefault(); wrapper.classList.add('dragover'); });
            wrapper.addEventListener('dragleave', function() { wrapper.classList.remove('dragover'); });
            wrapper.addEventListener('drop', function(e) {
                e.preventDefault();
                wrapper.classList.remove('dragover');
                var file = e.dataTransfer.files[0];
                if (file) {
                    document.getElementById('fileInput_product').files = e.dataTransfer.files;
                    document.getElementById('fileInput_product').dispatchEvent(new Event('change'));
                }
            });
            document.getElementById('imagePreview_product').addEventListener('click', function() {
                document.getElementById('fileInput_product').click();
            });
        })();

        function toggleUrlInput_product() {
            var el = document.getElementById('urlInputWrapper_product');
            el.style.display = el.style.display === 'none' ? 'block' : 'none';
        }

        function applyManualUrl_product() {
            var url = document.getElementById('manualUrl_product').value.trim();
            if (url) {
                document.getElementById('imageURL_product').value = url;
                document.getElementById('previewImg_product').src = url;
                document.getElementById('previewImg_product').style.display = 'block';
                document.getElementById('placeholder_product').style.display = 'none';
                document.getElementById('urlInputWrapper_product').style.display = 'none';
            }
        }
    </script>
</body>
</html>
<%
If Not rsProducts Is Nothing Then
    rsProducts.Close
    Set rsProducts = Nothing
End If
If Not rsFragranceNotes Is Nothing Then
    rsFragranceNotes.Close
    Set rsFragranceNotes = Nothing
End If
If Not rsVolumes Is Nothing Then
    rsVolumes.Close
    Set rsVolumes = Nothing
End If
If Not rsBottleStyles Is Nothing Then
    rsBottleStyles.Close
    Set rsBottleStyles = Nothing
End If
If Not rsFormulas Is Nothing Then
    rsFormulas.Close
    Set rsFormulas = Nothing
End If
If Not rsRecipes Is Nothing Then
    rsRecipes.Close
    Set rsRecipes = Nothing
End If
Call CloseConnection()
%>
