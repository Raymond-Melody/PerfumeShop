<%
' ============================================
' V14.6 产品设置 - 数据加载器
' 从 product_settings.asp 提取
' 包含：产品筛选、查询构建、列表加载、类型配置、
'       香调/容量/瓶型/配方预加载、统计数据
' ============================================

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
