<%
' ============================================
' 商品类型工具函数库
' 被其他 ASP 页面 include 使用，不应有任何 HTML 输出
' ============================================

' 获取所有已启用的商品类型，返回二维数组
' 数组格式：arr(i, 0)=TypeCode, arr(i, 1)=DisplayName, arr(i, 2)=NavName, 
'           arr(i, 3)=Description, arr(i, 4)=Icon, arr(i, 5)=RequiresReview,
'           arr(i, 6)=RequiresRatio, arr(i, 7)=DisplayOrder, arr(i, 8)=IsActive
' 返回空数组如果无数据
Function GetActiveProductTypes()
    Dim rs, arr, count, i
    Dim sql
    
    sql = "SELECT TypeCode, DisplayName, NavName, Description, Icon, RequiresReview, RequiresRatio, DisplayOrder, IsActive FROM ProductTypeConfig WHERE IsActive<>0 ORDER BY DisplayOrder ASC"
    
    Set rs = ExecuteQuery(sql)
    
    If rs Is Nothing Then
        GetActiveProductTypes = Array()
        Exit Function
    End If
    If rs.EOF Then
        GetActiveProductTypes = Array()
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    ' 获取记录数
    count = 0
    Do While Not rs.EOF
        count = count + 1
        rs.MoveNext
    Loop
    
    ' 返回第一条记录
    rs.MoveFirst
    
    ' 创建数组
    ReDim arr(count - 1, 8)
    
    ' 填充数组
    i = 0
    Do While Not rs.EOF
        arr(i, 0) = rs("TypeCode").Value
        arr(i, 1) = rs("DisplayName").Value
        arr(i, 2) = rs("NavName").Value
        arr(i, 3) = rs("Description").Value
        arr(i, 4) = rs("Icon").Value
        arr(i, 5) = rs("RequiresReview").Value
        arr(i, 6) = rs("RequiresRatio").Value
        arr(i, 7) = rs("DisplayOrder").Value
        arr(i, 8) = rs("IsActive").Value
        i = i + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    GetActiveProductTypes = arr
End Function

' 获取所有商品类型（包括禁用的），返回二维数组，格式同上
Function GetAllProductTypes()
    Dim rs, arr, count, i
    Dim sql
    
    sql = "SELECT TypeCode, DisplayName, NavName, Description, Icon, RequiresReview, RequiresRatio, DisplayOrder, IsActive FROM ProductTypeConfig ORDER BY DisplayOrder ASC"
    
    Set rs = ExecuteQuery(sql)
    
    If rs Is Nothing Then
        GetAllProductTypes = Array()
        Exit Function
    End If
    If rs.EOF Then
        GetAllProductTypes = Array()
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    ' 获取记录数
    count = 0
    Do While Not rs.EOF
        count = count + 1
        rs.MoveNext
    Loop
    
    ' 返回第一条记录
    rs.MoveFirst
    
    ' 创建数组
    ReDim arr(count - 1, 8)
    
    ' 填充数组
    i = 0
    Do While Not rs.EOF
        arr(i, 0) = rs("TypeCode").Value
        arr(i, 1) = rs("DisplayName").Value
        arr(i, 2) = rs("NavName").Value
        arr(i, 3) = rs("Description").Value
        arr(i, 4) = rs("Icon").Value
        arr(i, 5) = rs("RequiresReview").Value
        arr(i, 6) = rs("RequiresRatio").Value
        arr(i, 7) = rs("DisplayOrder").Value
        arr(i, 8) = rs("IsActive").Value
        i = i + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    GetAllProductTypes = arr
End Function

' 获取指定TypeCode的显示名称，返回字符串
Function GetProductTypeDisplayName(typeCode)
    Dim rs
    Dim sql, result
    
    result = ""
    
    If IsNull(typeCode) Or typeCode = "" Then
        GetProductTypeDisplayName = result
        Exit Function
    End If
    
    sql = "SELECT DisplayName FROM ProductTypeConfig WHERE TypeCode='" & SafeSQL(typeCode) & "'"
    
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            result = rs("DisplayName").Value
            If IsNull(result) Then result = ""
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    GetProductTypeDisplayName = result
End Function

' 判断商品类型是否存在且启用，返回Boolean
Function IsProductTypeActive(typeCode)
    Dim rs
    Dim sql, result
    
    result = False
    
    If IsNull(typeCode) Or typeCode = "" Then
        IsProductTypeActive = result
        Exit Function
    End If
    
    sql = "SELECT TypeCode FROM ProductTypeConfig WHERE TypeCode='" & SafeSQL(typeCode) & "' AND IsActive<>0"
    
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            result = True
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    IsProductTypeActive = result
End Function

' 构建商品类型的WHERE过滤条件字符串
' 如果typeCode有值且该类型启用，返回 " AND ProductType='xxx'"
' 如果该类型需要审核(RequiresReview=1)，额外添加 " AND ReviewStatus='Approved'"
Function BuildProductTypeFilter(typeCode)
    Dim rs
    Dim sql, filterStr, requiresReview
    
    filterStr = ""
    
    If IsNull(typeCode) Or typeCode = "" Then
        BuildProductTypeFilter = filterStr
        Exit Function
    End If
    
    ' 检查类型是否存在且启用
    sql = "SELECT TypeCode, RequiresReview FROM ProductTypeConfig WHERE TypeCode='" & SafeSQL(typeCode) & "' AND IsActive<>0"
    
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            ' 类型存在且启用，添加类型过滤
            filterStr = " AND ProductType='" & SafeSQL(typeCode) & "'"
            
            ' 检查是否需要审核
            requiresReview = rs("RequiresReview").Value
            If Not IsNull(requiresReview) Then
                If CBool(requiresReview) = True Then
                    filterStr = filterStr & " AND ReviewStatus='Approved'"
                End If
            End If
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    BuildProductTypeFilter = filterStr
End Function

' 获取所有已启用类型的类型代码，返回逗号分隔字符串用于IN查询
' 例如 "'Fixed','Custom','KOL'"
Function GetActiveTypeCodesForSQL()
    Dim rs
    Dim sql, result, typeCodeList
    
    result = ""
    typeCodeList = ""
    
    sql = "SELECT TypeCode FROM ProductTypeConfig WHERE IsActive<>0 ORDER BY DisplayOrder ASC"
    
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If typeCodeList <> "" Then
                typeCodeList = typeCodeList & ","
            End If
            typeCodeList = typeCodeList & "'" & SafeSQL(rs("TypeCode").Value) & "'"
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    
    GetActiveTypeCodesForSQL = typeCodeList
End Function

' ============================================
' V17.2: i18n辅助 - 根据TypeCode获取翻译后的产品类型名称
' 参数:
'   typeCode     - 产品类型代码 (如 "Custom", "standard", "kol", "Fixed")
'   dbName       - 数据库中的名称 (作为回退值)
'   nameType     - "display" 获取完整显示名, "nav" 获取简短导航名
' 返回: 翻译后的名称，若翻译缺失则返回数据库原始名称
' ============================================
Function GetProductTypeI18nName(typeCode, dbName, nameType)
    Dim lcType, i18nKey, translated
    
    If IsNull(dbName) Or dbName = "" Then
        GetProductTypeI18nName = ""
        Exit Function
    End If
    
    ' i18n未启用时直接返回数据库名称
    If Not FEATURE_I18N Then
        GetProductTypeI18nName = dbName
        Exit Function
    End If
    
    lcType = LCase(typeCode & "")
    
    ' 构建i18n键名: product_type_<typecode> (display) 或 nav_<typecode> (nav)
    If nameType = "nav" Then
        i18nKey = "nav_" & lcType
        ' 先尝试nav键
        translated = T(i18nKey, Empty)
        ' 检查是否命中翻译 (未命中时返回 [key] 格式)
        If Left(translated, 1) = "[" And Right(translated, 1) = "]" Then
            ' 回退: 尝试 product_type_<typecode>
            i18nKey = "product_type_" & lcType
            translated = T(i18nKey, Empty)
            If Left(translated, 1) = "[" And Right(translated, 1) = "]" Then
                GetProductTypeI18nName = dbName
                Exit Function
            End If
        End If
    Else
        ' display 名称: 先尝试 product_type_<typecode>
        i18nKey = "product_type_" & lcType
        translated = T(i18nKey, Empty)
        If Left(translated, 1) = "[" And Right(translated, 1) = "]" Then
            ' 回退到数据库名称
            GetProductTypeI18nName = dbName
            Exit Function
        End If
    End If
    
    GetProductTypeI18nName = translated
End Function
%>
