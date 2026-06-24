<%
' ============================================
' V15.0 DAL - 产品数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_products.asp"-->
' ============================================

' ============================================
' 根据ID获取产品
' ============================================
Function DAL_Products_GetByID(productId)
    Dim sql, params(0)
    sql = "SELECT * FROM Products WHERE ProductID=@ProductID"
    params(0) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    Set DAL_Products_GetByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 获取产品列表（分页+筛选）
' 参数: search(可选), category(可选), productType(可选), isActive(可选)
' ============================================
Function DAL_Products_GetList(search, category, productType, isActiveOnly, page, pageSize, ByRef pageInfo)
    Dim sql, params(), paramCount, whereAdded
    
    sql = "SELECT p.ProductID, p.ProductName, p.BasePrice, p.Category, p.ProductType, " & _
          "p.ImageURL, p.IsActive, p.Description, p.UnitCost, p.CreatedAt, " & _
          "ISNULL(rp.PurchaseCount, 0) AS PurchaseCount " & _
          "FROM Products p " & _
          "LEFT JOIN RecipePopularity rp ON p.ProductID=rp.ProductID WHERE 1=1"
    
    paramCount = -1
    ReDim params(0)
    
    ' 搜索关键词
    If Not IsNull(search) And search <> "" Then
        sql = sql & " AND (p.ProductName LIKE '%' + @Search + '%' OR p.Description LIKE '%' + @Search + '%')"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Search", DAL_adVarChar, 200, search)
    End If
    
    ' 分类筛选
    If Not IsNull(category) And category <> "" Then
        sql = sql & " AND p.Category=@Category"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@Category", DAL_adVarChar, 50, category)
    End If
    
    ' 产品类型筛选
    If Not IsNull(productType) And productType <> "" Then
        sql = sql & " AND p.ProductType=@ProductType"
        paramCount = paramCount + 1
        ReDim Preserve params(paramCount)
        params(paramCount) = Array("@ProductType", DAL_adVarChar, 50, productType)
    End If
    
    ' 仅活跃产品
    If isActiveOnly Then
        sql = sql & " AND p.IsActive=1"
    End If
    
    sql = sql & " ORDER BY p.CreatedAt DESC, p.ProductID DESC"
    
    If paramCount >= 0 Then
        Set DAL_Products_GetList = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
    Else
        Set DAL_Products_GetList = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
    End If
End Function

' ============================================
' 获取活跃产品（简单列表，不分页）
' ============================================
Function DAL_Products_GetActive()
    Dim sql
    sql = "SELECT ProductID, ProductName, BasePrice, Category, ProductType, " & _
          "ImageURL, Description, IsActive FROM Products " & _
          "WHERE IsActive=1 ORDER BY ProductName ASC"
    Set DAL_Products_GetActive = DAL_GetList(sql, Null)
End Function

' ============================================
' 获取热门产品（按购买次数排序）
' ============================================
Function DAL_Products_GetHot(limit)
    Dim sql, params(0)
    If IsNull(limit) Or limit < 1 Then limit = 10
    sql = "SELECT TOP " & CLng(limit) & " p.ProductID, p.ProductName, p.BasePrice, " & _
          "p.Category, p.ProductType, p.ImageURL, p.Description, " & _
          "ISNULL(rp.PurchaseCount, 0) AS HotScore " & _
          "FROM Products p " & _
          "LEFT JOIN RecipePopularity rp ON p.ProductID=rp.ProductID " & _
          "WHERE p.IsActive=1 " & _
          "ORDER BY ISNULL(rp.PurchaseCount, 0) DESC, p.CreatedAt DESC"
    Set DAL_Products_GetHot = DAL_GetList(sql, Null)
End Function

' ============================================
' 根据KOL获取产品
' ============================================
Function DAL_Products_GetByKOL(kolId)
    Dim sql, params(0)
    sql = "SELECT * FROM Products WHERE KOLID=@KOLID AND IsActive=1"
    params(0) = Array("@KOLID", DAL_adInteger, 0, CLng(kolId))
    Set DAL_Products_GetByKOL = DAL_GetList(sql, params)
End Function

' ============================================
' 产品搜索（综合搜索：名称+描述+分类）
' ============================================
Function DAL_Products_Search(keyword, page, pageSize, ByRef pageInfo)
    Dim sql, params(0)
    sql = "SELECT ProductID, ProductName, BasePrice, Category, ProductType, " & _
          "ImageURL, Description FROM Products " & _
          "WHERE IsActive=1 AND (ProductName LIKE '%' + @Keyword + '%' " & _
          "OR Description LIKE '%' + @Keyword + '%' OR Category LIKE '%' + @Keyword + '%') " & _
          "ORDER BY ProductName ASC"
    params(0) = Array("@Keyword", DAL_adVarChar, 100, keyword)
    Set DAL_Products_Search = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
End Function

' ============================================
' 获取产品评论（分页）
' ============================================
Function DAL_Products_GetReviews(productId, page, pageSize, ByRef pageInfo)
    Dim sql, params(0)
    sql = "SELECT pr.ReviewID, pr.ProductID, pr.UserID, u.Username, " & _
          "pr.Rating, pr.Comment, pr.CreatedAt, pr.Status " & _
          "FROM ProductReviews pr " & _
          "LEFT JOIN Users u ON pr.UserID=u.UserID " & _
          "WHERE pr.ProductID=@ProductID AND pr.Status='Approved' " & _
          "ORDER BY pr.CreatedAt DESC"
    params(0) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    Set DAL_Products_GetReviews = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
End Function

' ============================================
' 获取产品平均评分
' ============================================
Function DAL_Products_GetAverageRating(productId)
    Dim sql, params(0)
    sql = "SELECT ISNULL(AVG(CAST(Rating AS FLOAT)), 0) FROM ProductReviews " & _
          "WHERE ProductID=@ProductID AND Status='Approved'"
    params(0) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    DAL_Products_GetAverageRating = CDbl(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 获取产品容量价格
' ============================================
Function DAL_Products_GetVolumePrices(productId)
    Dim sql, params(0)
    sql = "SELECT pvp.PVPriceID, pvp.ProductID, pvp.VolumeID, pvp.Price, " & _
          "v.VolumeML, v.VolumeName, v.PriceMultiplier " & _
          "FROM ProductVolumePrices pvp " & _
          "INNER JOIN Volumes v ON pvp.VolumeID=v.VolumeID AND v.IsActive=1 " & _
          "WHERE pvp.ProductID=@ProductID " & _
          "ORDER BY v.VolumeML ASC"
    params(0) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    Set DAL_Products_GetVolumePrices = DAL_GetList(sql, params)
End Function

' ============================================
' 获取产品香调（关联的FragranceNotes）
' ============================================
Function DAL_Products_GetFragranceNotes(productId)
    Dim sql, params(0)
    sql = "SELECT pn.ProductNoteID, pn.ProductID, pn.NoteID, " & _
          "fn.NoteName, fn.NoteType, fn.PriceAddition, fn.Description, fn.ImageURL, " & _
          "pnr.Percentage " & _
          "FROM ProductNotes pn " & _
          "INNER JOIN FragranceNotes fn ON pn.NoteID=fn.NoteID AND fn.IsActive=1 " & _
          "LEFT JOIN ProductNoteRatios pnr ON pn.ProductID=pnr.ProductID AND pn.NoteID=pnr.NoteID " & _
          "WHERE pn.ProductID=@ProductID " & _
          "ORDER BY CASE fn.NoteType WHEN 'Top' THEN 1 WHEN 'Middle' THEN 2 WHEN 'Base' THEN 3 ELSE 4 END"
    params(0) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    Set DAL_Products_GetFragranceNotes = DAL_GetList(sql, params)
End Function

' ============================================
' 获取产品库存
' ============================================
Function DAL_Products_GetStock(productId)
    Dim sql, params(0)
    sql = "SELECT ISNULL(SUM(StockQty), 0) FROM ProductInventory " & _
          "WHERE ProductID=@ProductID"
    params(0) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    DAL_Products_GetStock = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 更新产品库存
' ============================================
Function DAL_Products_UpdateStock(productId, quantity)
    Dim sql, params(1)
    sql = "UPDATE ProductInventory SET StockQty=@StockQty, UpdatedAt=GETDATE() " & _
          "WHERE ProductID=@ProductID"
    params(0) = Array("@StockQty", DAL_adInteger, 0, CLng(quantity))
    params(1) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    DAL_Products_UpdateStock = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 检查产品是否存在
' ============================================
Function DAL_Products_Exists(productId)
    DAL_Products_Exists = DAL_Exists("Products", "ProductID=@ProductID", _
        Array(Array("@ProductID", DAL_adInteger, 0, CLng(productId))))
End Function

' ============================================
' 获取用户收藏列表
' ============================================
Function DAL_Products_GetFavorites(userId, page, pageSize, ByRef pageInfo)
    Dim sql, params(0)
    sql = "SELECT uf.FavoriteID, uf.CreatedTime, " & _
          "p.ProductID, p.ProductName, p.BasePrice, p.Category, p.ProductType, p.ImageURL " & _
          "FROM UserFavorites uf " & _
          "INNER JOIN Products p ON uf.ProductID=p.ProductID AND p.IsActive=1 " & _
          "WHERE uf.UserID=@UserID " & _
          "ORDER BY uf.CreatedTime DESC"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    Set DAL_Products_GetFavorites = DAL_GetListPaged(sql, params, page, pageSize, pageInfo)
End Function

' ============================================
' 添加收藏
' ============================================
Function DAL_Products_AddFavorite(userId, productId)
    Dim sql, params(2)
    ' 先检查是否已收藏
    If DAL_Exists("UserFavorites", "UserID=@UserID AND ProductID=@ProductID", _
        Array(Array("@UserID", DAL_adInteger, 0, CLng(userId)), _
              Array("@ProductID", DAL_adInteger, 0, CLng(productId)))) Then
        DAL_Products_AddFavorite = True
        Exit Function
    End If
    
    Dim fields(1)
    fields(0) = "UserID"
    fields(1) = "ProductID"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    params(1) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    DAL_Products_AddFavorite = (DAL_Insert("UserFavorites", fields, params) > 0)
End Function

' ============================================
' 取消收藏
' ============================================
Function DAL_Products_RemoveFavorite(userId, productId)
    Dim sql, params(1)
    sql = "DELETE FROM UserFavorites WHERE UserID=@UserID AND ProductID=@ProductID"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    params(1) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    DAL_Products_RemoveFavorite = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 判断用户是否已收藏某产品
' ============================================
Function DAL_Products_IsFavorite(userId, productId)
    DAL_Products_IsFavorite = DAL_Exists("UserFavorites", "UserID=@UserID AND ProductID=@ProductID", _
        Array(Array("@UserID", DAL_adInteger, 0, CLng(userId)), _
              Array("@ProductID", DAL_adInteger, 0, CLng(productId))))
End Function

' ============================================
' 获取分类列表
' ============================================
Function DAL_Products_GetCategories()
    Dim sql
    sql = "SELECT CategoryID, CategoryName, SortOrder FROM Categories " & _
          "WHERE IsActive=1 ORDER BY SortOrder ASC, CategoryID ASC"
    Set DAL_Products_GetCategories = DAL_GetList(sql, Null)
End Function

' ============================================
' 获取产品类型配置
' ============================================
Function DAL_Products_GetProductTypes()
    Dim sql
    sql = "SELECT ConfigID, TypeCode, DisplayName, NavName, Icon, Description, " & _
          "RequiresRatio, RequiresReview, DisplayOrder " & _
          "FROM ProductTypeConfig WHERE IsActive=1 ORDER BY DisplayOrder ASC"
    Set DAL_Products_GetProductTypes = DAL_GetList(sql, Null)
End Function

' ============================================
' 获取容量选项列表
' ============================================
Function DAL_Products_GetVolumes()
    Dim sql
    sql = "SELECT VolumeID, VolumeML, VolumeName, PriceMultiplier FROM Volumes " & _
          "WHERE IsActive=1 ORDER BY VolumeML ASC"
    Set DAL_Products_GetVolumes = DAL_GetList(sql, Null)
End Function
%>