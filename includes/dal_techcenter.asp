<%
' ============================================
' V18.0 DAL - 技术中心数据访问层
' 依赖: dal.asp, connection.asp
' 用法: <!--#include file="dal_techcenter.asp"-->
' 涵盖：Recipes, RecipeNotes, RecipeIngredients,
'       RecipeAccords, RecipeAccordMaterials,
'       RecipeProducts, RecipeProductNotes,
'       RecipePublishLog, FragranceNotes, BaseNotes
' ============================================

' ============================================
' 配方 - 获取活跃配方列表
' ============================================
Function DAL_TC_GetActiveRecipes()
    Dim sql
    sql = "SELECT RecipeID, RecipeName, RecipeCode, ProductType FROM Recipes WHERE IsActive=1 ORDER BY RecipeCode"
    Set DAL_TC_GetActiveRecipes = DAL_GetList(sql, Null)
End Function

' ============================================
' 配方 - 按ID获取
' ============================================
Function DAL_TC_GetRecipeByID(recipeId)
    Dim sql, params(0)
    sql = "SELECT RecipeName, RecipeCode, ProductType FROM Recipes WHERE RecipeID=@RecipeID"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    Set DAL_TC_GetRecipeByID = DAL_GetRow(sql, params)
End Function

' ============================================
' RecipeIngredients - 检查配方是否有数据
' ============================================
Function DAL_TC_CountIngredients(recipeId)
    Dim sql, params(0)
    sql = "SELECT COUNT(*) FROM RecipeIngredients WHERE RecipeID=@RecipeID"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    DAL_TC_CountIngredients = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' RecipeIngredients - 按配方+香调统计数量
' ============================================
Function DAL_TC_CountIngredientsByNote(recipeId, noteId)
    Dim sql, params(1)
    sql = "SELECT COUNT(*) FROM RecipeIngredients WHERE RecipeID=@RecipeID AND NoteID=@NoteID"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    params(1) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    DAL_TC_CountIngredientsByNote = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' RecipeNotes - 获取配方中的香调列表（用于发布选择）
' ============================================
Function DAL_TC_GetRecipeNotesForAccord(recipeId)
    Dim sql, params(0)
    sql = "SELECT DISTINCT rn.NoteID, fn.NoteName, fn.NoteType FROM RecipeNotes rn " & _
          "INNER JOIN FragranceNotes fn ON rn.NoteID=fn.NoteID " & _
          "WHERE rn.RecipeID=@RecipeID ORDER BY fn.NoteType, fn.NoteName"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    Set DAL_TC_GetRecipeNotesForAccord = DAL_GetList(sql, params)
End Function

' ============================================
' RecipeIngredients - 获取配方中某香调的原材料成分
' ============================================
Function DAL_TC_GetAccordIngredients(recipeId, noteId)
    Dim sql, params(1)
    sql = "SELECT IngredientName, Percentage FROM RecipeIngredients " & _
          "WHERE RecipeID=@RecipeID AND NoteID=@NoteID ORDER BY ID"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    params(1) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    Set DAL_TC_GetAccordIngredients = DAL_GetList(sql, params)
End Function

' ============================================
' RawMaterialInventory - 按名称匹配原料
' ============================================
Function DAL_TC_MatchRawMaterial(itemName)
    Dim sql, params(0)
    sql = "SELECT TOP 1 MaterialID FROM RawMaterialInventory WHERE ItemName=@ItemName"
    params(0) = Array("@ItemName", DAL_adVarChar, 200, itemName)
    DAL_TC_MatchRawMaterial = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' RecipeNotes - 按百分比获取配方香调（用于产品配方发布）
' ============================================
Function DAL_TC_GetRecipeNotesForProduct(recipeId)
    Dim sql, params(0)
    sql = "SELECT rn.NoteID, rn.Percentage, fn.NoteName, fn.NoteType FROM RecipeNotes rn " & _
          "INNER JOIN FragranceNotes fn ON rn.NoteID=fn.NoteID " & _
          "WHERE rn.RecipeID=@RecipeID ORDER BY rn.ID"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    Set DAL_TC_GetRecipeNotesForProduct = DAL_GetList(sql, params)
End Function

' ============================================
' RecipeNotes - 检查百分比总和
' ============================================
Function DAL_TC_GetNotesPercentSum(recipeId)
    Dim sql, params(0)
    sql = "SELECT SUM(Percentage) FROM RecipeNotes WHERE RecipeID=@RecipeID"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    DAL_TC_GetNotesPercentSum = CDbl(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' RecipeAccords - 创建发布记录
' ============================================
Function DAL_TC_CreateAccord(recipeId, noteId, batchSize, publishedBy, recipeName)
    Dim sql, params(4)
    sql = "INSERT INTO RecipeAccords (RecipeID, NoteID, BatchSize, Status, PublishedBy, PublishedAt, CreatedAt, RecipeName) " & _
          "VALUES (@RecipeID, @NoteID, @BatchSize, 'Published', @PublishedBy, GETDATE(), GETDATE(), @RecipeName)"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    params(1) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    params(2) = Array("@BatchSize", DAL_adInteger, 0, CLng(batchSize))
    params(3) = Array("@PublishedBy", DAL_adVarChar, 50, publishedBy)
    params(4) = Array("@RecipeName", DAL_adVarChar, 255, Left(recipeName, 255))
    DAL_TC_CreateAccord = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeAccordMaterials - 创建原材料明细
' ============================================
Function DAL_TC_CreateAccordMaterial(accordRecipeId, materialId, materialName, percentage, plannedQty)
    Dim sql, params(4)
    sql = "INSERT INTO RecipeAccordMaterials (AccordRecipeID, MaterialID, MaterialName, Percentage, PlannedQty) " & _
          "VALUES (@AccordRecipeID, @MaterialID, @MaterialName, @Percentage, @PlannedQty)"
    params(0) = Array("@AccordRecipeID", DAL_adInteger, 0, CLng(accordRecipeId))
    params(1) = Array("@MaterialID", DAL_adInteger, 0, CLng(materialId))
    params(2) = Array("@MaterialName", DAL_adVarChar, 200, Left(materialName, 200))
    params(3) = Array("@Percentage", DAL_adDouble, 0, CDbl(percentage))
    params(4) = Array("@PlannedQty", DAL_adDouble, 0, CDbl(plannedQty))
    DAL_TC_CreateAccordMaterial = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeProducts - 创建产品配方发布记录
' ============================================
Function DAL_TC_CreateProduct(recipeId, productId, batchSize, publishedBy)
    Dim sql, params(3)
    sql = "INSERT INTO RecipeProducts (RecipeID, ProductID, BatchSize, Status, PublishedBy, PublishedAt, CreatedAt) " & _
          "VALUES (@RecipeID, @ProductID, @BatchSize, 'Published', @PublishedBy, GETDATE(), GETDATE())"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    params(1) = Array("@ProductID", DAL_adInteger, 0, CLng(productId))
    params(2) = Array("@BatchSize", DAL_adInteger, 0, CLng(batchSize))
    params(3) = Array("@PublishedBy", DAL_adVarChar, 50, publishedBy)
    DAL_TC_CreateProduct = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeProductNotes - 创建产品香调明细
' ============================================
Function DAL_TC_CreateProductNote(productRecipeId, noteId, noteName, percentage, plannedQty)
    Dim sql, params(4)
    sql = "INSERT INTO RecipeProductNotes (ProductRecipeID, NoteID, NoteName, Percentage, PlannedQty) " & _
          "VALUES (@ProductRecipeID, @NoteID, @NoteName, @Percentage, @PlannedQty)"
    params(0) = Array("@ProductRecipeID", DAL_adInteger, 0, CLng(productRecipeId))
    params(1) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    params(2) = Array("@NoteName", DAL_adVarChar, 200, Left(noteName, 200))
    params(3) = Array("@Percentage", DAL_adDouble, 0, CDbl(percentage))
    params(4) = Array("@PlannedQty", DAL_adDouble, 0, CDbl(plannedQty))
    DAL_TC_CreateProductNote = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipePublishLog - 创建审计日志
' ============================================
Function DAL_TC_CreatePublishLog(recipeId, publishType, targetRecipeId, publishedBy, ipAddress)
    Dim sql, params(4)
    sql = "INSERT INTO RecipePublishLog (RecipeID, PublishType, TargetRecipeID, PublishedBy, PublishedAt, IPAddress) " & _
          "VALUES (@RecipeID, @PublishType, @TargetRecipeID, @PublishedBy, GETDATE(), @IPAddress)"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    params(1) = Array("@PublishType", DAL_adVarChar, 20, Left(publishType, 20))
    params(2) = Array("@TargetRecipeID", DAL_adInteger, 0, CLng(targetRecipeId))
    params(3) = Array("@PublishedBy", DAL_adVarChar, 50, publishedBy)
    params(4) = Array("@IPAddress", DAL_adVarChar, 50, Left(ipAddress, 50))
    DAL_TC_CreatePublishLog = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeAccords - 废弃旧版本
' ============================================
Function DAL_TC_DeprecateAccord(accordRecipeId, recipeId, noteId)
    Dim sql, params(2)
    sql = "UPDATE RecipeAccords SET Status='Deprecated' WHERE AccordRecipeID<>@AccordRecipeID " & _
          "AND RecipeID=@RecipeID AND NoteID=@NoteID AND Status='Published'"
    params(0) = Array("@AccordRecipeID", DAL_adInteger, 0, CLng(accordRecipeId))
    params(1) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    params(2) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    DAL_TC_DeprecateAccord = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeAccords - 废弃单个
' ============================================
Function DAL_TC_DeprecateAccordSingle(accordRecipeId)
    Dim sql, params(0)
    sql = "UPDATE RecipeAccords SET Status='Deprecated' WHERE AccordRecipeID=@AccordRecipeID"
    params(0) = Array("@AccordRecipeID", DAL_adInteger, 0, CLng(accordRecipeId))
    DAL_TC_DeprecateAccordSingle = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeProducts - 废弃旧版本
' ============================================
Function DAL_TC_DeprecateProductVersions(productRecipeId, recipeId)
    Dim sql, params(1)
    sql = "UPDATE RecipeProducts SET Status='Deprecated' WHERE ProductRecipeID<>@ProductRecipeID " & _
          "AND RecipeID=@RecipeID AND Status='Published'"
    params(0) = Array("@ProductRecipeID", DAL_adInteger, 0, CLng(productRecipeId))
    params(1) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    DAL_TC_DeprecateProductVersions = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeProducts - 废弃单个
' ============================================
Function DAL_TC_DeprecateProductSingle(productRecipeId)
    Dim sql, params(0)
    sql = "UPDATE RecipeProducts SET Status='Deprecated' WHERE ProductRecipeID=@ProductRecipeID"
    params(0) = Array("@ProductRecipeID", DAL_adInteger, 0, CLng(productRecipeId))
    DAL_TC_DeprecateProductSingle = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' RecipeAccords - 获取已发布列表
' ============================================
Function DAL_TC_GetPublishedAccords()
    Dim sql
    sql = "SELECT ra.AccordRecipeID, ra.BatchSize, ra.Status, ra.PublishedBy, ra.PublishedAt, ra.RecipeName, " & _
          "r.RecipeName AS FullRecipeName, fn.NoteName FROM RecipeAccords ra " & _
          "LEFT JOIN Recipes r ON ra.RecipeID=r.RecipeID " & _
          "LEFT JOIN FragranceNotes fn ON ra.NoteID=fn.NoteID " & _
          "ORDER BY ra.PublishedAt DESC"
    Set DAL_TC_GetPublishedAccords = DAL_GetList(sql, Null)
End Function

' ============================================
' RecipeProducts - 获取已发布列表
' ============================================
Function DAL_TC_GetPublishedProducts()
    Dim sql
    sql = "SELECT rp.ProductRecipeID, rp.BatchSize, rp.Status, rp.PublishedBy, rp.PublishedAt, " & _
          "r.RecipeName, p.ProductName FROM RecipeProducts rp " & _
          "LEFT JOIN Recipes r ON rp.RecipeID=r.RecipeID " & _
          "LEFT JOIN Products p ON rp.ProductID=p.ProductID " & _
          "ORDER BY rp.PublishedAt DESC"
    Set DAL_TC_GetPublishedProducts = DAL_GetList(sql, Null)
End Function

' ============================================
' RecipePublishLog - 获取最近日志
' ============================================
Function DAL_TC_GetRecentPublishLogs(topCount)
    Dim sql
    If IsNull(topCount) Or topCount <= 0 Then topCount = 20
    sql = "SELECT TOP " & CLng(topCount) & " PublishedAt, PublishType, TargetRecipeID, PublishedBy, IPAddress " & _
          "FROM RecipePublishLog ORDER BY PublishedAt DESC"
    Set DAL_TC_GetRecentPublishLogs = DAL_GetList(sql, Null)
End Function

' ============================================
' 统计 - 已发布香调配方数
' ============================================
Function DAL_TC_CountPublishedAccords()
    DAL_TC_CountPublishedAccords = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM RecipeAccords WHERE Status='Published'", Null, 0))
End Function

' ============================================
' 统计 - 已发布产品配方数
' ============================================
Function DAL_TC_CountPublishedProducts()
    DAL_TC_CountPublishedProducts = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM RecipeProducts WHERE Status='Published'", Null, 0))
End Function

' ============================================
' 统计 - 发布日志总数
' ============================================
Function DAL_TC_CountPublishLogs()
    DAL_TC_CountPublishLogs = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM RecipePublishLog", Null, 0))
End Function

' ============================================
' 统计 - 活跃配方总数
' ============================================
Function DAL_TC_CountActiveRecipes()
    DAL_TC_CountActiveRecipes = CLng(DAL_GetScalar( _
        "SELECT COUNT(*) FROM Recipes WHERE IsActive=1", Null, 0))
End Function

' ============================================
' 统计 - 配方中某香调的RecipeAccord发布数量
' ============================================
Function DAL_TC_CountAccordsByRecipe(recipeId)
    Dim sql, params(0)
    sql = "SELECT COUNT(*) FROM RecipeAccords WHERE RecipeID=@RecipeID AND Status='Published'"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    DAL_TC_CountAccordsByRecipe = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 统计 - 配方中某产品的RecipeProduct发布数量
' ============================================
Function DAL_TC_CountProductsByRecipe(recipeId)
    Dim sql, params(0)
    sql = "SELECT COUNT(*) FROM RecipeProducts WHERE RecipeID=@RecipeID AND Status='Published'"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    DAL_TC_CountProductsByRecipe = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 统计 - 配方中的香调数量
' ============================================
Function DAL_TC_CountNotesByRecipe(recipeId)
    Dim sql, params(0)
    sql = "SELECT COUNT(*) FROM RecipeNotes WHERE RecipeID=@RecipeID"
    params(0) = Array("@RecipeID", DAL_adInteger, 0, CLng(recipeId))
    DAL_TC_CountNotesByRecipe = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 产品 - 获取所有产品列表（用于下拉选择）
' ============================================
Function DAL_TC_GetAllProducts()
    Dim sql
    sql = "SELECT ProductID, ProductName FROM Products ORDER BY ProductName"
    Set DAL_TC_GetAllProducts = DAL_GetList(sql, Null)
End Function

' ============================================
' 获取最后插入的IDENTITY值
' ============================================
Function DAL_TC_GetLastIdentity()
    DAL_TC_GetLastIdentity = CLng(DAL_GetScalar("SELECT SCOPE_IDENTITY()", Null, 0))
End Function

' ============================================
' NoteInventory - 更新香调库存（DAL参数化版本）
' ============================================
Function DAL_TC_UpdateNoteStock(noteId, stockQty, minStockLevel)
    Dim sql, params(2)
    sql = "UPDATE NoteInventory SET StockQuantity=@StockQty, MinStockLevel=@MinLevel, UpdatedAt=GETDATE() WHERE NoteID=@NoteID"
    params(0) = Array("@StockQty", DAL_adInteger, 0, CLng(stockQty))
    params(1) = Array("@MinLevel", DAL_adInteger, 0, CLng(minStockLevel))
    params(2) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    DAL_TC_UpdateNoteStock = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' NoteInventory - 入库增加库存（原子操作）
' ============================================
Function DAL_TC_RestockNote(noteId, addQty)
    Dim sql, params(1)
    sql = "UPDATE NoteInventory SET StockQuantity=StockQuantity+@AddQty, LastRestockDate=GETDATE(), UpdatedAt=GETDATE() WHERE NoteID=@NoteID"
    params(0) = Array("@AddQty", DAL_adInteger, 0, CLng(addQty))
    params(1) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    DAL_TC_RestockNote = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' InventoryTransactions - 创建交易记录
' ============================================
Function DAL_TC_CreateInvTransaction(noteId, quantity, transType, direction, notes, createdBy)
    Dim sql, params(5)
    sql = "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) " & _
          "VALUES (@NoteID, @Quantity, @TransType, @Direction, @Notes, @CreatedBy, GETDATE())"
    params(0) = Array("@NoteID", DAL_adInteger, 0, CLng(noteId))
    params(1) = Array("@Quantity", DAL_adInteger, 0, CLng(quantity))
    params(2) = Array("@TransType", DAL_adVarChar, 30, Left(transType, 30))
    params(3) = Array("@Direction", DAL_adVarChar, 20, Left(direction, 20))
    params(4) = Array("@Notes", DAL_adVarChar, 500, Left(notes, 500))
    params(5) = Array("@CreatedBy", DAL_adVarChar, 50, Left(createdBy, 50))
    DAL_TC_CreateInvTransaction = (DAL_Execute(sql, params) >= 0)
End Function
%>
