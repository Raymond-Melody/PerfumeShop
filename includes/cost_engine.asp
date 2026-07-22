<%
' ============================================
' 成本自动传导引擎 - Cost Engine
' 实现三级成本自动计算与传导：
' 原材料 → 香调(Accord/Note) → 产品(Product) → 订单(Order)
' ============================================

' ============================================
' 性能优化：批量预加载全局缓存字典
' 使用方式：在页面顶部（OpenConnection之后）调用 CE_PreloadAllCostData()
' 后续使用 CE_GetCached* 系列函数代替原 CE_Calculate* 函数
' ============================================

' 基础数据字典
Dim CE_MaterialPrices       ' MaterialID → UnitPrice
Dim CE_BaseNotePrices       ' BaseNoteID → UnitPrice (V9: 基香单价)
Dim CE_AccordRecipes        ' NoteID → "AccordRecipeID|BatchSize"
Dim CE_AccordMaterials      ' AccordRecipeID → "MatID1:Pct1:Qty1,MatID2:Pct2:Qty2,..."
Dim CE_NoteIngredients      ' NoteID → "BaseID1:Pct1,BaseID2:Pct2,..."
Dim CE_NotePriceAdditions   ' NoteID → PriceAddition
Dim CE_ProductNoteRatios    ' ProductID → "NoteID1:Pct1,NoteID2:Pct2,..."
Dim CE_BottleAdditions      ' ProductID → PriceAddition
Dim CE_ProductExtraCosts    ' ProductID → "PackagingCost|OtherCost"

' V10: 批次加权成本字典 (ItemCode → WeightedUnitCost)
Dim CE_WeightedBatchCosts   ' ItemCode → weighted average UnitCost

' 结果缓存字典（避免同一对象重复计算）
Dim CE_NoteCostCache        ' NoteID → computed cost
Dim CE_ProductBOMCache      ' ProductID → computed BOM cost
Dim CE_ProductUnitCache     ' ProductID → computed unit cost

' 品牌定香(Fixed)采购成本缓存
Dim CE_FixedBrandCosts      ' ProductID → AvgUnitCost (from FixedBrandInventory)

' 统计缓存
Dim CE_Stats                ' Key → Value (various COUNT stats)

' ============================================
' 批量预加载所有成本参考数据（仅需调用一次）
' 加载：原料价格、Accord配方、原料配比、成分聚合、香调配比、瓶身成本、包装/人工分摊
' ============================================
Sub CE_PreloadAllCostData()
    Dim rs, key, val, itemCode, matPrice, existing, parts, costType, costVal
    Dim rsCount
    
    On Error Resume Next
    
    ' 初始化所有字典
    Set CE_MaterialPrices = CreateObject("Scripting.Dictionary")
    Set CE_BaseNotePrices = CreateObject("Scripting.Dictionary")
    Set CE_AccordRecipes = CreateObject("Scripting.Dictionary")
    Set CE_AccordMaterials = CreateObject("Scripting.Dictionary")
    Set CE_NoteIngredients = CreateObject("Scripting.Dictionary")
    Set CE_NotePriceAdditions = CreateObject("Scripting.Dictionary")
    Set CE_ProductNoteRatios = CreateObject("Scripting.Dictionary")
    Set CE_BottleAdditions = CreateObject("Scripting.Dictionary")
    Set CE_ProductExtraCosts = CreateObject("Scripting.Dictionary")
    Set CE_NoteCostCache = CreateObject("Scripting.Dictionary")
    Set CE_ProductBOMCache = CreateObject("Scripting.Dictionary")
    Set CE_ProductUnitCache = CreateObject("Scripting.Dictionary")
    Set CE_Stats = CreateObject("Scripting.Dictionary")
    Set CE_WeightedBatchCosts = CreateObject("Scripting.Dictionary")
    Set CE_FixedBrandCosts = CreateObject("Scripting.Dictionary")
    
    ' === 1. 加载最新采购价 (ItemCode → UnitPrice) ===
    Dim CE_SPPrices
    Set CE_SPPrices = CreateObject("Scripting.Dictionary")
    
    Set rs = conn.Execute("SELECT sp.ItemCode, sp.UnitPrice FROM SupplierPrices sp INNER JOIN (SELECT ItemCode, MAX(CreatedAt) AS MaxCreated FROM SupplierPrices WHERE IsActive=1 GROUP BY ItemCode) latest ON sp.ItemCode = latest.ItemCode AND sp.CreatedAt = latest.MaxCreated WHERE sp.IsActive=1")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("ItemCode") & "")
            If key <> "" Then
                val = CE_SafeNum(rs("UnitPrice"))
                If Not CE_SPPrices.Exists(key) Then CE_SPPrices.Add key, val
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 2. 加载原料信息并构建 MaterialPrices (V10: 优先使用加权成本) ===
    Set rs = conn.Execute("SELECT MaterialID, ItemCode, UnitPrice, ISNULL(WeightedUnitCost,0) as WUC FROM RawMaterialInventory WHERE StockQty > 0")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("MaterialID"))
            If key <> "" And key <> "0" Then
                itemCode = CStr(rs("ItemCode") & "")
                ' V10: 优先使用加权平均成本，否则使用供应商最新报价
                Dim wuc : wuc = CE_SafeNum(rs("WUC"))
                If wuc > 0 Then
                    matPrice = wuc
                ElseIf itemCode <> "" And CE_SPPrices.Exists(itemCode) Then
                    matPrice = CE_SPPrices(itemCode)
                Else
                    matPrice = CE_SafeNum(rs("UnitPrice"))
                End If
                If Not CE_MaterialPrices.Exists(key) Then CE_MaterialPrices.Add key, matPrice
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    Set CE_SPPrices = Nothing
    
    ' === 2.5. V10: 加载跨品类加权批次成本 (ItemCode → WeightedUnitCost) ===
    Call CE_PreloadBatchCosts()
    
    ' === 3. 加载Accord配方 (NoteID → AccordRecipeID|BatchSize) ===
    Set rs = conn.Execute("SELECT ra.NoteID, ra.AccordRecipeID, ra.BatchSize FROM RecipeAccords ra INNER JOIN (SELECT NoteID, MAX(PublishedAt) AS MaxPub FROM RecipeAccords WHERE Status='Published' GROUP BY NoteID) latest ON ra.NoteID = latest.NoteID AND ra.PublishedAt = latest.MaxPub WHERE ra.Status='Published'")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("NoteID"))
            If key <> "" And key <> "0" Then
                val = CStr(rs("AccordRecipeID")) & "|" & CStr(CE_SafeNum(rs("BatchSize")))
                If Not CE_AccordRecipes.Exists(key) Then CE_AccordRecipes.Add key, val
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 4. 加载Accord原料配比 (AccordRecipeID → "MatID:Pct:Qty,...") ===
    Set rs = conn.Execute("SELECT AccordRecipeID, MaterialID, Percentage, PlannedQty FROM RecipeAccordMaterials ORDER BY AccordRecipeID")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("AccordRecipeID"))
            If key <> "" And key <> "0" Then
                val = CStr(rs("MaterialID")) & ":" & CStr(CE_SafeNum(rs("Percentage"))) & ":" & CStr(CE_SafeNum(rs("PlannedQty")))
                If CE_AccordMaterials.Exists(key) Then
                    CE_AccordMaterials(key) = CE_AccordMaterials(key) & "," & val
                Else
                    CE_AccordMaterials.Add key, val
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 2.5. V9: 加载 BaseNotes 单价 (BaseNoteID → UnitPrice) ===
    Set rs = conn.Execute("SELECT BaseNoteID, UnitPrice FROM BaseNotes WHERE IsActive <> 0")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("BaseNoteID"))
            If key <> "" And key <> "0" Then
                val = CE_SafeNum(rs("UnitPrice"))
                If val > 0 And Not CE_BaseNotePrices.Exists(key) Then
                    CE_BaseNotePrices.Add key, val
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 5. 加载NoteIngredients (NoteID → "BaseNoteID:Pct,...") ===
    Set rs = conn.Execute("SELECT NoteID, BaseNoteID, Percentage FROM NoteIngredients ORDER BY NoteID")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("NoteID"))
            If key <> "" And key <> "0" Then
                val = CStr(rs("BaseNoteID")) & ":" & CStr(CE_SafeNum(rs("Percentage")))
                If CE_NoteIngredients.Exists(key) Then
                    CE_NoteIngredients(key) = CE_NoteIngredients(key) & "," & val
                Else
                    CE_NoteIngredients.Add key, val
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 6. 加载FragranceNotes PriceAddition ===
    Set rs = conn.Execute("SELECT NoteID, PriceAddition FROM FragranceNotes WHERE IsActive=1")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("NoteID"))
            If key <> "" And key <> "0" Then
                val = CE_SafeNum(rs("PriceAddition"))
                If Not CE_NotePriceAdditions.Exists(key) Then CE_NotePriceAdditions.Add key, val
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 7. 加载ProductNoteRatios (ProductID → "NoteID:Pct,...") ===
    Set rs = conn.Execute("SELECT ProductID, NoteID, Percentage FROM ProductNoteRatios ORDER BY ProductID")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("ProductID"))
            If key <> "" And key <> "0" Then
                val = CStr(rs("NoteID")) & ":" & CStr(CE_SafeNum(rs("Percentage")))
                If CE_ProductNoteRatios.Exists(key) Then
                    CE_ProductNoteRatios(key) = CE_ProductNoteRatios(key) & "," & val
                Else
                    CE_ProductNoteRatios.Add key, val
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 8. 加载BottleStyles PriceAddition (ProductID → PriceAddition) ===
    Set rs = conn.Execute("SELECT pbs.ProductID, bs.PriceAddition FROM ProductBottleStyles pbs LEFT JOIN BottleStyles bs ON pbs.BottleID = bs.BottleID WHERE bs.IsActive = 1")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("ProductID"))
            If key <> "" And key <> "0" Then
                val = CE_SafeNum(rs("PriceAddition"))
                If Not CE_BottleAdditions.Exists(key) Then CE_BottleAdditions.Add key, val
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 9. 加载ProductCosts (ProductID → "PackagingCost|OtherCost") ===
    Set rs = conn.Execute("SELECT ProductID, CostType, SUM(TotalCost) AS Total FROM ProductCosts GROUP BY ProductID, CostType ORDER BY ProductID")
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            key = CStr(rs("ProductID"))
            If key <> "" And key <> "0" Then
                costType = CStr(rs("CostType") & "")
                costVal = CE_SafeNum(rs("Total"))
                If CE_ProductExtraCosts.Exists(key) Then
                    existing = CE_ProductExtraCosts(key)
                Else
                    existing = "0|0"
                End If
                parts = Split(existing, "|")
                If costType = "Packaging" Then
                    CE_ProductExtraCosts(key) = CStr(costVal) & "|" & parts(1)
                ElseIf costType = "Other" Then
                    CE_ProductExtraCosts(key) = parts(0) & "|" & CStr(costVal)
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' === 9.5. 预加载品牌定香(Fixed)采购成本 (ProductID → AvgUnitCost) ===
    On Error Resume Next
    Dim rsFB
    Set rsFB = conn.Execute("SELECT fbp.ProductID, ISNULL(fbi.AvgUnitCost, fbp.UnitPrice) AS Cost FROM FixedBrandProducts fbp LEFT JOIN FixedBrandInventory fbi ON fbp.FixedProductID = fbi.FixedProductID WHERE fbp.ProductID IS NOT NULL AND fbp.ProductID > 0 AND fbp.Status = 'Active'")
    If Err.Number = 0 Then
        If Not rsFB Is Nothing Then
            Do While Not rsFB.EOF
                Dim fbKey : fbKey = CStr(rsFB("ProductID"))
                Dim fbCost : fbCost = CE_SafeNum(rsFB("Cost"))
                If fbKey <> "" And fbCost > 0 And Not CE_FixedBrandCosts.Exists(fbKey) Then
                    CE_FixedBrandCosts.Add fbKey, fbCost
                End If
                rsFB.MoveNext
            Loop
            rsFB.Close
        End If
    Else
        Err.Clear
    End If
    Set rsFB = Nothing
    On Error GoTo 0
    
    ' === 10. 预加载统计数据 ===
    Set rsCount = conn.Execute("SELECT COUNT(*) FROM Products WHERE IsActive=1 AND UnitCost > 0")
    If Not rsCount Is Nothing Then CE_Stats.Add "updatedProducts", CE_SafeNum(rsCount(0)) : rsCount.Close
    Set rsCount = Nothing
    
    Set rsCount = conn.Execute("SELECT COUNT(*) FROM Products WHERE IsActive=1")
    If Not rsCount Is Nothing Then CE_Stats.Add "totalProducts", CE_SafeNum(rsCount(0)) : rsCount.Close
    Set rsCount = Nothing
    
    Set rsCount = conn.Execute("SELECT COUNT(*) FROM Orders WHERE CostAmount > 0")
    If Not rsCount Is Nothing Then CE_Stats.Add "updatedOrders", CE_SafeNum(rsCount(0)) : rsCount.Close
    Set rsCount = Nothing
    
    Set rsCount = conn.Execute("SELECT COUNT(*) FROM Orders WHERE Status NOT IN ('Pending','Cancelled')")
    If Not rsCount Is Nothing Then CE_Stats.Add "totalValidOrders", CE_SafeNum(rsCount(0)) : rsCount.Close
    Set rsCount = Nothing
    
    Set rsCount = conn.Execute("SELECT COUNT(*) FROM Orders")
    If Not rsCount Is Nothing Then CE_Stats.Add "allOrders", CE_SafeNum(rsCount(0)) : rsCount.Close
    Set rsCount = Nothing
    
    Set rsCount = conn.Execute("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty > 0")
    If Not rsCount Is Nothing Then CE_Stats.Add "rawMaterials", CE_SafeNum(rsCount(0)) : rsCount.Close
    Set rsCount = Nothing
    
    On Error GoTo 0
End Sub

' ============================================
' V10: 预加载跨品类加权批次成本
' 来源：RawMaterialInventory.WeightedUnitCost 及各类库存的 WeightedUnitCost
' ============================================
Sub CE_PreloadBatchCosts()
    Dim rs, key, val, wuc
    On Error Resume Next
    
    ' 加载所有库存品类的加权成本（ItemCode → WeightedUnitCost）
    Dim sqlUnion
    sqlUnion = "SELECT ItemCode, ISNULL(WeightedUnitCost,ISNULL(UnitPrice,0)) as WUC, 'RawMaterial' as ItemType FROM RawMaterialInventory WHERE StockQty > 0 " & _
             "UNION ALL SELECT ISNULL(ItemCode,''), ISNULL(WeightedUnitCost,ISNULL(UnitPrice,0)), 'Packaging' FROM PackagingInventory WHERE StockQty > 0 AND IsActive=1 " & _
             "UNION ALL SELECT ISNULL(ItemCode,''), ISNULL(WeightedUnitCost,ISNULL(UnitPrice,0)), 'Bottle' FROM BottleStyles WHERE StockQty > 0 " & _
             "UNION ALL SELECT ISNULL(ItemCode,''), ISNULL(WeightedUnitCost,ISNULL(UnitPrice,0)), 'Printing' FROM PrintingInventory WHERE StockQty > 0 AND IsActive=1 " & _
             "UNION ALL SELECT ISNULL(ItemCode,''), ISNULL(WeightedUnitCost,ISNULL(UnitPrice,0)), 'SprayHead' FROM SprayHeadInventory WHERE StockQty > 0 AND IsActive=1"
    
    ' 尝试加载印刷品和喷头（表可能还不存在）
    On Error Resume Next
    Set rs = conn.Execute(sqlUnion)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            Do While Not rs.EOF
                key = CStr(rs("ItemCode") & "")
                If key <> "" Then
                    wuc = CE_SafeNum(rs("WUC"))
                    If wuc > 0 Then
                        If Not CE_WeightedBatchCosts.Exists(key) Then
                            CE_WeightedBatchCosts.Add key, wuc
                        End If
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing
    
End Sub

' ============================================
' 品牌定香(Fixed): 从缓存获取采购加权平均成本
' 优先使用 FixedBrandInventory.AvgUnitCost，回退到 FixedBrandProducts.UnitPrice
' ============================================
Function CE_GetCachedFixedBrandCost(productId)
    Dim key : key = CStr(productId)
    
    If IsObject(CE_FixedBrandCosts) Then
        If CE_FixedBrandCosts.Exists(key) Then
            CE_GetCachedFixedBrandCost = CE_FixedBrandCosts(key)
            Exit Function
        End If
    End If
    
    ' 缓存未命中，实时查询
    Dim fbCost : fbCost = 0
    On Error Resume Next
    Dim rsFB2
    Set rsFB2 = conn.Execute("SELECT ISNULL(fbi.AvgUnitCost, fbp.UnitPrice) AS Cost FROM FixedBrandProducts fbp LEFT JOIN FixedBrandInventory fbi ON fbp.FixedProductID = fbi.FixedProductID WHERE fbp.ProductID=" & productId & " AND fbp.Status='Active'")
    If Err.Number = 0 Then
        If Not rsFB2 Is Nothing And Not rsFB2.EOF Then
            fbCost = CE_SafeNum(rsFB2("Cost"))
            rsFB2.Close
            If IsObject(CE_FixedBrandCosts) And fbCost > 0 Then CE_FixedBrandCosts.Add key, fbCost
        End If
    Else
        Err.Clear
    End If
    Set rsFB2 = Nothing
    On Error GoTo 0
    
    CE_GetCachedFixedBrandCost = fbCost
End Function

' ============================================
' V10: 获取指定物料的加权批次成本
' ============================================
Function CE_GetCachedBatchCost(itemCode)
    Dim key
    key = CStr(itemCode)
    If IsObject(CE_WeightedBatchCosts) Then
        If CE_WeightedBatchCosts.Exists(key) Then
            CE_GetCachedBatchCost = CE_WeightedBatchCosts(key)
            Exit Function
        End If
    End If
    CE_GetCachedBatchCost = 0
End Function

' ============================================
' 缓存版：获取原料成本（O(1) 内存查找，零DB查询）
' ============================================
Function CE_GetCachedMaterialCost(materialId)
    Dim key
    key = CStr(materialId)
    If IsObject(CE_MaterialPrices) Then
        If CE_MaterialPrices.Exists(key) Then
            CE_GetCachedMaterialCost = CE_MaterialPrices(key)
            Exit Function
        End If
    End If
    CE_GetCachedMaterialCost = 0
End Function

' ============================================
' 缓存版：获取香调成本（首次计算后缓存，零DB查询）
' ============================================
Function CE_GetCachedNoteCost(noteId)
    Dim key, totalCost, hasAccord
    Dim accordData, accordParts, accordId, batchSize
    Dim matList, matItems, matData, matParts, matId, matPct, matUnitCost
    Dim ingList, ingItems, ingData, ingParts, baseNoteId, ingPct, baseCost
    Dim i, j
    
    key = CStr(noteId)
    
    ' 检查缓存
    If IsObject(CE_NoteCostCache) Then
        If CE_NoteCostCache.Exists(key) Then
            CE_GetCachedNoteCost = CE_NoteCostCache(key)
            Exit Function
        End If
    End If
    
    totalCost = 0
    hasAccord = False
    
    ' 路径A: Accord配方
    If IsObject(CE_AccordRecipes) Then
        If CE_AccordRecipes.Exists(key) Then
            hasAccord = True
            accordData = CE_AccordRecipes(key)
            accordParts = Split(accordData, "|")
            accordId = accordParts(0)
            batchSize = CDbl(accordParts(1))
            If batchSize <= 0 Then batchSize = 100
            
            If CE_AccordMaterials.Exists(accordId) Then
                matList = CE_AccordMaterials(accordId)
                matItems = Split(matList, ",")
                For i = 0 To UBound(matItems)
                    If matItems(i) <> "" Then
                        matParts = Split(matItems(i), ":")
                        If UBound(matParts) >= 1 Then
                            matId = matParts(0)
                            matPct = CDbl(matParts(1))
                            matUnitCost = CE_GetCachedMaterialCost(matId)
                            totalCost = totalCost + (matPct / batchSize) * matUnitCost
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    ' 路径B: NoteIngredients
    If Not hasAccord And IsObject(CE_NoteIngredients) Then
        If CE_NoteIngredients.Exists(key) Then
            ingList = CE_NoteIngredients(key)
            ingItems = Split(ingList, ",")
            For j = 0 To UBound(ingItems)
                If ingItems(j) <> "" Then
                    ingParts = Split(ingItems(j), ":")
                    If UBound(ingParts) >= 1 Then
                        baseNoteId = ingParts(0)
                        ingPct = CDbl(ingParts(1))
                        baseCost = CE_GetCachedNoteCost(baseNoteId)
                        If baseCost <= 0 Then
                            ' V9: 优先从 BaseNotes 单价获取
                            If IsObject(CE_BaseNotePrices) Then
                                If CE_BaseNotePrices.Exists(baseNoteId) Then
                                    baseCost = CE_BaseNotePrices(baseNoteId)
                                End If
                            End If
                            ' 回退到 PriceAddition
                            If baseCost <= 0 And IsObject(CE_NotePriceAdditions) Then
                                If CE_NotePriceAdditions.Exists(baseNoteId) Then
                                    baseCost = CE_NotePriceAdditions(baseNoteId)
                                End If
                            End If
                        End If
                        totalCost = totalCost + (baseCost * ingPct / 100)
                    End If
                End If
            Next
        End If
    End If
    
    ' 兜底: PriceAddition
    If totalCost <= 0 And IsObject(CE_NotePriceAdditions) Then
        If CE_NotePriceAdditions.Exists(key) Then
            totalCost = CE_NotePriceAdditions(key)
        End If
    End If
    
    ' 存入缓存
    If IsObject(CE_NoteCostCache) Then CE_NoteCostCache.Add key, totalCost
    CE_GetCachedNoteCost = totalCost
End Function

' ============================================
' 缓存版：获取产品BOM成本（首次计算后缓存，零DB查询）
' ============================================
Function CE_GetCachedProductBOMCost(productId)
    Dim key, totalCost
    Dim ratioList, ratioItems, ratioData, ratioParts, noteId, notePct, noteCost
    Dim i
    
    key = CStr(productId)
    
    If IsObject(CE_ProductBOMCache) Then
        If CE_ProductBOMCache.Exists(key) Then
            CE_GetCachedProductBOMCost = CE_ProductBOMCache(key)
            Exit Function
        End If
    End If
    
    totalCost = 0
    
    ' 香调配比成本
    If IsObject(CE_ProductNoteRatios) Then
        If CE_ProductNoteRatios.Exists(key) Then
            ratioList = CE_ProductNoteRatios(key)
            ratioItems = Split(ratioList, ",")
            For i = 0 To UBound(ratioItems)
                If ratioItems(i) <> "" Then
                    ratioParts = Split(ratioItems(i), ":")
                    If UBound(ratioParts) >= 1 Then
                        noteId = ratioParts(0)
                        notePct = CDbl(ratioParts(1))
                        noteCost = CE_GetCachedNoteCost(noteId)
                        totalCost = totalCost + (noteCost * notePct / 100)
                    End If
                End If
            Next
        End If
    End If
    
    ' 瓶身成本
    If IsObject(CE_BottleAdditions) Then
        If CE_BottleAdditions.Exists(key) Then
            totalCost = totalCost + CE_BottleAdditions(key)
        End If
    End If
    
    If IsObject(CE_ProductBOMCache) Then CE_ProductBOMCache.Add key, totalCost
    CE_GetCachedProductBOMCost = totalCost
End Function

' ============================================
' 缓存版：获取产品单位总成本（首次计算后缓存，零DB查询）
' ============================================
Function CE_GetCachedProductUnitCost(productId)
    Dim key, totalCost
    Dim extraParts
    
    key = CStr(productId)
    
    If IsObject(CE_ProductUnitCache) Then
        If CE_ProductUnitCache.Exists(key) Then
            CE_GetCachedProductUnitCost = CE_ProductUnitCache(key)
            Exit Function
        End If
    End If
    
    totalCost = CE_GetCachedProductBOMCost(productId)
    
    ' 包装成本 + 其他成本
    If IsObject(CE_ProductExtraCosts) Then
        If CE_ProductExtraCosts.Exists(key) Then
            extraParts = Split(CE_ProductExtraCosts(key), "|")
            If UBound(extraParts) >= 1 Then
                totalCost = totalCost + CDbl(extraParts(0)) + CDbl(extraParts(1))
            End If
        End If
    End If
    
    If IsObject(CE_ProductUnitCache) Then CE_ProductUnitCache.Add key, totalCost
    CE_GetCachedProductUnitCost = totalCost
End Function

' 安全数值转换
' V19修复：MSOLEDBSQL(DataTypeCompatibility=0) 下 DECIMAL/NUMERIC 以 adNumeric 变体返回，
' VBScript 的 IsNumeric() 对该类型返回 False（TypeName 为空），原判断误将有效小数当成非数字返回0，
' 进而导致成本引擎把全部成本清零。改为直接 CDbl 转换并用错误兜底，不再依赖 IsNumeric。
Function CE_SafeNum(val)
    On Error Resume Next
    CE_SafeNum = 0
    If IsNull(val) Or IsEmpty(val) Then Exit Function
    CE_SafeNum = CDbl(val)
    If Err.Number <> 0 Then
        CE_SafeNum = 0
        Err.Clear
    End If
End Function

' V10: 获取指定物料的加权平均成本（ItemCode → WeightedUnitCost, 覆盖5个品类）
Function CE_GetBatchWeightedCost(itemCode)
    Dim rs, cost
    cost = 0
    On Error Resume Next
    ' 按优先级依次查询5个品类的库存表
    Set rs = conn.Execute("SELECT TOP 1 ISNULL(WeightedUnitCost, UnitPrice) FROM RawMaterialInventory WHERE ItemCode='" & CE_SafeSQL(itemCode) & "'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then cost = CE_SafeNum(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    If cost <= 0 Then
        Set rs = conn.Execute("SELECT TOP 1 ISNULL(WeightedUnitCost, UnitPrice) FROM PackagingInventory WHERE ItemCode='" & CE_SafeSQL(itemCode) & "' AND IsActive=1")
        If Not rs Is Nothing Then
            If Not rs.EOF Then cost = CE_SafeNum(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    If cost <= 0 Then
        Set rs = conn.Execute("SELECT TOP 1 ISNULL(WeightedUnitCost, UnitPrice) FROM BottleStyles WHERE ItemCode='" & CE_SafeSQL(itemCode) & "'")
        If Not rs Is Nothing Then
            If Not rs.EOF Then cost = CE_SafeNum(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    If cost <= 0 Then
        Set rs = conn.Execute("SELECT TOP 1 ISNULL(WeightedUnitCost, UnitPrice) FROM PrintingInventory WHERE ItemCode='" & CE_SafeSQL(itemCode) & "' AND IsActive=1")
        If Not rs Is Nothing Then
            If Not rs.EOF Then cost = CE_SafeNum(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    If cost <= 0 Then
        Set rs = conn.Execute("SELECT TOP 1 ISNULL(WeightedUnitCost, UnitPrice) FROM SprayHeadInventory WHERE ItemCode='" & CE_SafeSQL(itemCode) & "' AND IsActive=1")
        If Not rs Is Nothing Then
            If Not rs.EOF Then cost = CE_SafeNum(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    CE_GetBatchWeightedCost = cost
End Function

' 安全SQL
Function CE_SafeSQL(str)
    If IsNull(str) Or str = "" Then CE_SafeSQL = "" Else CE_SafeSQL = Replace(str, "'", "''")
End Function

' ============================================
' Level 1: 计算原材料加权平均成本
' V10: 优先使用 WeightedUnitCost，回退到 SupplierPrices 最新报价
' ============================================
Function CE_CalculateMaterialCost(materialId)
    Dim rs, cost, itemCode
    cost = 0
    On Error Resume Next
    
    ' V10: 先检查是否有加权平均成本
    Set rs = conn.Execute("SELECT ItemCode, ISNULL(WeightedUnitCost,0) as WUC FROM RawMaterialInventory WHERE MaterialID=" & materialId)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            itemCode = CStr(rs("ItemCode") & "")
            Dim wucVal : wucVal = CE_SafeNum(rs("WUC"))
            If wucVal > 0 Then
                cost = wucVal
                rs.Close : Set rs = Nothing
                CE_CalculateMaterialCost = cost
                Exit Function
            End If
        End If
        rs.Close
    End If
    Set rs = Nothing

    ' 回退到最新采购价
    If cost <= 0 Then
        Set rs = conn.Execute("SELECT TOP 1 UnitPrice FROM SupplierPrices WHERE ItemCode IN (SELECT ItemCode FROM RawMaterialInventory WHERE MaterialID=" & materialId & ") AND IsActive=1 ORDER BY CreatedAt DESC")
        If Not rs Is Nothing Then
            If Not rs.EOF Then cost = CE_SafeNum(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    ' 如果无采购价，使用库存中的最后采购价
    If cost <= 0 Then
        Set rs = conn.Execute("SELECT UnitPrice FROM RawMaterialInventory WHERE MaterialID=" & materialId)
        If Not rs Is Nothing Then
            If Not rs.EOF Then cost = CE_SafeNum(rs(0))
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    CE_CalculateMaterialCost = cost
End Function

' ============================================
' Level 2: 计算香调(Note)的单位成本
' 路径A：通过 RecipeAccords → RecipeAccordMaterials 计算
' 路径B：通过 NoteIngredients → BaseNotes → 成分聚合计算
' ============================================
Function CE_CalculateNoteCost(noteId)
    Dim totalCost, hasAccord
    totalCost = 0
    hasAccord = False
    
    On Error Resume Next
    
    ' === 路径A: 通过 Accord 生产配方计算 ===
    Dim rsAccord
    Set rsAccord = conn.Execute("SELECT TOP 1 AccordRecipeID, BatchSize FROM RecipeAccords WHERE NoteID=" & noteId & " AND Status='Published' ORDER BY PublishedAt DESC")
    If Not rsAccord Is Nothing Then
        If Not rsAccord.EOF Then
            hasAccord = True
            Dim accordId, batchSize
            accordId = CE_SafeNum(rsAccord(0))
            batchSize = CE_SafeNum(rsAccord(1))
            If batchSize <= 0 Then batchSize = 100
            rsAccord.Close
            Set rsAccord = Nothing
            
            ' 计算该Accord中所有原材料的成本
            Dim rsMaterials, matId, matPct, matUnitCost, matCost
            Set rsMaterials = conn.Execute("SELECT MaterialID, Percentage, PlannedQty FROM RecipeAccordMaterials WHERE AccordRecipeID=" & accordId)
            If Not rsMaterials Is Nothing Then
                Do While Not rsMaterials.EOF
                    matId = CE_SafeNum(rsMaterials("MaterialID"))
                    matPct = CE_SafeNum(rsMaterials("Percentage"))
                    matUnitCost = CE_CalculateMaterialCost(matId)
                    ' 成本 = (配比/批量) × 材料单价
                    matCost = (matPct / batchSize) * matUnitCost
                    totalCost = totalCost + matCost
                    rsMaterials.MoveNext
                Loop
                rsMaterials.Close
            End If
            Set rsMaterials = Nothing
        Else
            rsAccord.Close
        End If
    End If
    Set rsAccord = Nothing
    
    ' === 路径B: 通过 BaseNote 成分聚合计算 ===
    If Not hasAccord Then
        Dim rsIngredients
        Set rsIngredients = conn.Execute("SELECT ni.BaseNoteID, ni.Percentage, bn.BaseNoteName FROM (NoteIngredients ni LEFT JOIN BaseNotes bn ON ni.BaseNoteID=bn.BaseNoteID) WHERE ni.NoteID=" & noteId)
        If Not rsIngredients Is Nothing Then
            If Not rsIngredients.EOF Then
                ' 通过NoteIngredients关联获取成分成本
                Do While Not rsIngredients.EOF
                    Dim baseNoteId, ingPct, baseNoteCost
                    baseNoteId = CE_SafeNum(rsIngredients("BaseNoteID"))
                    ingPct = CE_SafeNum(rsIngredients("Percentage"))
                    
                    ' 计算基香成本：检查是否有对应的Accord
                    baseNoteCost = CE_CalculateNoteCost(baseNoteId)
                    If baseNoteCost <= 0 Then
                        ' V9: 优先从 BaseNotes 单价获取
                        Dim rsBN2
                        Set rsBN2 = conn.Execute("SELECT UnitPrice FROM BaseNotes WHERE BaseNoteID=" & baseNoteId & " AND IsActive <> 0")
                        If Not rsBN2 Is Nothing Then
                            If Not rsBN2.EOF Then baseNoteCost = CE_SafeNum(rsBN2(0))
                            rsBN2.Close
                        End If
                        Set rsBN2 = Nothing
                        ' 如果基香也没有Accord，尝试从FragranceNotes.PriceAddition获取
                        If baseNoteCost <= 0 Then
                            Dim rsBN
                            Set rsBN = conn.Execute("SELECT PriceAddition FROM FragranceNotes WHERE NoteID=" & baseNoteId)
                            If Not rsBN Is Nothing Then
                                If Not rsBN.EOF Then baseNoteCost = CE_SafeNum(rsBN(0))
                                rsBN.Close
                            End If
                            Set rsBN = Nothing
                        End If
                    End If
                    
                    totalCost = totalCost + (baseNoteCost * ingPct / 100)
                    rsIngredients.MoveNext
                Loop
            End If
            rsIngredients.Close
        End If
        Set rsIngredients = Nothing
    End If
    
    ' 如果上述方法都无数据，尝试使用 FragranceNotes.PriceAddition
    If totalCost <= 0 Then
        Dim rsFN
        Set rsFN = conn.Execute("SELECT PriceAddition FROM FragranceNotes WHERE NoteID=" & noteId)
        If Not rsFN Is Nothing Then
            If Not rsFN.EOF Then totalCost = CE_SafeNum(rsFN(0))
            rsFN.Close
        End If
        Set rsFN = Nothing
    End If
    
    CE_CalculateNoteCost = totalCost
End Function

' ============================================
' Level 3: 计算产品的BOM成本
' 来源：ProductNoteRatios × Note成本 + 瓶身成本 + 包装成本
' ============================================
Function CE_CalculateProductBOMCost(productId)
    Dim totalCost, rsPN, noteId, notePct, noteCost
    
    totalCost = 0
    On Error Resume Next
    
    ' 1. 香调配比成本
    Set rsPN = conn.Execute("SELECT NoteID, Percentage FROM ProductNoteRatios WHERE ProductID=" & productId)
    If Not rsPN Is Nothing Then
        Do While Not rsPN.EOF
            noteId = CE_SafeNum(rsPN("NoteID"))
            notePct = CE_SafeNum(rsPN("Percentage"))
            noteCost = CE_CalculateNoteCost(noteId)
            totalCost = totalCost + (noteCost * notePct / 100)
            rsPN.MoveNext
        Loop
        rsPN.Close
    End If
    Set rsPN = Nothing
    
    ' 2. 瓶身成本（从 BottleStyles.PriceAddition 获取附加成本）
    Dim rsBottle, bottleAdd
    Set rsBottle = conn.Execute("SELECT TOP 1 bs.PriceAddition FROM ProductBottleStyles pbs LEFT JOIN BottleStyles bs ON pbs.BottleID=bs.BottleID WHERE pbs.ProductID=" & productId & " AND bs.IsActive=1")
    If Not rsBottle Is Nothing Then
        If Not rsBottle.EOF Then
            bottleAdd = CE_SafeNum(rsBottle(0))
            totalCost = totalCost + bottleAdd
        End If
        rsBottle.Close
    End If
    Set rsBottle = Nothing
    
    CE_CalculateProductBOMCost = totalCost
End Function

' ============================================
' Level 4: 计算产品完整单位成本 (BOM + 包装 + 人工分摊)
' ============================================
Function CE_CalculateProductUnitCost(productId)
    Dim totalCost, packagingCost, laborOverhead
    
    ' 品牌定香(Fixed)产品：优先使用采购加权平均成本
    Dim productType : productType = ""
    On Error Resume Next
    Dim rsPType : Set rsPType = conn.Execute("SELECT ProductType FROM Products WHERE ProductID=" & productId)
    If Not rsPType Is Nothing And Not rsPType.EOF Then
        productType = CStr(rsPType("ProductType") & "")
        rsPType.Close
    End If
    Set rsPType = Nothing
    On Error GoTo 0
    
    If productType = "standard" Then
        Dim fbCost : fbCost = CE_GetCachedFixedBrandCost(productId)
        If fbCost > 0 Then
            CE_CalculateProductUnitCost = fbCost
            Exit Function
        End If
        ' 回退：使用 Products.UnitCost（可能已由采购模块同步）
        Dim existingUnitCost : existingUnitCost = 0
        On Error Resume Next
        Dim rsUC : Set rsUC = conn.Execute("SELECT ISNULL(UnitCost,0) FROM Products WHERE ProductID=" & productId)
        If Not rsUC Is Nothing And Not rsUC.EOF Then
            existingUnitCost = CE_SafeNum(rsUC(0))
            rsUC.Close
        End If
        Set rsUC = Nothing
        On Error GoTo 0
        If existingUnitCost > 0 Then
            CE_CalculateProductUnitCost = existingUnitCost
            Exit Function
        End If
    End If
    
    totalCost = CE_CalculateProductBOMCost(productId)
    
    ' 包装成本：从 ProductCosts 获取包装项
    On Error Resume Next
    Dim rsPkg
    Set rsPkg = conn.Execute("SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID=" & productId & " AND CostType='Packaging'")
    If Not rsPkg Is Nothing Then
        If Not rsPkg.EOF Then packagingCost = CE_SafeNum(rsPkg(0))
        rsPkg.Close
    End If
    Set rsPkg = Nothing
    
    ' 人工与管理费用分摊
    Set rsPkg = conn.Execute("SELECT SUM(TotalCost) FROM ProductCosts WHERE ProductID=" & productId & " AND CostType='Other'")
    If Not rsPkg Is Nothing Then
        If Not rsPkg.EOF Then laborOverhead = CE_SafeNum(rsPkg(0))
        rsPkg.Close
    End If
    Set rsPkg = Nothing
    
    CE_CalculateProductUnitCost = totalCost + packagingCost + laborOverhead
End Function

' ============================================
' Level 5: 批量更新所有产品的成本
' ============================================
Sub CE_UpdateAllProductCosts()
    Dim rsProducts, pid, bomCost, unitCost, updateSQL, count
    
    count = 0
    On Error Resume Next
    
    ' V19修复：预加载成本缓存（顺序查询、互不嵌套），随后用O(1)缓存计算，
    ' 从根本上规避 forward-only 游标下嵌套查询导致的"连接繁忙"清零问题。
    Call CE_PreloadAllCostData()
    
    ' 先将所有 ProductID 缓冲到数组，再关闭 rsProducts，
    ' 避免在遍历 forward-only 游标时嵌套 conn.Execute 触发"连接繁忙"（无MARS），
    ' 该问题会导致内层计算查询静默失败并返回0，进而把全部产品成本清零。
    Dim idBuf, idArr, bufI
    idBuf = ""
    Set rsProducts = conn.Execute("SELECT ProductID FROM Products WHERE IsActive=1")
    If Not rsProducts Is Nothing Then
        Do While Not rsProducts.EOF
            idBuf = idBuf & CStr(rsProducts("ProductID")) & ","
            rsProducts.MoveNext
        Loop
        rsProducts.Close
    End If
    Set rsProducts = Nothing
    
    If Len(idBuf) > 0 Then
        idArr = Split(Left(idBuf, Len(idBuf)-1), ",")
        For bufI = 0 To UBound(idArr)
            pid = CE_SafeNum(idArr(bufI))
            If pid > 0 Then
                bomCost = CE_GetCachedProductBOMCost(pid)
                unitCost = CE_GetCachedProductUnitCost(pid)
                If unitCost <= 0 Then unitCost = bomCost
                
                updateSQL = "UPDATE Products SET BOMCost=" & bomCost & ", UnitCost=" & unitCost & " WHERE ProductID=" & pid
                conn.Execute updateSQL
                If Err.Number = 0 Then count = count + 1
                Err.Clear
            End If
        Next
    End If
    
    Session("CE_LastUpdateCount") = count
    Session("CE_LastUpdateTime") = Now()
End Sub

' ============================================
' Level 6: 更新指定产品的成本
' ============================================
Sub CE_UpdateProductCost(productId)
    Dim bomCost, unitCost
    On Error Resume Next
    
    bomCost = CE_CalculateProductBOMCost(productId)
    unitCost = CE_CalculateProductUnitCost(productId)
    
    conn.Execute "UPDATE Products SET BOMCost=" & bomCost & ", UnitCost=" & unitCost & " WHERE ProductID=" & productId
End Sub

' ============================================
' Level 7: 更新订单的成本和利润
' 计算：CostAmount = 各商品数量 × 单位成本
'       ProfitAmount = TotalAmount - CostAmount - ShippingFee
' ============================================
Sub CE_UpdateOrderCosts(orderId)
    Dim rsDetails, detailId, productId, qty, unitCost, orderCost
    orderCost = 0
    
    On Error Resume Next
    
    ' V19修复：用单条联表查询直接从 Products.UnitCost 汇总订单成本，
    ' 避免逐行嵌套 CE_CalculateProductUnitCost 及"记录集打开时执行UPDATE"导致的连接繁忙
    Dim rsC
    Set rsC = conn.Execute("SELECT ISNULL(SUM(od.Quantity * ISNULL(p.UnitCost,0)),0) AS Cost " & _
        "FROM OrderDetails od LEFT JOIN Products p ON od.ProductID=p.ProductID WHERE od.OrderID=" & orderId)
    If Not rsC Is Nothing Then
        If Not rsC.EOF Then orderCost = CE_SafeNum(rsC(0))
        rsC.Close
    End If
    Set rsC = Nothing
    
    ' 更新订单的成本和利润
    Dim totalAmount, shippingFee, profitAmount, gotOrder
    Dim rsOrder
    gotOrder = False
    Set rsOrder = conn.Execute("SELECT TotalAmount, ShippingFee FROM Orders WHERE OrderID=" & orderId)
    If Not rsOrder Is Nothing Then
        If Not rsOrder.EOF Then
            totalAmount = CE_SafeNum(rsOrder("TotalAmount"))
            shippingFee = CE_SafeNum(rsOrder("ShippingFee"))
            gotOrder = True
        End If
        rsOrder.Close
    End If
    Set rsOrder = Nothing
    
    If gotOrder Then
        profitAmount = totalAmount - orderCost - shippingFee
        If profitAmount < 0 Then profitAmount = 0
        conn.Execute "UPDATE Orders SET CostAmount=" & orderCost & ", ProfitAmount=" & profitAmount & " WHERE OrderID=" & orderId
    End If
    
    ' V10: 记录订单成本分摊明细
    CE_RecordOrderCostAllocation orderId
End Sub

' ============================================
' Level 8: 批量更新所有已支付订单的成本和利润
' ============================================
Sub CE_UpdateAllOrderCosts()
    Dim rsOrders, oid
    
    On Error Resume Next
    ' V19修复：先缓冲所有 OrderID，再关闭游标，避免遍历时嵌套查询导致连接繁忙
    Dim oidBuf, oidArr, oBufI
    oidBuf = ""
    Set rsOrders = conn.Execute("SELECT OrderID FROM Orders WHERE Status NOT IN ('Pending','Cancelled')")
    If Not rsOrders Is Nothing Then
        Do While Not rsOrders.EOF
            oidBuf = oidBuf & CStr(rsOrders("OrderID")) & ","
            rsOrders.MoveNext
        Loop
        rsOrders.Close
    End If
    Set rsOrders = Nothing
    
    If Len(oidBuf) > 0 Then
        oidArr = Split(Left(oidBuf, Len(oidBuf)-1), ",")
        For oBufI = 0 To UBound(oidArr)
            oid = CE_SafeNum(oidArr(oBufI))
            If oid > 0 Then CE_UpdateOrderCosts oid
        Next
    End If
    
    Session("CE_LastOrderUpdateTime") = Now()
End Sub

' ============================================
' Level 9: 获取成本传导状态摘要
' ============================================
Function CE_GetCostSummary()
    Dim summary, rs
    
    summary = ""
    On Error Resume Next
    
    ' 已更新成本的产品数
    Dim updatedProducts, totalProducts, updatedOrders, totalOrders
    Set rs = conn.Execute("SELECT COUNT(*) FROM Products WHERE IsActive=1 AND UnitCost > 0")
    If Not rs Is Nothing Then updatedProducts = CE_SafeNum(rs(0)) : rs.Close
    Set rs = Nothing
    
    Set rs = conn.Execute("SELECT COUNT(*) FROM Products WHERE IsActive=1")
    If Not rs Is Nothing Then totalProducts = CE_SafeNum(rs(0)) : rs.Close
    Set rs = Nothing
    
    Set rs = conn.Execute("SELECT COUNT(*) FROM Orders WHERE CostAmount > 0")
    If Not rs Is Nothing Then updatedOrders = CE_SafeNum(rs(0)) : rs.Close
    Set rs = Nothing
    
    Set rs = conn.Execute("SELECT COUNT(*) FROM Orders WHERE Status NOT IN ('Pending','Cancelled')")
    If Not rs Is Nothing Then totalOrders = CE_SafeNum(rs(0)) : rs.Close
    Set rs = Nothing
    
    summary = "{""totalProducts"":" & totalProducts & ",""updatedProducts"":" & updatedProducts _
            & ",""totalOrders"":" & totalOrders & ",""updatedOrders"":" & updatedOrders _
            & ",""lastUpdate"":""" & Session("CE_LastUpdateTime") & """}"
    
    CE_GetCostSummary = summary
End Function

' ============================================
' V10: 记录订单成本分摊到 OrderCostAllocation 表
' 将每个订单明细的产品成本记录为独立的成本分摊条目
' ============================================
Sub CE_RecordOrderCostAllocation(orderId)
    Dim rsOrder, orderNo, rsDetails
    Dim detailId, productId, productName, qty, unitCost, totalCost
    Dim allocSql
    
    On Error Resume Next
    Err.Clear
    
    ' 获取订单号
    Set rsOrder = conn.Execute("SELECT OrderNo FROM Orders WHERE OrderID=" & orderId)
    If Not rsOrder Is Nothing And Not rsOrder.EOF Then
        orderNo = CStr(rsOrder("OrderNo"))
        rsOrder.Close
    End If
    Set rsOrder = Nothing
    If orderNo = "" Then Exit Sub
    
    ' V19修复：先清除该订单旧的分摊记录，避免重复执行"自动更新订单利润"时累积重复行
    conn.Execute "DELETE FROM OrderCostAllocation WHERE OrderID=" & orderId
    Err.Clear
    
    ' 获取订单明细及产品成本
    ' V19修复：先缓冲明细（Chr(31)字段分隔/Chr(30)行分隔），再关闭游标，
    ' 避免遍历时嵌套 INSERT 与 CE_CalculateProductUnitCost 查询导致连接繁忙
    Dim recBuf, recArr, rBufI, recFields
    recBuf = ""
    Set rsDetails = conn.Execute("SELECT od.ProductID, p.ProductName, od.Quantity, ISNULL(p.UnitCost,0) AS UnitCost " & _
        "FROM OrderDetails od LEFT JOIN Products p ON od.ProductID = p.ProductID " & _
        "WHERE od.OrderID=" & orderId)
    If Not rsDetails Is Nothing Then
        Do While Not rsDetails.EOF
            recBuf = recBuf & CStr(CE_SafeNum(rsDetails("ProductID"))) & Chr(31) & _
                     Replace(CStr(rsDetails("ProductName") & ""), Chr(30), "") & Chr(31) & _
                     CStr(CE_SafeNum(rsDetails("Quantity"))) & Chr(31) & _
                     CStr(CE_SafeNum(rsDetails("UnitCost"))) & Chr(30)
            rsDetails.MoveNext
        Loop
        rsDetails.Close
    End If
    Set rsDetails = Nothing
    
    If Len(recBuf) > 0 Then
        recArr = Split(Left(recBuf, Len(recBuf)-1), Chr(30))
        For rBufI = 0 To UBound(recArr)
            recFields = Split(recArr(rBufI), Chr(31))
            If UBound(recFields) >= 3 Then
                productId = CE_SafeNum(recFields(0))
                productName = recFields(1)
                qty = CE_SafeNum(recFields(2))
                unitCost = CE_SafeNum(recFields(3))
                
                ' 如果产品成本未计算，重新计算
                If unitCost <= 0 Then
                    unitCost = CE_CalculateProductUnitCost(productId)
                End If
                
                totalCost = unitCost * qty
                
                If totalCost > 0 Then
                    allocSql = "INSERT INTO OrderCostAllocation (OrderID, OrderNo, CostType, ItemCode, ItemName, UnitCost, Quantity, TotalCost, AllocatedAt, CreatedAt) VALUES (" & _
                        orderId & ", '" & orderNo & "', 'Product', '" & CStr(productId) & "', '" & Replace(productName, "'", "''") & "', " & _
                        unitCost & ", " & qty & ", " & totalCost & ", GETDATE(), GETDATE())"
                    conn.Execute allocSql
                    If Err.Number <> 0 Then Err.Clear
                End If
            End If
        Next
    End If
End Sub
%>
