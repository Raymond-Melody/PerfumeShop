<!-- ============================================
     V14.6 产品设置 - 产品管理标签页模板
     从 product_settings.asp 提取
     包含：统计区域、筛选栏、产品卡片网格
     ============================================ -->
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
                    Select Case LCase(pType)
                        Case "standard": pTypeClass = "status-fixed": pBadgeClass = "fixed"
                        Case "custom": pTypeClass = "status-custom": pBadgeClass = "custom"
                        Case "kol": pTypeClass = "status-kol": pBadgeClass = "kol"
                        Case Else: pTypeClass = "": pBadgeClass = ""
                    End Select
                    
                    ' 计算实际显示价格
                    Dim pDisplayPrice
                    pDisplayPrice = SafeNum(rsProducts("BasePrice"))
                    If LCase(pType) = "standard" Then
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
                        <% Select Case LCase(pType)
                            Case "standard" %><i class="fas fa-box"></i> 品牌定香
                        <%  Case "custom" %><i class="fas fa-paint-brush"></i> 用户定制
                        <%  Case "kol" %><i class="fas fa-star"></i> KOL推荐
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
        
