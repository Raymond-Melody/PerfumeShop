<!-- ============================================
     V14.6 产品设置 - 模态框模板
     包含：产品新增/编辑模态框、类型编辑模态框
     ============================================ -->

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
