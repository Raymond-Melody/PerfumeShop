<!-- ============================================
     V14.6 采购订单 - 新建/编辑表单模板
     包含：预选择逻辑、供应商/分类/类型选择、
           明细行表格、基香快速选择
     ============================================ -->
        <%
        ' ========== 新建订单预选类型 ==========
        Dim preselectCategory, preselectOrderType
        preselectCategory = ""
        preselectOrderType = ""
        If Request.QueryString("new") = "1" Then
            Dim qsOT : qsOT = Trim(Request.QueryString("order_type"))
            If qsOT <> "" Then
                preselectOrderType = qsOT
                Select Case qsOT
                    Case "RawMaterial" : preselectCategory = "RAW"
                    Case "Packaging" : preselectCategory = "PACK"
                    Case "Bottle" : preselectCategory = "BOTTLE"
                    Case "Printing" : preselectCategory = "PRINTING"
                    Case "SprayHead" : preselectCategory = "SPRAYHEAD"
                End Select
            End If
        End If

        ' ========== 预计算采购分类选中状态 ==========
        Dim catSelRAW, catSelBASE, catSelPACK, catSelMARKET, catSelBOTTLE, catSelPRINTING, catSelSPRAYHEAD
        catSelRAW = "" : catSelBASE = "" : catSelPACK = "" : catSelMARKET = "" : catSelBOTTLE = "" : catSelPRINTING = "" : catSelSPRAYHEAD = ""
        If editMode Then
            Dim edCatVal : edCatVal = CStr(editOrderData("CategoryCode"))
            If edCatVal = "RAW" Then catSelRAW = "selected"
            If edCatVal = "BASE" Then catSelBASE = "selected"
            If edCatVal = "PACK" Then catSelPACK = "selected"
            If edCatVal = "MARKET" Then catSelMARKET = "selected"
            If edCatVal = "BOTTLE" Then catSelBOTTLE = "selected"
            If edCatVal = "PRINTING" Then catSelPRINTING = "selected"
            If edCatVal = "SPRAYHEAD" Then catSelSPRAYHEAD = "selected"
        Else
            If preselectCategory = "RAW" Then catSelRAW = "selected"
            If preselectCategory = "BASE" Then catSelBASE = "selected"
            If preselectCategory = "PACK" Then catSelPACK = "selected"
            If preselectCategory = "MARKET" Then catSelMARKET = "selected"
            If preselectCategory = "BOTTLE" Then catSelBOTTLE = "selected"
            If preselectCategory = "PRINTING" Then catSelPRINTING = "selected"
            If preselectCategory = "SPRAYHEAD" Then catSelSPRAYHEAD = "selected"
        End If

        ' ========== 预计算采购类型选中状态 ==========
        Dim otSelRAW, otSelPACK, otSelBOTTLE, otSelPRINT, otSelSPRAY
        otSelRAW = "" : otSelPACK = "" : otSelBOTTLE = "" : otSelPRINT = "" : otSelSPRAY = ""
        If editMode Then
            Dim edOTVal : edOTVal = CStr(editOrderData("OrderType"))
            If edOTVal = "RawMaterial" Then otSelRAW = "selected"
            If edOTVal = "Packaging" Then otSelPACK = "selected"
            If edOTVal = "Bottle" Then otSelBOTTLE = "selected"
            If edOTVal = "Printing" Then otSelPRINT = "selected"
            If edOTVal = "SprayHead" Then otSelSPRAY = "selected"
        Else
            If preselectOrderType = "RawMaterial" Or preselectOrderType = "" Then otSelRAW = "selected"
            If preselectOrderType = "Packaging" Then otSelPACK = "selected"
            If preselectOrderType = "Bottle" Then otSelBOTTLE = "selected"
            If preselectOrderType = "Printing" Then otSelPRINT = "selected"
            If preselectOrderType = "SprayHead" Then otSelSPRAY = "selected"
        End If

        ' ========== 基香按钮可见性 ==========
        Dim showBaseNoteSection
        showBaseNoteSection = False
        If editMode Then
            If CStr(editOrderData("CategoryCode")) = "BASE" Or CStr(editOrderData("OrderType")) = "RawMaterial" Then
                showBaseNoteSection = True
            End If
        Else
            If preselectCategory = "BASE" Or preselectOrderType = "RawMaterial" Or preselectOrderType = "" Then
                showBaseNoteSection = True
            End If
        End If
        %>

        <% If editMode Or Request.QueryString("new") = "1" Then %>
        ' ========== 创建/编辑表单 ==========
        <div class="form-section">
            <h3>
                <% If editMode Then %>
                <i class="fas fa-edit"></i> 编辑采购订单
                <% Else %>
                <i class="fas fa-plus-circle"></i> 新建采购订单
                <% End If %>
            </h3>

            <form method="post" id="purchaseForm">
                <input type="hidden" name="action" value="<%= IIf(editMode, "update", "create") %>">
                <% If editMode Then %>
                <input type="hidden" name="purchase_id" value="<%= editOrderID %>">
                <% End If %>

                <div class="form-row">
                    <div class="form-group">
                        <label>采购单号</label>
                        <% If editMode Then %>
                        <input type="text" value="<%= Server.HTMLEncode(CStr(editOrderData("PurchaseNo"))) %>" readonly style="background:#1a1a2e;">
                        <% Else %>
                        <input type="text" value="<%= GeneratePurchaseNo() %>（自动生成）" readonly style="background:#1a1a2e;">
                        <% End If %>
                    </div>
                    <div class="form-group">
                        <label>供应商 <span style="color:#F44336;">*</span></label>
                        <select name="supplier_id" required>
                            <option value="">请选择供应商</option>
                            <%
                            If Not rsSuppliers Is Nothing Then
                                Do While Not rsSuppliers.EOF
                                    Dim selected
                                    selected = ""
                                    If editMode Then
                                        If SafeNum(editOrderData("SupplierID")) = SafeNum(rsSuppliers("SupplierID")) Then
                                            selected = "selected"
                                        End If
                                    End If
                            %>
                            <option value="<%= rsSuppliers("SupplierID") %>" <%= selected %>><%= Server.HTMLEncode(CStr(IIf(IsNull(rsSuppliers("SupplierName")), "", rsSuppliers("SupplierName")))) %></option>
                            <%
                                    rsSuppliers.MoveNext
                                Loop
                                rsSuppliers.Close
                                Set rsSuppliers = Nothing
                            End If
                            %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>采购分类 <span style="color:#F44336;">*</span></label>
                        <select name="category_code" required onchange="updateOrderTypeByCategory()">
                            <option value="">请选择分类</option>
                            <option value="RAW" <%= catSelRAW %>>原材料</option>
                            <option value="BASE" <%= catSelBASE %>>基香原料</option>
                            <option value="PACK" <%= catSelPACK %>>包装材料</option>
                            <option value="MARKET" <%= catSelMARKET %>>营销物料</option>
                            <option value="BOTTLE" <%= catSelBOTTLE %>>瓶子包装</option>
                            <option value="PRINTING" <%= catSelPRINTING %>>印刷品</option>
                            <option value="SPRAYHEAD" <%= catSelSPRAYHEAD %>>喷头配件</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>采购类型 <span style="color:#F44336;">*</span></label>
                        <select name="order_type" required onchange="updateCategoryByOrderType()">
                            <option value="">请选择类型</option>
                            <option value="RawMaterial" <%= otSelRAW %>>原料采购</option>
                            <option value="Packaging" <%= otSelPACK %>>包装物采购</option>
                            <option value="Bottle" <%= otSelBOTTLE %>>瓶子采购</option>
                            <option value="Printing" <%= otSelPRINT %>>印刷品采购</option>
                            <option value="SprayHead" <%= otSelSPRAY %>>喷头采购</option>
                        </select>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>期望交期</label>
                        <input type="date" name="expected_date"
                            <% If editMode Then %>
                            <% If Not IsNull(editOrderData("ExpectedDate")) And IsDate(editOrderData("ExpectedDate")) Then %>
                            value="<% If IsDate(editOrderData("ExpectedDate")) Then Response.Write FormatDateTime(editOrderData("ExpectedDate"), 2) End If %>"
                            <% End If %>
                            <% End If %>>
                    </div>
                    <div class="form-group" style="flex:2;">
                        <label>备注</label>
                        <input type="text" name="remarks" maxlength="500" placeholder="可选填"
                            <% If editMode Then %>
                            <% If Not IsNull(editOrderData("Remarks")) Then %>
                            value="<%= Server.HTMLEncode(CStr(editOrderData("Remarks"))) %>"
                            <% End If %>
                            <% End If %>>
                    </div>
                </div>

                <h4 style="margin:25px 0 15px 0;color:#fff;"><i class="fas fa-list"></i> 采购明细</h4>
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
                    <div id="baseNoteSection" style="display:<% If showBaseNoteSection Then Response.Write "flex" Else Response.Write "none" End If %>;align-items:center;gap:10px;">
                        <button type="button" class="btn-select-base" onclick="openBaseNoteModal()">
                            <i class="fas fa-flask"></i> 选择基香原料
                        </button>
                        <span style="font-size:11px;color:#888;">快速选择系统已录入的基香，自动填充物料信息（编码BN-XXX、单价等）</span>
                    </div>
                    <button type="button" class="btn-select-base" onclick="openHistoryModal()" style="margin-left:auto;">
                        <i class="fas fa-history"></i> 历史产品快速选择
                    </button>
                </div>
                <table class="details-table" id="detailsTable">
                    <thead>
                        <tr>
                            <th style="width:25%;">物料名称 <span style="color:#F44336;">*</span></th>
                            <th style="width:12%;">物料编码</th>
                            <th style="width:15%;">规格</th>
                            <th style="width:10%;">单位</th>
                            <th style="width:10%;">数量</th>
                            <th style="width:12%;">单价</th>
                            <th style="width:12%;">小计</th>
                            <th style="width:4%;"></th>
                        </tr>
                    </thead>
                    <tbody id="detailsBody">
                        <%
                        Dim rowCount
                        rowCount = 0

                        If editMode Then
                            If Not rsEditDetails Is Nothing Then
                                Do While Not rsEditDetails.EOF
                                    rowCount = rowCount + 1
                        %>
                        <tr>
                            <td><input type="text" name="item_name_<%= rowCount %>" value="<%= Server.HTMLEncode(CStr(rsEditDetails("ItemName"))) %>" required></td>
                            <td><input type="text" name="item_code_<%= rowCount %>" value="<% If Not IsNull(rsEditDetails("ItemCode")) Then Response.Write Server.HTMLEncode(CStr(rsEditDetails("ItemCode"))) %>"></td>
                            <td><input type="text" name="spec_<%= rowCount %>" value="<% If Not IsNull(rsEditDetails("Specification")) Then Response.Write Server.HTMLEncode(CStr(rsEditDetails("Specification"))) %>"></td>
                            <td><input type="text" name="unit_<%= rowCount %>" value="<% If Not IsNull(rsEditDetails("Unit")) Then Response.Write Server.HTMLEncode(CStr(rsEditDetails("Unit"))) %>"></td>
                            <td><input type="number" name="qty_<%= rowCount %>" class="num-input qty" value="<%= SafeNum(rsEditDetails("Quantity")) %>" min="0" step="0.01" onchange="calculateRow(this)"></td>
                            <td><input type="number" name="price_<%= rowCount %>" class="num-input price" value="<%= SafeNum(rsEditDetails("UnitPrice")) %>" min="0" step="0.01" onchange="calculateRow(this)"></td>
                            <td class="row-total">¥<%= FormatNumber(SafeNum(rsEditDetails("TotalPrice")), 2) %></td>
                            <td><button type="button" class="btn btn-danger btn-sm" onclick="removeRow(this)"><i class="fas fa-trash"></i></button></td>
                        </tr>
                        <%
                                    rsEditDetails.MoveNext
                                Loop
                                rsEditDetails.Close
                                Set rsEditDetails = Nothing
                            End If
                        End If
                        %>
                    </tbody>
                    <tfoot>
                        <tr>
                            <td colspan="8" style="text-align:center;">
                                <button type="button" class="btn btn-secondary" onclick="addRow()">
                                    <i class="fas fa-plus"></i> 添加明细行
                                </button>
                            </td>
                        </tr>
                    </tfoot>
                </table>

                <input type="hidden" name="item_count" id="itemCount" value="<%= rowCount %>">

                <div style="margin-top:25px;display:flex;gap:10px;justify-content:flex-end;">
                    <a href="purchase_orders.asp" class="btn btn-secondary">取消</a>
                    <button type="submit" class="btn btn-primary">
                        <i class="fas fa-save"></i> <%= IIf(editMode, "保存修改", "创建订单") %>
                    </button>
                </div>
            </form>
        </div>
        <%
        If editMode Then
            editOrderData.Close
            Set editOrderData = Nothing
        End If
        End If
        %>