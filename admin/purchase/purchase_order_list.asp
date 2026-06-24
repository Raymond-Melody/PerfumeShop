<!-- ============================================
     V14.6 采购订单 - 筛选栏 + 订单列表 + 模态框
     ============================================ -->
        <% If Not viewMode And Request.QueryString("new") <> "1" And Not editMode Then %>
        ' ========== 筛选栏 ==========
        <div class="filter-section">
            <form method="get" class="filter-row">
                <div class="filter-group">
                    <label>状态</label>
                    <select name="status">
                        <option value="">全部</option>
                        <option value="Draft" <% If filterStatus="Draft" Then Response.Write "selected" %>>草稿</option>
                        <option value="Submitted" <% If filterStatus="Submitted" Then Response.Write "selected" %>>待审批</option>
                        <option value="FinanceApproved" <% If filterStatus="FinanceApproved" Then Response.Write "selected" %>>已审批</option>
                        <option value="Ordered" <% If filterStatus="Ordered" Then Response.Write "selected" %>>已下单</option>
                        <option value="PartialReceived" <% If filterStatus="PartialReceived" Then Response.Write "selected" %>>部分收货</option>
                        <option value="Received" <% If filterStatus="Received" Then Response.Write "selected" %>>已收货</option>
                        <option value="Completed" <% If filterStatus="Completed" Then Response.Write "selected" %>>已完成</option>
                        <option value="Rejected" <% If filterStatus="Rejected" Then Response.Write "selected" %>>已拒绝</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>采购分类</label>
                    <select name="category">
                        <option value="">全部</option>
                        <option value="RAW" <% If filterCategory="RAW" Then Response.Write "selected" %>>原材料</option>
                        <option value="BASE" <% If filterCategory="BASE" Then Response.Write "selected" %>>基香原料</option>
                        <option value="PACK" <% If filterCategory="PACK" Then Response.Write "selected" %>>包装材料</option>
                        <option value="MARKET" <% If filterCategory="MARKET" Then Response.Write "selected" %>>营销物料</option>
                        <option value="BOTTLE" <% If filterCategory="BOTTLE" Then Response.Write "selected" %>>瓶子包装</option>
                        <option value="PRINTING" <% If filterCategory="PRINTING" Then Response.Write "selected" %>>印刷品</option>
                        <option value="SPRAYHEAD" <% If filterCategory="SPRAYHEAD" Then Response.Write "selected" %>>喷头配件</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>采购类型</label>
                    <select name="order_type">
                        <option value="">全部</option>
                        <option value="RawMaterial" <% If filterOrderType="RawMaterial" Then Response.Write "selected" %>>原料采购</option>
                        <option value="Packaging" <% If filterOrderType="Packaging" Then Response.Write "selected" %>>包装物采购</option>
                        <option value="Bottle" <% If filterOrderType="Bottle" Then Response.Write "selected" %>>瓶子采购</option>
                        <option value="Printing" <% If filterOrderType="Printing" Then Response.Write "selected" %>>印刷品采购</option>
                        <option value="SprayHead" <% If filterOrderType="SprayHead" Then Response.Write "selected" %>>喷头采购</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label>开始日期</label>
                    <input type="date" name="start_date" value="<%= filterStartDate %>">
                </div>
                <div class="filter-group">
                    <label>结束日期</label>
                    <input type="date" name="end_date" value="<%= filterEndDate %>">
                </div>
                <div class="filter-group">
                    <label>搜索</label>
                    <input type="text" name="keyword" value="<%= Server.HTMLEncode(filterKeyword) %>" placeholder="单号/备注">
                </div>
                <div class="filter-group">
                    <button type="submit" class="btn btn-secondary">
                        <i class="fas fa-filter"></i> 筛选
                    </button>
                </div>
                <div class="filter-group" style="margin-left:auto;">
                    <a href="purchase_orders.asp?new=1<% If filterOrderType <> "" Then %>&order_type=<%= filterOrderType %><% End If %>" class="btn btn-primary">
                        <i class="fas fa-plus"></i> 新建订单
                    </a>
                </div>
            </form>
        </div>

        ' ========== 订单列表 ==========
        <div class="data-section">
            <!-- V11: 批量操作工具栏 -->
            <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap;">
                <button type="button" class="btn btn-sm" style="background:#2196F3;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;" onclick="batchAction('approve')" title="批量财务审批（仅待审批状态）"><i class="fas fa-check"></i> 批量审批</button>
                <button type="button" class="btn btn-sm" style="background:#FF9800;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;" onclick="batchAction('order')" title="批量确认下单（仅已审批状态）"><i class="fas fa-shopping-cart"></i> 批量下单</button>
                <button type="button" class="btn btn-sm" style="background:#4CAF50;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;" onclick="batchAction('receive')" title="批量确认收货（仅已下单状态）"><i class="fas fa-box"></i> 批量收货</button>
                <span id="batchCount" style="color:#888;font-size:12px;margin-left:10px;"></span>
            </div>

            <table class="data-table">
                <thead>
                    <tr>
                        <th style="width:40px;"><input type="checkbox" id="selectAll" onchange="toggleSelectAll(this)" title="全选"></th>
                        <th>采购单号</th>
                        <th>供应商</th>
                        <th>类型</th>
                        <th style="text-align:right;">金额</th>
                        <th>状态</th>
                        <th>订单日期</th>
                        <th style="text-align:center;">操作</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If rsOrders Is Nothing Then
                    %>
                    <tr>
                        <td colspan="8" class="empty-row">
                            <i class="fas fa-inbox" style="font-size: 24px; margin-bottom: 10px; display: block;"></i>
                            暂无采购订单数据
                        </td>
                    </tr>
                    <%
                    Else
                        If rsOrders.EOF Then
                    %>
                    <tr>
                        <td colspan="8" class="empty-row">
                            <i class="fas fa-inbox" style="font-size: 24px; margin-bottom: 10px; display: block;"></i>
                            暂无符合条件的订单
                        </td>
                    </tr>
                    <%
                        Else
                            Do While Not rsOrders.EOF
                    %>
                    <tr>
                        <td><input type="checkbox" class="row-check" value="<%= rsOrders("PurchaseID") %>" onchange="updateBatchCount()"></td>
                        <td><%= Server.HTMLEncode(CStr(rsOrders("PurchaseNo"))) %></td>
                        <td><%= Server.HTMLEncode(GetSupplierName(rsOrders("SupplierID"))) %></td>
                        <td>
                            <%
                            Dim oType : oType = CStr(rsOrders("OrderType") & "")
                            If oType = "Packaging" Then
                                Response.Write "<span class='type-badge type-packaging'>包装物</span>"
                            ElseIf oType = "Bottle" Then
                                Response.Write "<span class='type-badge type-bottle'>瓶子</span>"
                            ElseIf oType = "Printing" Then
                                Response.Write "<span class='type-badge type-printing'>印刷品</span>"
                            ElseIf oType = "SprayHead" Then
                                Response.Write "<span class='type-badge type-sprayhead'>喷头</span>"
                            Else
                                Response.Write "<span class='type-badge type-raw'>原料</span>"
                            End If
                            %>
                        </td>
                        <td style="text-align:right;">¥<%= FormatNumber(SafeNum(rsOrders("TotalAmount")), 2) %></td>
                        <td>
                            <span class="status-badge <%= GetStatusClass(CStr(rsOrders("Status"))) %>">
                                <%= GetStatusName(CStr(rsOrders("Status"))) %>
                            </span>
                        </td>
                        <td><%
                            Dim od : od = rsOrders("OrderDate")
                            If Not IsNull(od) And IsDate(od) Then
                                Response.Write FormatDateTime(od, 2)
                            End If
                        %></td>
                        <td style="text-align:center;">
                            <div class="action-btns" style="justify-content:center;">
                                <a href="purchase_orders.asp?view=<%= rsOrders("PurchaseID") %>" class="btn btn-secondary btn-sm" title="查看">
                                    <i class="fas fa-eye"></i>
                                </a>
                                <% If CStr(rsOrders("Status")) = "Draft" Then %>
                                <a href="purchase_orders.asp?edit=<%= rsOrders("PurchaseID") %>" class="btn btn-primary btn-sm" title="编辑">
                                    <i class="fas fa-edit"></i>
                                </a>
                                <% End If %>
                                <% ' V11: 复制订单 %>
                                <form method="post" style="display:inline;" onsubmit="return confirm('确定复制该订单吗？将创建一个新的草稿订单。')">
                                    <input type="hidden" name="action" value="copy">
                                    <input type="hidden" name="purchase_id" value="<%= rsOrders("PurchaseID") %>">
                                    <button type="submit" class="btn btn-sm" style="background:#9C27B0;color:#fff;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;" title="复制订单">
                                        <i class="fas fa-copy"></i>
                                    </button>
                                </form>
                            </div>
                        </td>
                    </tr>
                    <%
                                rsOrders.MoveNext
                            Loop
                            rsOrders.Close
                            Set rsOrders = Nothing
                        End If
                    End If
                    %>
                </tbody>
            </table>

            ' ========== 分页 ==========
            <% If totalPages > 1 Then %>
            <div class="pagination">
                <% If page > 1 Then %>
                <a href="purchase_orders.asp?page=<%= page-1 %>&status=<%= filterStatus %>&category=<%= filterCategory %>&start_date=<%= filterStartDate %>&end_date=<%= filterEndDate %>&keyword=<%= Server.URLEncode(filterKeyword) %>"><i class="fas fa-chevron-left"></i></a>
                <% End If %>

                <%
                Dim p
                For p = 1 To totalPages
                    If p = page Then
                %>
                <span class="current"><%= p %></span>
                <% Else %>
                <a href="purchase_orders.asp?page=<%= p %>&status=<%= filterStatus %>&category=<%= filterCategory %>&start_date=<%= filterStartDate %>&end_date=<%= filterEndDate %>&keyword=<%= Server.URLEncode(filterKeyword) %>"><%= p %></a>
                <%
                    End If
                Next
                %>

                <% If page < totalPages Then %>
                <a href="purchase_orders.asp?page=<%= page+1 %>&status=<%= filterStatus %>&category=<%= filterCategory %>&start_date=<%= filterStartDate %>&end_date=<%= filterEndDate %>&keyword=<%= Server.URLEncode(filterKeyword) %>"><i class="fas fa-chevron-right"></i></a>
                <% End If %>
            </div>
            <% End If %>
        </div>
        <% End If %>
    </div>

    ' V9: 基香选择模态框
    %>
    <div class="modal-overlay" id="baseNoteModal">
        <div class="modal-dialog">
            <div class="modal-header">
                <h4><i class="fas fa-flask" style="color:#FF9800;"></i> 选择基香原料</h4>
                <button type="button" class="modal-close" onclick="closeBaseNoteModal()">&times;</button>
            </div>
            <div class="modal-body">
                <input type="text" class="modal-search" id="baseNoteSearch" placeholder="搜索基香名称..." oninput="filterBaseNotes()">
                <div class="base-note-grid" id="baseNoteGrid">
                    <div style="grid-column:1/-1;text-align:center;color:#666;padding:20px;">
                        加载中...
                    </div>
                </div>
            </div>
        </div>
    </div>
    <%
    %>
    <div class="modal-overlay" id="historyModal">
        <div class="modal-dialog" style="width:750px;">
            <div class="modal-header">
                <h4><i class="fas fa-history" style="color:#FF9800;"></i> 历史采购产品快速选择</h4>
                <button type="button" class="modal-close" onclick="closeHistoryModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div style="display:flex;gap:10px;margin-bottom:12px;">
                    <input type="text" class="modal-search" id="historySearch" placeholder="搜索产品名称或编码..." oninput="searchHistory()" style="flex:2;">
                    <select id="historySupplierFilter" onchange="searchHistory()" style="flex:1;padding:8px 12px;border-radius:6px;border:1px solid #3a3a3a;background:#252538;color:#e0e0e0;font-size:14px;">
                        <option value="">全部供应商</option>
                    </select>
                </div>
                <div style="max-height:400px;overflow-y:auto;" id="historyGrid">
                    <div style="text-align:center;color:#666;padding:20px;">加载中...</div>
                </div>
            </div>
        </div>
    </div>
    <%
    %>
