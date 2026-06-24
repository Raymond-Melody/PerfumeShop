<!-- ============================================
     V14.6 采购订单 - 查看订单详情模板
     ============================================ -->
        <% If viewMode Then %>
        ' ========== 查看订单详情 ==========
        <div class="view-section">
            <div class="view-header">
                <div>
                    <h3 style="margin:0 0 10px 0;">
                        <i class="fas fa-eye" style="color:#FF9800;"></i>
                        订单详情：<%= Server.HTMLEncode(CStr(viewOrderData("PurchaseNo"))) %>
                    </h3>
                    <span class="status-badge <%= GetStatusClass(CStr(viewOrderData("Status"))) %>">
                        <%= GetStatusName(CStr(viewOrderData("Status"))) %>
                    </span>
                </div>
                <div>
                    <a href="purchase_orders.asp" class="btn btn-secondary">
                        <i class="fas fa-arrow-left"></i> 返回列表
                    </a>
                    <button type="button" class="btn btn-secondary" style="margin-left:8px;" onclick="copyOrderFromView(<%= viewOrderID %>)">
                        <i class="fas fa-copy"></i> 复制订单
                    </button>
                </div>
            </div>

            <div class="view-info-grid">
                <div class="view-info-item">
                    <label>供应商</label>
                    <value><%= Server.HTMLEncode(GetSupplierName(viewOrderData("SupplierID"))) %></value>
                </div>
                <div class="view-info-item">
                    <label>采购分类</label>
                    <value><%= GetCategoryName(CStr(viewOrderData("CategoryCode"))) %></value>
                </div>
                <div class="view-info-item">
                    <label>订单日期</label>
                    <value><% If IsDate(viewOrderData("OrderDate")) Then Response.Write FormatDateTime(viewOrderData("OrderDate"), 2) End If %></value>
                </div>
                <div class="view-info-item">
                    <label>期望交期</label>
                    <value>
                        <% If IsNull(viewOrderData("ExpectedDate")) Then %>
                            未设置
                        <% Else %>
                            <% If IsDate(viewOrderData("ExpectedDate")) Then Response.Write FormatDateTime(viewOrderData("ExpectedDate"), 2) End If %>
                        <% End If %>
                    </value>
                </div>
                <div class="view-info-item">
                    <label>订单金额</label>
                    <value style="color:#FF9800;font-size:18px;">¥<%= FormatNumber(SafeNum(viewOrderData("TotalAmount")), 2) %></value>
                </div>
                <div class="view-info-item">
                    <label>创建人</label>
                    <value><%= SafeNum(viewOrderData("CreatedBy")) %></value>
                </div>
            </div>

            <% Dim remarksVal : remarksVal = viewOrderData("Remarks") & "" : If remarksVal <> "" Then %>
            <div style="margin-bottom:20px;">
                <label style="font-size:12px;color:#888;display:block;margin-bottom:5px;">备注</label>
                <div style="background:rgba(255,255,255,0.02);padding:12px 15px;border-radius:8px;">
                    <%= Server.HTMLEncode(remarksVal) %>
                </div>
            </div>
            <% End If %>

            <h4 style="margin:20px 0 15px 0;color:#fff;"><i class="fas fa-list"></i> 采购明细</h4>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>物料名称</th>
                        <th>物料编码</th>
                        <th>规格</th>
                        <th>单位</th>
                        <th style="text-align:right;">数量</th>
                        <th style="text-align:right;">单价</th>
                        <th style="text-align:right;">小计</th>
                        <th style="text-align:right;">已收货</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If rsViewDetails Is Nothing Then
                    %>
                    <tr>
                        <td colspan="7" class="empty-row">暂无明细</td>
                    </tr>
                    <%
                    Else
                        If rsViewDetails.EOF Then
                    %>
                    <tr>
                        <td colspan="7" class="empty-row">暂无明细</td>
                    </tr>
                    <%
                        Else
                            Do While Not rsViewDetails.EOF
                    %>
                    <tr>
                        <td><%= Server.HTMLEncode(CStr(rsViewDetails("ItemName"))) %></td>
                        <td>
                            <% If IsNull(rsViewDetails("ItemCode")) Then %>
                                -
                            <% Else %>
                                <%= Server.HTMLEncode(CStr(rsViewDetails("ItemCode"))) %>
                            <% End If %>
                        </td>
                        <td>
                            <% If IsNull(rsViewDetails("Specification")) Then %>
                                -
                            <% Else %>
                                <%= Server.HTMLEncode(CStr(rsViewDetails("Specification"))) %>
                            <% End If %>
                        </td>
                        <td>
                            <% If IsNull(rsViewDetails("Unit")) Then %>
                                -
                            <% Else %>
                                <%= Server.HTMLEncode(CStr(rsViewDetails("Unit"))) %>
                            <% End If %>
                        </td>
                        <td style="text-align:right;"><%= SafeNum(rsViewDetails("Quantity")) %></td>
                        <td style="text-align:right;">¥<%= FormatNumber(SafeNum(rsViewDetails("UnitPrice")), 2) %></td>
                        <td style="text-align:right;">¥<%= FormatNumber(SafeNum(rsViewDetails("TotalPrice")), 2) %></td>
                        <td style="text-align:right;"><%= SafeNum(rsViewDetails("ReceivedQty")) %></td>
                    </tr>
                    <%
                                rsViewDetails.MoveNext
                            Loop
                            rsViewDetails.Close
                            Set rsViewDetails = Nothing
                        End If
                    End If
                    %>
                </tbody>
            </table>


            <% ' ========== V12: 操作时间线 ========== %>
            <h4 style="margin:20px 0 15px 0;color:#fff;"><i class="fas fa-history"></i> 操作时间线</h4>
            <div class="timeline">
                <%
                If Not rsStatusLog Is Nothing Then
                    If Not rsStatusLog.EOF Then
                        Dim tlIdx : tlIdx = 0
                        Do While Not rsStatusLog.EOF
                            Dim tlFromStatus : tlFromStatus = CStr(rsStatusLog("FromStatus") & "")
                            Dim tlToStatus : tlToStatus = CStr(rsStatusLog("ToStatus"))
                            Dim tlLogTime : tlLogTime = rsStatusLog("ChangedAt")
                            Dim tlChanger : tlChanger = CStr(rsStatusLog("ChangedBy") & "")
                            Dim tlRemark : tlRemark = CStr(rsStatusLog("Remarks") & "")
                %>
                <div class="timeline-item">
                    <div class="timeline-dot <%= IIf(tlIdx=0, "active", "") %>"></div>
                    <div class="timeline-content">
                        <div class="timeline-header">
                            <span class="status-badge <%= GetStatusClass(tlToStatus) %>"><%= GetStatusName(tlToStatus) %></span>
                            <span class="timeline-time"><% If IsDate(tlLogTime) Then Response.Write FormatDateTime(tlLogTime, 2) & " " & FormatDateTime(tlLogTime, 4) End If %></span>
                        </div>
                        <div class="timeline-desc">
                            <% If tlFromStatus <> "" Then %>
                                <%= GetStatusName(tlFromStatus) %> &rarr; <%= GetStatusName(tlToStatus) %>
                            <% Else %>
                                创建订单（初始状态：<%= GetStatusName(tlToStatus) %>）
                            <% End If %>
                            <% If tlRemark <> "" Then %> &mdash; <em><%= Server.HTMLEncode(tlRemark) %></em><% End If %>
                        </div>
                        <% If tlChanger <> "" Then %>
                        <div class="timeline-actor"><i class="fas fa-user"></i> <%= Server.HTMLEncode(tlChanger) %></div>
                        <% End If %>
                    </div>
                </div>
                <%
                            tlIdx = tlIdx + 1
                            rsStatusLog.MoveNext
                        Loop
                        rsStatusLog.Close
                        Set rsStatusLog = Nothing
                    Else
                        rsStatusLog.Close
                        Set rsStatusLog = Nothing
                %>
                <div class="timeline-item">
                    <div class="timeline-dot"></div>
                    <div class="timeline-content">
                        <div class="timeline-desc" style="color:#666;">暂无状态变更记录</div>
                    </div>
                </div>
                <%
                    End If
                Else
                %>
                <div class="timeline-item">
                    <div class="timeline-dot"></div>
                    <div class="timeline-content">
                        <div class="timeline-desc" style="color:#666;">暂无状态变更记录</div>
                    </div>
                </div>
                <% End If %>
            </div>
<% ' ========== 状态操作按钮 ========== %>
            <div style="margin-top:25px;padding-top:20px;border-top:1px solid rgba(255,255,255,0.05);">
                <h4 style="margin:0 0 15px 0;color:#fff;"><i class="fas fa-exchange-alt"></i> 状态操作</h4>
                <div class="action-btns">
                    <%
                    Dim viewStatus
                    viewStatus = CStr(viewOrderData("Status"))

                    If viewStatus = "Draft" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Submitted">
                        <button type="submit" class="btn btn-success" onclick="return confirm('确定提交审批吗？');">
                            <i class="fas fa-paper-plane"></i> 提交审批
                        </button>
                    </form>
                    <a href="purchase_orders.asp?edit=<%= viewOrderID %>" class="btn btn-primary">
                        <i class="fas fa-edit"></i> 编辑
                    </a>
                    <%
                    ElseIf viewStatus = "Submitted" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="FinanceApproved">
                        <button type="submit" class="btn btn-primary" onclick="return confirm('确定通过财务审批吗？');">
                            <i class="fas fa-check"></i> 财务审批
                        </button>
                    </form>
                    <%
                    ElseIf viewStatus = "FinanceApproved" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Ordered">
                        <button type="submit" class="btn btn-primary" onclick="return confirm('确定标记为已下单吗？');">
                            <i class="fas fa-shopping-cart"></i> 确认下单
                        </button>
                    </form>
                    <%
                    ElseIf viewStatus = "Ordered" Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Received">
                        <button type="submit" class="btn btn-success" onclick="return confirm('确定标记为已收货吗？');">
                            <i class="fas fa-box"></i> 确认收货
                        </button>
                    </form>
                    <%
                    ElseIf viewStatus = "Received" And isManager Then
                    %>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="changestatus">
                        <input type="hidden" name="purchase_id" value="<%= viewOrderID %>">
                        <input type="hidden" name="new_status" value="Completed">
                        <button type="submit" class="btn btn-success" onclick="return confirm('确定完成订单吗？');">
                            <i class="fas fa-check-circle"></i> 完成订单
                        </button>
                    </form>
                    <% End If %>
                </div>
            </div>
        </div>
        <%
        viewOrderData.Close
        Set viewOrderData = Nothing
        End If
        %>