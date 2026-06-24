<!-- ============================================
     V14.6 产品设置 - 类型配置标签页模板
     ============================================ -->
        <% ElseIf currentTab = "types" Then %>
        <!-- 类型配置Tab -->
        <div class="admin-card">
            <div class="admin-card-header">
                <h3 class="admin-card-title"><i class="fas fa-tags"></i> 产品类型配置</h3>
            </div>
            <div class="admin-card-body">
                <div class="type-grid">
                    <% 
                    If Not rsTypeConfig Is Nothing Then
                        Do While Not rsTypeConfig.EOF
                            Dim tcId, tcCode, tcDisplay, tcNav, tcDesc, tcIcon, tcReview, tcRatio, tcOrder, tcActive
                            tcId = rsTypeConfig("ConfigID")
                            tcCode = rsTypeConfig("TypeCode")
                            tcDisplay = rsTypeConfig("DisplayName")
                            tcNav = rsTypeConfig("NavName") & ""
                            tcDesc = rsTypeConfig("Description") & ""
                            tcIcon = rsTypeConfig("Icon") & ""
                            tcReview = rsTypeConfig("RequiresReview")
                            tcRatio = rsTypeConfig("RequiresRatio")
                            tcOrder = rsTypeConfig("DisplayOrder")
                            tcActive = rsTypeConfig("IsActive")
                            
                            ' 获取该类型产品数量
                            Dim typeTotal, typeActive
                            typeTotal = 0
                            typeActive = 0
                            If productStats.Exists(tcCode) Then
                                typeTotal = productStats(tcCode)(0)
                                typeActive = productStats(tcCode)(1)
                            End If
                            
                            ' 设置类型样式
                            Dim tcClass
                            Select Case LCase(tcCode)
                                Case "standard": tcClass = "status-fixed"
                                Case "custom": tcClass = "status-custom"
                                Case "kol": tcClass = "status-kol"
                                Case Else: tcClass = ""
                            End Select
                    %>
                    <div class="type-card">
                        <div class="type-card-header">
                            <div class="type-icon">
                                <% If tcIcon <> "" Then %>
                                <i class="<%= tcIcon %>"></i>
                                <% Else %>
                                <i class="fas fa-box"></i>
                                <% End If %>
                            </div>
                            <div class="type-info">
                                <h4><%= HTMLEncode(tcDisplay) %></h4>
                                <span class="type-code"><%= tcCode %></span>
                            </div>
                        </div>
                        
                        <div class="type-stats">
                            <div class="type-stat">
                                <div class="type-stat-value"><%= typeActive %></div>
                                <div class="type-stat-label">上架产品</div>
                            </div>
                            <div class="type-stat">
                                <div class="type-stat-value"><%= typeTotal %></div>
                                <div class="type-stat-label">总产品</div>
                            </div>
                            <div class="type-stat">
                                <div class="type-stat-value"><%= tcOrder %></div>
                                <div class="type-stat-label">排序</div>
                            </div>
                        </div>
                        
                        <div class="type-features">
                            <span class="type-feature <%= IIf(tcReview, "active", "") %>">
                                <i class="fas <%= IIf(tcReview, "fa-check", "fa-times") %>"></i> 需要审核
                            </span>
                            <span class="type-feature <%= IIf(tcRatio, "active", "") %>">
                                <i class="fas <%= IIf(tcRatio, "fa-check", "fa-times") %>"></i> 需要配比
                            </span>
                            <span class="type-feature <%= IIf(tcActive, "active", "") %>">
                                <i class="fas <%= IIf(tcActive, "fa-check", "fa-times") %>"></i> 已启用
                            </span>
                        </div>
                        
                        <% If tcDesc <> "" Then %>
                        <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                            <%= HTMLEncode(Left(tcDesc, 50)) %><%= IIf(Len(tcDesc) > 50, "...", "") %>
                        </p>
                        <% End If %>
                        
                        <div class="action-btns">
                            <button class="admin-btn admin-btn-sm admin-btn-outline" onclick="showEditTypeForm(this)"
                                data-id="<%= tcId %>"
                                data-code="<%= SafeOutput(tcCode) %>"
                                data-display="<%= SafeOutput(tcDisplay) %>"
                                data-nav="<%= SafeOutput(tcNav) %>"
                                data-desc="<%= SafeOutput(tcDesc) %>"
                                data-icon="<%= SafeOutput(tcIcon) %>"
                                data-review="<%= tcReview %>"
                                data-ratio="<%= tcRatio %>"
                                data-order="<%= tcOrder %>"
                                data-active="<%= tcActive %>">
                                <i class="fas fa-edit"></i> 编辑
                            </button>
                        </div>
                    </div>
                    <% 
                            rsTypeConfig.MoveNext
                        Loop
                        rsTypeConfig.Close
                        Set rsTypeConfig = Nothing
                    End If
                    %>
                </div>
            </div>
        </div>