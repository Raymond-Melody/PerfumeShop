<!-- ============================================
     V14.6 产品设置 - 香调配比参数标签页模板
     ============================================ -->
        <% ElseIf currentTab = "ratio" Then %>
        <!-- 香调配比参数Tab -->
        <div class="admin-card">
            <div class="admin-card-header">
                <h3 class="admin-card-title"><i class="fas fa-percentage"></i> 香调配比参数设置</h3>
                <p style="color: #888; font-size: 14px; margin-top: 10px;">
                    设置定制香水和KOL推荐商品的前、中、后调最小比例限制，确保配方平衡
                </p>
            </div>
            <div class="admin-card-body">
                <% If Request.QueryString("msg") <> "" Then %>
                <div class="alert alert-success" style="margin-bottom: 20px;">
                    <i class="fas fa-check-circle"></i> <%= Server.HTMLEncode(Request.QueryString("msg")) %>
                </div>
                <% End If %>
                
                <form method="post" action="product_settings.asp?tab=ratio">
                    <input type="hidden" name="action" value="save_ratio_settings">
                    
                    <div class="type-grid">
                        <!-- 前调最小比例 -->
                        <div class="type-card">
                            <div class="type-card-header">
                                <div class="type-icon" style="background: #e8f5e9; color: #2e7d32;">
                                    <i class="fas fa-wind"></i>
                                </div>
                                <div class="type-info">
                                    <h4>前调最小比例</h4>
                                    <span class="type-code">Top Note</span>
                                </div>
                            </div>
                            <div style="padding: 15px;">
                                <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                                    前调是香水的第一印象，设置最小比例确保香水有足够的首香特征。
                                </p>
                                <div class="admin-form-group">
                                    <label class="admin-form-label">最小比例 (%)</label>
                                    <input type="number" name="minTopPercent" value="<%= minTopPercent %>" 
                                        min="0" max="100" step="1" class="admin-form-control" required>
                                </div>
                                <div style="margin-top: 10px; padding: 8px; background: #e8f5e9; border-radius: 4px; font-size: 13px; color: #2e7d32;">
                                    <i class="fas fa-info-circle"></i> 建议值：10% - 30%
                                </div>
                            </div>
                        </div>
                        
                        <!-- 中调最小比例 -->
                        <div class="type-card">
                            <div class="type-card-header">
                                <div class="type-icon" style="background: #fff3e0; color: #e65100;">
                                    <i class="fas fa-heart"></i>
                                </div>
                                <div class="type-info">
                                    <h4>中调最小比例</h4>
                                    <span class="type-code">Middle Note</span>
                                </div>
                            </div>
                            <div style="padding: 15px;">
                                <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                                    中调是香水的核心灵魂，设置最小比例确保香水有持久的主香特征。
                                </p>
                                <div class="admin-form-group">
                                    <label class="admin-form-label">最小比例 (%)</label>
                                    <input type="number" name="minMiddlePercent" value="<%= minMiddlePercent %>" 
                                        min="0" max="100" step="1" class="admin-form-control" required>
                                </div>
                                <div style="margin-top: 10px; padding: 8px; background: #fff3e0; border-radius: 4px; font-size: 13px; color: #e65100;">
                                    <i class="fas fa-info-circle"></i> 建议值：10% - 40%
                                </div>
                            </div>
                        </div>
                        
                        <!-- 后调最小比例 -->
                        <div class="type-card">
                            <div class="type-card-header">
                                <div class="type-icon" style="background: #f3e5f5; color: #7b1fa2;">
                                    <i class="fas fa-moon"></i>
                                </div>
                                <div class="type-info">
                                    <h4>后调最小比例</h4>
                                    <span class="type-code">Base Note</span>
                                </div>
                            </div>
                            <div style="padding: 15px;">
                                <p style="font-size: 13px; color: #888; margin-bottom: 15px;">
                                    后调是香水的持久余韵，设置最小比例确保香水有足够的留香时间。
                                </p>
                                <div class="admin-form-group">
                                    <label class="admin-form-label">最小比例 (%)</label>
                                    <input type="number" name="minBasePercent" value="<%= minBasePercent %>" 
                                        min="0" max="100" step="1" class="admin-form-control" required>
                                </div>
                                <div style="margin-top: 10px; padding: 8px; background: #f3e5f5; border-radius: 4px; font-size: 13px; color: #7b1fa2;">
                                    <i class="fas fa-info-circle"></i> 建议值：10% - 30%
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div style="margin-top: 30px; padding: 20px; background: #e3f2fd; border: 1px solid #2196F3; border-radius: 8px;">
                        <h4 style="margin: 0 0 15px 0; color: #1565c0;">
                            <i class="fas fa-lightbulb"></i> 设置说明
                        </h4>
                        <ul style="margin: 0; padding-left: 20px; color: #1565c0; line-height: 1.8;">
                            <li>该设置适用于<strong>定制香水</strong>和<strong>KOL推荐</strong>两种商品类型</li>
                            <li>用户在前台购买时，系统会验证前、中、后调的比例是否都达到最小值</li>
                            <li>后台管理员新增KOL商品时，也会验证该配比规则</li>
                            <li>建议三种调性的最小比例之和不超过60%，以保留调配灵活性</li>
                            <li>当前设置：前调 <strong><%= minTopPercent %>%</strong> | 中调 <strong><%= minMiddlePercent %>%</strong> | 后调 <strong><%= minBasePercent %>%</strong></li>
                        </ul>
                    </div>
                    
                    <div style="margin-top: 30px; text-align: center;">
                        <button type="submit" class="admin-btn admin-btn-primary" style="padding: 12px 40px; font-size: 16px;" <%= IIf(isManager, "", "disabled") %>>
                            <i class="fas fa-save"></i> 保存设置
                        </button>
                    </div>
                </form>
            </div>
        </div>