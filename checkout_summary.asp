<!-- ============================================
     V14.6 结算页 - 订单摘要和支付方式
     从 checkout.asp 提取
     ============================================ -->
            <div class="checkout-summary">
                <h3><% If FEATURE_I18N Then %><%= T("cart_summary_title", Empty) %><% Else %>订单摘要<% End If %></h3>
                
                <div class="summary-row">
                    <span><% If FEATURE_I18N Then %><%= T("cart_summary_amount", Empty) %><% Else %>商品金额<% End If %>:</span>
                    <span><%= FormatMoney(cartTotal) %></span>
                </div>
                
                <% If totalEngravingFee > 0 Then %>
                <div class="summary-row">
                    <span><% If FEATURE_I18N Then %><%= T("cart_summary_engraving", Empty) %><% Else %>刻字费用<% End If %>:</span>
                    <span><%= FormatMoney(totalEngravingFee) %></span>
                </div>
                <% End If %>
                
                <% If memberDiscountAmount > 0 Then %>
                <div class="summary-row" style="color:#e74c3c;">
                    <span><%= MU_GetLevelName(memberLevel) %><% If FEATURE_I18N Then %><%= T("cart_discount", Empty) %><% Else %>折扣<% End If %> (<%= FormatNumber((1-memberDiscount)*100, 0) %>%OFF):</span>
                    <span>-<%= FormatMoney(memberDiscountAmount) %></span>
                </div>
                <% End If %>
                
                <div class="summary-row">
                    <span><% If FEATURE_I18N Then %><%= T("cart_summary_shipping", Empty) %><% Else %>运费<% End If %>:</span>
                    <% If discountedGrandTotal >= FREE_SHIPPING_AMOUNT Then %>
                    <span><% If FEATURE_I18N Then %><%= T("cart_summary_free_shipping", Empty) %><% Else %>免运费<% End If %></span>
                    <% Else %>
                    <span><%= FormatMoney(SHIPPING_FEE) %></span>
                    <% End If %>
                </div>
                
                <div class="summary-divider"></div>
                
                <div class="summary-total">
                    <span><% If FEATURE_I18N Then %><%= T("cart_summary_total", Empty) %><% Else %>应付总额<% End If %>:</span>
                    <% If discountedGrandTotal >= FREE_SHIPPING_AMOUNT Then %>
                    <span class="total-amount"><%= FormatMoney(discountedGrandTotal) %></span>
                    <% Else %>
                    <span class="total-amount"><%= FormatMoney(discountedGrandTotal + SHIPPING_FEE) %></span>
                    <% End If %>
                </div>
                
                <form method="post" id="paymentForm">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="cart_ids" value="<%= cartIdList %>">
                <h3><% If FEATURE_I18N Then %><%= T("checkout_shipping_info", Empty) %><% Else %>收货信息<% End If %></h3>
                
                <div class="form-group">
                    <label><% If FEATURE_I18N Then %><%= T("checkout_select_address", Empty) %><% Else %>选择收货地址<% End If %></label>
                    <div class="address-selector">
                        <select name="selectedAddress" id="selectedAddress" class="form-control" onchange="loadAddressDetails()">
                            <option value="">-- <% If FEATURE_I18N Then %><%= T("checkout_select_address", Empty) %><% Else %>选择已有地址或添加新地址<% End If %> --</option>
                            <% 
                            ' 获取用户地址列表
                            Dim rsUserAddresses, addrId, addrConsignee, addrPhone, addrProvince, addrCity, addrDistrict, addrDetail, addrIsDefault
                            Dim selectedAddressParam
                            selectedAddressParam = Request.QueryString("selected_address")
                            Set rsUserAddresses = ExecuteQuery("SELECT * FROM UserAddresses WHERE UserID = " & userId & " ORDER BY IsDefault DESC, CreatedAt DESC")
                            If Not rsUserAddresses Is Nothing Then
                                If Not rsUserAddresses.EOF Then
                                    Do While Not rsUserAddresses.EOF
                                        addrId = rsUserAddresses("AddressID")
                                        addrConsignee = rsUserAddresses("Consignee")
                                        addrPhone = rsUserAddresses("Phone")
                                        addrProvince = rsUserAddresses("Province")
                                        addrCity = rsUserAddresses("City")
                                        addrDistrict = rsUserAddresses("District")
                                        addrDetail = rsUserAddresses("Address")
                                        addrIsDefault = rsUserAddresses("IsDefault")
                                        %>
                                    <option value="<%= addrId %>"<% If (addrIsDefault <> 0) Or (selectedAddressParam <> "" And CLng(selectedAddressParam) = CLng(addrId)) Then Response.Write " selected" End If %>><%= HTMLEncode(addrConsignee) %> <%= HTMLEncode(addrPhone) %> <%= HTMLEncode(BuildFullAddress(addrProvince, addrCity, addrDistrict, addrDetail)) %></option>
                                    <%
                                        rsUserAddresses.MoveNext
                                    Loop
                                End If
                                rsUserAddresses.Close
                                Set rsUserAddresses = Nothing
                            End If
                            %>
                            <option value="new"<% If selectedAddressParam <> "" And selectedAddressParam = "new" Then Response.Write " selected" End If %>>+ <% If FEATURE_I18N Then %><%= T("checkout_add_address", Empty) %><% Else %>添加新地址<% End If %></option>
                        </select>
                        <button type="button" class="btn btn-secondary" onclick="showAddressForm()" style="margin-top: 10px;">
                            <i class="fas fa-plus"></i> <% If FEATURE_I18N Then %><%= T("checkout_new_address_btn", Empty) %><% Else %>新增收货地址<% End If %>
                        </button>
                    </div>
                </div>
                
                <div id="selectedAddressDisplay" style="display:block;">
                    <!-- BUG FIX: 原代码使用 userRealName（未定义），已修正为 userFullName -->
                    <% If userFullName <> "" Then %>
                    <div class="selected-address-info">
                        <p><strong><% If FEATURE_I18N Then %><%= T("checkout_address_current", Empty) %><% Else %>当前地址<% End If %>：</strong><%= HTMLEncode(userFullName) %> <%= HTMLEncode(userPhone) %> <%= HTMLEncode(userAddress) %></p>
                    </div>
                    <% End If %>
                </div>
                
                <h3><% If FEATURE_I18N Then %><%= T("checkout_payment_method", Empty) %><% Else %>支付方式<% End If %></h3>
                
                    <div class="payment-methods">
                        <% If enableCOD = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_COD %>" checked>
                                <span class="radio-text"><% If FEATURE_I18N Then %><%= T("payment_cod", Empty) %><% Else %>货到付款<% End If %></span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enableWechat = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_WECHAT %>" <% If enableCOD <> "1" Then Response.Write "checked" End If %>>
                                <span class="radio-text"><% If FEATURE_I18N Then %><%= T("payment_wechat", Empty) %><% Else %>微信支付<% End If %></span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enableAlipay = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_ALIPAY %>">
                                <span class="radio-text"><% If FEATURE_I18N Then %><%= T("payment_alipay", Empty) %><% Else %>支付宝<% End If %></span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enablePaypal = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="<%= PAYMENT_METHOD_PAYPAL %>">
                                <span class="radio-text">PayPal</span>
                            </label>
                        </div>
                        <% End If %>
                        
                        <% If enableBankTransfer = "1" Then %>
                        <div class="payment-method">
                            <label class="radio-label">
                                <input type="radio" name="payment_method" value="5">
                                <span class="radio-text"><% If FEATURE_I18N Then %><%= T("payment_bank_transfer", Empty) %><% Else %>银行转账<% End If %></span>
                            </label>
                        </div>
                        <% End If %>
                    </div>
                    
                    <% If enableCOD <> "1" And enableWechat <> "1" And enableAlipay <> "1" And enablePaypal <> "1" And enableBankTransfer <> "1" Then %>
                    <div class="alert alert-warning">
                        <i class="fas fa-exclamation-triangle"></i> <% If FEATURE_I18N Then %><%= T("checkout_no_payment", Empty) %><% Else %>暂无可用的支付方式，请联系客服。<% End If %>
                    </div>
                    <% End If %>
                    
                    <button type="submit" class="btn btn-primary btn-lg btn-block">
                        <i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("checkout_confirm_pay", Empty) %><% Else %>确认订单并支付<% End If %>
                    </button>
                </form>