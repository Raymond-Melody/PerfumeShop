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
                
                <% If FEATURE_COUPON_SYSTEM Then %>
                <div class="summary-row coupon-row" style="background:#e8f5e9;margin:8px 0;padding:10px 12px;border-radius:8px;border:1px dashed #4CAF50;">
                    <span><i class="fas fa-ticket-alt" style="color:#4CAF50;"></i> 优惠码:</span>
                    <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                        <input type="text" 
                               id="couponCode" 
                               name="coupon_code" 
                               value="" 
                               style="width:150px;padding:6px 10px;border:1px solid #4CAF50;border-radius:6px;font-size:14px;text-transform:uppercase;"
                               placeholder="输入优惠码">
                        <button type="button" class="btn btn-sm" onclick="validateCoupon()" style="background:#4CAF50;color:#fff;border:none;padding:6px 14px;border-radius:6px;font-size:12px;cursor:pointer;">验证</button>
                        <span id="couponStatus" style="font-size:12px;"></span>
                    </div>
                </div>
                <script>
                var couponApplied = false;
                var couponDiscount = 0;
                function validateCoupon() {
                    var code = document.getElementById('couponCode').value.trim();
                    var statusEl = document.getElementById('couponStatus');
                    if (!code) { statusEl.innerHTML = ''; couponApplied = false; couponDiscount = 0; return; }
                    statusEl.innerHTML = '<span style="color:#888;"><i class="fas fa-spinner fa-spin"></i> 验证中...</span>';
                    var xhr = new XMLHttpRequest();
                    xhr.open('POST', '/api/coupon_validate.asp', true);
                    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
                    xhr.onload = function() {
                        try {
                            var resp = JSON.parse(xhr.responseText);
                            if (resp.success) {
                                statusEl.innerHTML = '<span style="color:#4CAF50;font-weight:bold;"><i class="fas fa-check-circle"></i> 可优惠 ¥' + resp.data.discount.toFixed(2) + '</span>';
                                couponApplied = true;
                                couponDiscount = resp.data.discount;
                            } else {
                                statusEl.innerHTML = '<span style="color:#f44336;">' + resp.message + '</span>';
                                couponApplied = false;
                                couponDiscount = 0;
                            }
                        } catch(e) {
                            statusEl.innerHTML = '<span style="color:#f44336;">网络错误</span>';
                        }
                    };
                    xhr.send('code=' + encodeURIComponent(code) + '&cart_total=' + <%= discountedGrandTotal %>);
                }
                </script>
                <% End If %>
                
                <% If FEATURE_POINTS_SYSTEM Then
                    Dim userAvailablePoints, maxRedeemPoints, redeemRate, maxRedeemPct, maxRedeemValue
                    userAvailablePoints = PE_GetAvailablePoints(userId)
                    redeemRate = PE_GetRule("redeem_discount_rate")
                    If redeemRate <= 0 Then redeemRate = 100
                    maxRedeemPct = PE_GetRule("max_redeem_pct") / 100
                    If maxRedeemPct <= 0 Then maxRedeemPct = 0.3
                    maxRedeemValue = discountedGrandTotal * maxRedeemPct
                    maxRedeemPoints = Int(maxRedeemValue * redeemRate)
                    If maxRedeemPoints > userAvailablePoints Then maxRedeemPoints = userAvailablePoints
                    If maxRedeemPoints < 0 Then maxRedeemPoints = 0
                %>
                <div class="summary-row points-redeem-row" style="background:#fff8e1;margin:8px 0;padding:10px 12px;border-radius:8px;border:1px dashed #ffcc02;">
                    <span><i class="fas fa-coins" style="color:#ff8f00;"></i> 积分抵扣 (可用 <strong><%= userAvailablePoints %></strong> 积分):</span>
                    <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                        <input type="number" 
                               id="pointsToRedeem" 
                               name="points_to_redeem" 
                               min="0" 
                               max="<%= maxRedeemPoints %>" 
                               value="0" 
                               step="<%= CInt(redeemRate) %>"
                               style="width:110px;padding:6px 10px;border:1px solid #ffcc02;border-radius:6px;font-size:14px;text-align:center;"
                               onchange="updatePointsPreview()" 
                               oninput="updatePointsPreview()"
                               placeholder="0">
                        <button type="button" class="btn btn-sm" onclick="useAllPoints()" style="background:#ff8f00;color:#fff;border:none;padding:4px 10px;border-radius:4px;font-size:12px;cursor:pointer;">全部使用</button>
                        <span id="pointsValuePreview" style="font-size:13px;color:#e65100;">-¥0.00</span>
                    </div>
                </div>
                <script>
                var redeemRate = <%= redeemRate %>;
                var maxRedeem = <%= maxRedeemPoints %>;
                function updatePointsPreview() {
                    var input = document.getElementById('pointsToRedeem');
                    var preview = document.getElementById('pointsValuePreview');
                    var val = parseInt(input.value) || 0;
                    if (val < 0) { input.value = 0; val = 0; }
                    if (val > maxRedeem) { input.value = maxRedeem; val = maxRedeem; }
                    var discount = (val / redeemRate).toFixed(2);
                    preview.textContent = '-¥' + discount;
                }
                function useAllPoints() {
                    var input = document.getElementById('pointsToRedeem');
                    input.value = maxRedeem;
                    updatePointsPreview();
                }
                </script>
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

                    <!-- V18: 记住支付方式 -->
                    <div class="remember-payment" style="margin:16px 0;display:flex;align-items:center;gap:10px;">
                        <label class="checkbox-label" style="display:flex;align-items:center;gap:8px;cursor:pointer;">
                            <input type="checkbox" name="remember_payment" value="1" id="rememberPayment" style="width:18px;height:18px;">
                            <span style="font-size:0.9rem;color:#666;"><% If FEATURE_I18N Then %>记住我的支付方式<% Else %>记住我的支付方式<% End If %></span>
                        </label>
                    </div>

                    <button type="submit" class="btn btn-primary btn-lg btn-block" style="padding:14px 24px;font-size:1rem;border-radius:10px;">
                        <i class="fas fa-check"></i> <% If FEATURE_I18N Then %><%= T("checkout_confirm_pay", Empty) %><% Else %>确认订单并支付<% End If %>
                    </button>
                </form>