<!-- ============================================
     V14.6 结算页 - 地址模态框
     从 checkout.asp 提取
     ============================================ -->
                <!-- 添加/编辑地址弹窗 -->
                <div class="modal" id="addressModal">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h3 id="modalTitle">新增收货地址</h3>
                            <span class="close" onclick="closeAddressForm()">&times;</span>
                        </div>
                        <div class="modal-body">
                            <form id="addressForm" method="post" action="checkout.asp">
                                <%= GetCSRFTokenField() %>
                                <input type="hidden" id="formAction" name="action" value="add">
                                <input type="hidden" id="formAddressId" name="addressId" value="">
                                <input type="hidden" name="cart_ids" value="<%= cartIdList %>">
                                <input type="hidden" name="payment_method" value="" id="addressFormPaymentMethod">
                                <div class="form-row">
                                    <div class="form-group">
                                        <label for="consignee">收货人姓名 *</label>
                                        <input type="text" id="consignee" name="realName" required>
                                    </div>
                                    <div class="form-group">
                                        <label for="phone">联系电话 *</label>
                                        <input type="tel" id="phone" name="phone" required>
                                    </div>
                                </div>
                                    
                                <div class="form-group">
                                    <label for="province">所在地区 *</label>
                                    <select id="province" name="province" required onchange="updateCities()">
                                        <option value="">请选择省份</option>
                                        <option value="北京市">北京市</option>
                                        <option value="上海市">上海市</option>
                                        <option value="天津市">天津市</option>
                                        <option value="重庆市">重庆市</option>
                                        <option value="河北省">河北省</option>
                                        <option value="山西省">山西省</option>
                                        <option value="辽宁省">辽宁省</option>
                                        <option value="吉林省">吉林省</option>
                                        <option value="黑龙江省">黑龙江省</option>
                                        <option value="江苏省">江苏省</option>
                                        <option value="浙江省">浙江省</option>
                                        <option value="安徽省">安徽省</option>
                                        <option value="福建省">福建省</option>
                                        <option value="江西省">江西省</option>
                                        <option value="山东省">山东省</option>
                                        <option value="河南省">河南省</option>
                                        <option value="湖北省">湖北省</option>
                                        <option value="湖南省">湖南省</option>
                                        <option value="广东省">广东省</option>
                                        <option value="海南省">海南省</option>
                                        <option value="四川省">四川省</option>
                                        <option value="贵州省">贵州省</option>
                                        <option value="云南省">云南省</option>
                                        <option value="陕西省">陕西省</option>
                                        <option value="甘肃省">甘肃省</option>
                                        <option value="青海省">青海省</option>
                                        <option value="台湾省">台湾省</option>
                                        <option value="内蒙古自治区">内蒙古自治区</option>
                                        <option value="广西壮族自治区">广西壮族自治区</option>
                                        <option value="西藏自治区">西藏自治区</option>
                                        <option value="宁夏回族自治区">宁夏回族自治区</option>
                                        <option value="新疆维吾尔自治区">新疆维吾尔自治区</option>
                                        <option value="香港特别行政区">香港特别行政区</option>
                                        <option value="澳门特别行政区">澳门特别行政区</option>
                                    </select>
                                    <select id="city" name="city" required onchange="updateDistricts()">
                                        <option value="">请选择城市</option>
                                    </select>
                                    <select id="district" name="district" required>
                                        <option value="">请选择区县</option>
                                    </select>
                                </div>
                                    
                                <div class="form-group">
                                    <label for="address">详细地址 *</label>
                                    <input type="text" id="address" name="address" placeholder="请输入详细地址，如街道、门牌号等" required>
                                </div>
                                    
                                <div class="form-group">
                                    <label class="checkbox-label">
                                        <input type="checkbox" id="isDefault" name="isDefault" value="1">
                                        设为默认地址
                                    </label>
                                </div>
                                    
                                <div class="form-actions">
                                    <button type="submit" class="btn btn-primary">保存地址</button>
                                    <button type="button" class="btn btn-text" onclick="closeAddressForm()">取消</button>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>
</div>