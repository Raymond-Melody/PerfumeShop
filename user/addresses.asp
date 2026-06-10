<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 检查用户是否登录
If Session("UserID") = "" Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("URL"))
End If

Call OpenConnection()

' 处理表单提交
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' CSRF验证
    If Not ValidateCSRFToken() Then
        Response.Write "<script>alert('安全验证失败，请刷新页面重试'); history.back();</script>"
        Response.End
    End If
    
    Dim action
    action = Request.Form("action")
        
    If action = "add" Or action = "edit" Then
        Dim consignee, phone, province, city, district, address, isDefault
        consignee = SafeSQL(Trim(Request.Form("consignee")))
        phone = SafeSQL(Trim(Request.Form("phone")))
        province = SafeSQL(Trim(Request.Form("province")))
        city = SafeSQL(Trim(Request.Form("city")))
        district = SafeSQL(Trim(Request.Form("district")))
        address = SafeSQL(Trim(Request.Form("address")))
        isDefault = Request.Form("isDefault")
            
        If isDefault <> "" And isDefault <> "0" Then
            isDefault = 1
        Else
            isDefault = 0
        End If
            
        Dim sql
        If action = "add" Then
            ' 如果设为默认地址，先取消其他默认地址
            If isDefault <> 0 Then
                Call ExecuteNonQuery("UPDATE UserAddresses SET IsDefault = 0 WHERE UserID = " & Session("UserID"))
            End If
                
            sql = "INSERT INTO UserAddresses (UserID, Consignee, Phone, Province, City, District, Address, IsDefault, CreatedAt) VALUES (" & Session("UserID") & ", '" & consignee & "', '" & phone & "', '" & province & "', '" & city & "', '" & district & "', '" & address & "', " & isDefault & ", GETDATE())"
        ElseIf action = "edit" Then
            Dim addressId
            addressId = Request.Form("addressId")
            If Not IsNumeric(addressId) Then
                Response.Write "<script>alert('无效的地址ID'); history.back();</script>"
                Response.End
            End If
                
            ' 如果设为默认地址，先取消其他默认地址
            If isDefault <> 0 Then
                Call ExecuteNonQuery("UPDATE UserAddresses SET IsDefault = 0 WHERE UserID = " & Session("UserID"))
            End If
                
            sql = "UPDATE UserAddresses SET Consignee = '" & consignee & "', Phone = '" & phone & "', Province = '" & province & "', City = '" & city & "', District = '" & district & "', Address = '" & address & "', IsDefault = " & isDefault & " WHERE AddressID = " & addressId & " AND UserID = " & Session("UserID")
        End If
            
        If ExecuteNonQuery(sql) Then
            Response.Write "<script>alert('地址保存成功'); location.href='addresses.asp';</script>"
        Else
            Response.Write "<script>alert('地址保存失败'); location.href='addresses.asp';</script>"
        End If
    ElseIf action = "delete" Then
        Dim deleteId
        deleteId = Request.Form("addressId")
        If Not IsNumeric(deleteId) Then
            Response.Write "<script>alert('无效的地址ID'); history.back();</script>"
            Response.End
        End If
        
        Dim deleteSql
        deleteSql = "DELETE FROM UserAddresses WHERE AddressID = " & deleteId & " AND UserID = " & Session("UserID")
        If ExecuteNonQuery(deleteSql) Then
            Response.Write "<script>alert('地址删除成功'); location.href='addresses.asp';</script>"
        Else
            Response.Write "<script>alert('地址删除失败'); location.href='addresses.asp';</script>"
        End If
    ElseIf action = "setDefault" Then
        Dim setAddressId
        setAddressId = Request.Form("addressId")
        If Not IsNumeric(setAddressId) Then
            Response.Write "<script>alert('无效的地址ID'); history.back();</script>"
            Response.End
        End If
        
        ' 先取消所有默认地址
        Call ExecuteNonQuery("UPDATE UserAddresses SET IsDefault = 0 WHERE UserID = " & Session("UserID"))
        
        ' 设置指定地址为默认地址
        Dim setDefaultSql
        setDefaultSql = "UPDATE UserAddresses SET IsDefault = 1 WHERE AddressID = " & setAddressId & " AND UserID = " & Session("UserID")
        If ExecuteNonQuery(setDefaultSql) Then
            Response.Write "<script>alert('默认地址设置成功'); location.href='addresses.asp';</script>"
        Else
            Response.Write "<script>alert('默认地址设置失败'); location.href='addresses.asp';</script>"
        End If
    End If
End If

Call OpenConnection()

' 确保CSRF令牌存在
Call EnsureCSRFToken()

Dim userId
userId = Session("UserID")
%>
<!--#include file="../includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/user/index.asp">个人中心</a>
        <span class="separator">/</span>
        <span>收货地址</span>
    </div>
</div>

<div class="container">
    <div class="user-center">
        <!-- 侧边栏 -->
        <aside class="user-sidebar">
            <div class="user-profile">
                <h3><%= HTMLEncode(Session("Username")) %></h3>
                <p><%= HTMLEncode(Session("Email")) %></p>
            </div>
            
            <nav class="user-nav">
                <a href="/user/index.asp"><i class="fas fa-home"></i> 个人中心</a>
                <a href="/user/orders.asp"><i class="fas fa-list"></i> 我的订单</a>
                <a href="/user/settings.asp"><i class="fas fa-user-edit"></i> 账户设置</a>
                <a href="/user/addresses.asp" class="active"><i class="fas fa-map-marker-alt"></i> 收货地址</a>
                <a href="/user/favorites.asp"><i class="fas fa-heart"></i> 我的收藏</a>
                <a href="/user/logout.asp"><i class="fas fa-sign-out-alt"></i> 退出登录</a>
            </nav>
        </aside>
        
        <!-- 主内容 -->
        <div class="user-main">
            <div class="user-card">
                <div class="card-header">
                    <h2 class="card-title"><i class="fas fa-map-marker-alt"></i> 收货地址</h2>
                    <button class="btn btn-primary" onclick="showAddressForm()">
                        <i class="fas fa-plus"></i> 新增收货地址
                    </button>
                </div>
                
                <div class="addresses-list">
                    <% 
                    ' 获取用户地址列表
                    Dim rsAddresses
                    Set rsAddresses = ExecuteQuery("SELECT * FROM UserAddresses WHERE UserID = " & userId & " ORDER BY IsDefault DESC, CreatedAt DESC")
                    If Not rsAddresses Is Nothing And Not rsAddresses.EOF Then
                        Do While Not rsAddresses.EOF
                    %>
                    <div class="address-item <%= IIF(rsAddresses("IsDefault") <> 0, "default", "") %>">
                        <div class="address-info">
                            <div class="address-name">
                                <%= HTMLEncode(rsAddresses("Consignee")) %> 
                                <% If rsAddresses("IsDefault") <> 0 Then %>
                                <span class="badge">默认</span>
                                <% End If %>
                            </div>
                            <div class="address-phone"><%= HTMLEncode(rsAddresses("Phone")) %></div>
                            <div class="address-detail">
                                <%= HTMLEncode(rsAddresses("Province")) %>
                                <%= HTMLEncode(rsAddresses("City")) %>
                                <%= HTMLEncode(rsAddresses("District")) %>
                                <%= HTMLEncode(rsAddresses("Address")) %>
                            </div>
                        </div>
                        <div class="address-actions">
                            <%
                            If rsAddresses("IsDefault") <> 0 Then
                                Response.Write "<button class=""btn btn-sm btn-outline disabled"">默认地址</button>"
                            Else
                                Response.Write "<button class=""btn btn-sm btn-outline"" onclick=""setDefaultAddress(" & rsAddresses("AddressID") & ")"">设为默认</button>"
                            End If
                            
                            Response.Write "<button class=""btn btn-sm btn-text"" onclick=""editAddress(" & rsAddresses("AddressID") & ")"">编辑</button>"
                            Response.Write "<button class=""btn btn-sm btn-text text-danger"" onclick=""deleteAddress(" & rsAddresses("AddressID") & ")"">删除</button>"
                            %>
                        </div>
                    </div>
                    <%
                        rsAddresses.MoveNext
                        Loop
                        rsAddresses.Close
                        Set rsAddresses = Nothing
                    Else
                    %>
                    <div class="address-empty">
                        <i class="fas fa-map-marker-alt"></i>
                        <h3>暂无收货地址</h3>
                        <p>请添加收货地址，方便商品配送</p>
                        <button class="btn btn-primary" onclick="showAddressForm()">立即添加</button>
                    </div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- 添加/编辑地址弹窗 -->
<div class="modal" id="addressModal">
    <div class="modal-content">
        <div class="modal-header">
            <h3 id="modalTitle">新增收货地址</h3>
            <span class="close" onclick="closeAddressForm()">&times;</span>
        </div>
        <div class="modal-body">
            <form id="addressForm" method="post" action="addresses.asp">
                            <%= GetCSRFTokenField() %>
                            <input type="hidden" id="formAction" name="action" value="add">
                            <input type="hidden" id="formAddressId" name="addressId" value="">
                <div class="form-row">
                    <div class="form-group">
                        <label for="consignee">收货人姓名 *</label>
                        <input type="text" id="consignee" name="consignee" required>
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

<script src="/js/area_data.js"></script>
<script>
// 省市区数据已移至外部文件 area_data.js
// addressData 和 districtData 都从 area_data.js 加载

function updateCities() {
    var provinceSelect = document.getElementById('province');
    var citySelect = document.getElementById('city');
    var districtSelect = document.getElementById('district');
    
    var selectedProvince = provinceSelect.value;
    
    // 清空城市和区县选项
    citySelect.innerHTML = '<option value="">请选择城市</option>';
    districtSelect.innerHTML = '<option value="">请选择区县</option>';
    
    if (selectedProvince && addressData[selectedProvince]) {
        var cities = addressData[selectedProvince];
        console.log('Found cities for province:', cities);
        
        for (var i = 0; i < cities.length; i++) {
            var option = document.createElement('option');
            option.value = cities[i];
            option.textContent = cities[i];
            citySelect.appendChild(option);
        }
    } else {
        console.log('No data found for province:', selectedProvince);
    }
}

function updateDistricts() {
    console.log('updateDistricts called, selected city:', document.getElementById('city').value);
    var provinceSelect = document.getElementById('province');
    var citySelect = document.getElementById('city');
    var districtSelect = document.getElementById('district');
    
    var selectedCity = citySelect.value;
    var selectedProvince = provinceSelect.value;
    
    // 清空区县选项
    districtSelect.innerHTML = '<option value="">请选择区县</option>';
    
    // 对于直辖市和自治区/特别行政区，使用省份名称作为键
    var searchKey = selectedCity;
    if (!districtData[searchKey] && (selectedProvince === '北京市' || selectedProvince === '上海市' || selectedProvince === '天津市' || selectedProvince === '重庆市' || selectedProvince === '新疆维吾尔自治区' || selectedProvince === '西藏自治区' || selectedProvince === '宁夏回族自治区' || selectedProvince === '广西壮族自治区' || selectedProvince === '内蒙古自治区' || selectedProvince === '香港特别行政区' || selectedProvince === '澳门特别行政区' || selectedProvince === '台湾省')) {
        searchKey = selectedProvince;
    }
    
    if (searchKey && districtData[searchKey]) {
        var districts = districtData[searchKey];
        console.log('Found districts for city/province:', districts);
        
        for (var i = 0; i < districts.length; i++) {
            var option = document.createElement('option');
            option.value = districts[i];
            option.textContent = districts[i];
            districtSelect.appendChild(option);
        }
    } else {
        console.log('No data found for city:', selectedCity, 'or province:', selectedProvince);
    }
}

function showAddressForm() {
    document.getElementById('modalTitle').textContent = '新增收货地址';
    document.getElementById('addressForm').reset();
    document.getElementById('formAction').value = 'add';
    document.getElementById('formAddressId').value = '';
    document.getElementById('addressModal').classList.add('show');
    document.body.style.overflow = 'hidden'; // 防止背景滚动
}

function editAddress(addressId) {
    document.getElementById('modalTitle').textContent = '编辑收货地址';
    document.getElementById('formAction').value = 'edit';
    document.getElementById('formAddressId').value = addressId;
    document.getElementById('addressModal').classList.add('show');
    document.body.style.overflow = 'hidden'; // 防止背景滚动
}

function closeAddressForm() {
    document.getElementById('addressModal').classList.remove('show');
    document.body.style.overflow = 'auto'; // 恢复背景滚动
}

function deleteAddress(addressId) {
    if (confirm('确定要删除这个地址吗？')) {
        // 创建一个隐藏表单来提交删除请求
        var form = document.createElement('form');
        form.method = 'POST';
        form.style.display = 'none';
        
        // 添加CSRF令牌
        var csrfInput = document.createElement('input');
        csrfInput.type = 'hidden';
        csrfInput.name = 'csrf_token';
        csrfInput.value = '<%= Session("CSRFToken") %>';
        form.appendChild(csrfInput);
        
        var actionInput = document.createElement('input');
        actionInput.type = 'hidden';
        actionInput.name = 'action';
        actionInput.value = 'delete';
        form.appendChild(actionInput);
        
        var addressIdInput = document.createElement('input');
        addressIdInput.type = 'hidden';
        addressIdInput.name = 'addressId';
        addressIdInput.value = addressId;
        form.appendChild(addressIdInput);
        
        document.body.appendChild(form);
        form.submit();
    }
}

function setDefaultAddress(addressId) {
    // 创建一个隐藏表单来提交设为默认请求
    var form = document.createElement('form');
    form.method = 'POST';
    form.style.display = 'none';
    
    // 添加CSRF令牌
    var csrfInput = document.createElement('input');
    csrfInput.type = 'hidden';
    csrfInput.name = 'csrf_token';
    csrfInput.value = '<%= Session("CSRFToken") %>';
    form.appendChild(csrfInput);
    
    var actionInput = document.createElement('input');
    actionInput.type = 'hidden';
    actionInput.name = 'action';
    actionInput.value = 'setDefault';
    form.appendChild(actionInput);
    
    var addressIdInput = document.createElement('input');
    addressIdInput.type = 'hidden';
    addressIdInput.name = 'addressId';
    addressIdInput.value = addressId;
    form.appendChild(addressIdInput);
    
    document.body.appendChild(form);
    form.submit();
}

// ESC键关闭模态框
document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        var modal = document.getElementById('addressModal');
        if (modal.classList.contains('show')) {
            modal.classList.remove('show');
            document.body.style.overflow = 'auto'; // 恢复背景滚动
        }
    }
});

// 点击模态框外部关闭
window.onclick = function(event) {
    var modal = document.getElementById('addressModal');
    if (event.target == modal) {
        modal.classList.remove('show');
        document.body.style.overflow = 'auto'; // 恢复背景滚动
    }
}
</script>

<!--#include file="../includes/footer.asp"-->
<%
Call CloseConnection()
%>