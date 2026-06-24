/* ============================================
   V14.6 结算页 - 页面脚本
   从 checkout.asp 提取
   ============================================ */

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
        for (var i = 0; i < cities.length; i++) {
            var option = document.createElement('option');
            option.value = cities[i];
            option.textContent = cities[i];
            citySelect.appendChild(option);
        }
    }
}

function updateDistricts() {
    var provinceSelect = document.getElementById('province');
    var citySelect = document.getElementById('city');
    var districtSelect = document.getElementById('district');
    
    var selectedCity = citySelect.value;
    var selectedProvince = provinceSelect.value;
    
    // 清空区县选项
    districtSelect.innerHTML = '<option value="">请选择区县</option>';
    
    // 对于直辖市，使用省份名称作为键
    var searchKey = selectedCity;
    if (!districtData[searchKey] && (selectedProvince === '北京市' || selectedProvince === '上海市' || selectedProvince === '天津市' || selectedProvince === '重庆市')) {
        searchKey = selectedProvince;
    }
    
    if (searchKey && districtData[searchKey]) {
        var districts = districtData[searchKey];
        for (var i = 0; i < districts.length; i++) {
            var option = document.createElement('option');
            option.value = districts[i];
            option.textContent = districts[i];
            districtSelect.appendChild(option);
        }
    }
}

// 显示地址表单弹窗
function showAddressForm() {
    document.getElementById('modalTitle').textContent = '新增收货地址';
    document.getElementById('addressForm').reset();
    document.getElementById('formAction').value = 'add';
    document.getElementById('formAddressId').value = '';
    
    // 复制当前选中的支付方式到地址表单
    var selectedPayment = document.querySelector('input[name="payment_method"]:checked');
    if (selectedPayment) {
        document.getElementById('addressFormPaymentMethod').value = selectedPayment.value;
    }
    
    document.getElementById('addressModal').classList.add('show');
    document.body.style.overflow = 'hidden';
}

// 关闭地址表单弹窗
function closeAddressForm() {
    document.getElementById('addressModal').classList.remove('show');
    document.body.style.overflow = 'auto';
}

// 加载地址详情
function loadAddressDetails() {
    var selectElement = document.getElementById('selectedAddress');
    var selectedValue = selectElement.value;
    var selectedAddressDisplay = document.getElementById('selectedAddressDisplay');
    
    if (selectedValue === 'new') {
        showAddressForm();
        selectedAddressDisplay.style.display = 'none';
    } else if (selectedValue !== '') {
        selectedAddressDisplay.style.display = 'block';
    } else {
        selectedAddressDisplay.style.display = 'block';
    }
}

// 在页面加载时初始化
window.addEventListener('DOMContentLoaded', function() {
    // 检查URL参数中是否有预选地址
    var urlParams = new URLSearchParams(window.location.search);
    var preselectedAddress = urlParams.get('selected_address');
    
    if (preselectedAddress) {
        var selectElement = document.getElementById('selectedAddress');
        if (selectElement) {
            selectElement.value = preselectedAddress;
        }
    }
    
    // 检查URL参数中是否有预选支付方式
    var preselectedPaymentMethod = urlParams.get('payment_method');
    if (preselectedPaymentMethod) {
        var paymentRadio = document.querySelector('input[name="payment_method"][value="' + preselectedPaymentMethod + '"]');
        if (paymentRadio) {
            paymentRadio.checked = true;
        }
    }
    
    // 初始化页面
    loadAddressDetails();
    
    // 绑定支付表单提交事件
    var paymentForm = document.getElementById('paymentForm');
    if(paymentForm) {
        paymentForm.addEventListener('submit', function(e) {
            var selectedPayment = document.querySelector('input[name="payment_method"]:checked');
            var selectedAddress = document.getElementById('selectedAddress').value;
            
            if (!selectedPayment) {
                e.preventDefault();
                alert('请选择支付方式');
                return false;
            }
            
            if (!selectedAddress || selectedAddress === '') {
                e.preventDefault();
                alert('请选择收货地址');
                return false;
            }
            
            if (selectedAddress === 'new') {
                e.preventDefault();
                showAddressForm();
                alert('请先添加收货地址');
                return false;
            }
            
            // 显示加载提示
            var submitBtn = this.querySelector('button[type="submit"]');
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 处理中...';
            submitBtn.disabled = true;
        });
    }
});

// ESC键关闭模态框
document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        var modal = document.getElementById('addressModal');
        if (modal && modal.classList.contains('show')) {
            modal.classList.remove('show');
            document.body.style.overflow = 'auto';
        }
    }
});

// 点击模态框外部关闭
window.onclick = function(event) {
    var modal = document.getElementById('addressModal');
    if (modal && event.target == modal) {
        modal.classList.remove('show');
        document.body.style.overflow = 'auto';
    }
}

// 地址表单提交处理
var addressForm = document.getElementById('addressForm');
if(addressForm) {
    addressForm.onsubmit = function(e) {
        // 同步最新的支付方式选择
        var selectedPayment = document.querySelector('input[name="payment_method"]:checked');
        if (selectedPayment) {
            document.getElementById('addressFormPaymentMethod').value = selectedPayment.value;
        }
        
        // 验证表单数据
        var consignee = document.getElementById('consignee').value;
        var phone = document.getElementById('phone').value;
        var province = document.getElementById('province').value;
        var city = document.getElementById('city').value;
        var district = document.getElementById('district').value;
        var address = document.getElementById('address').value;
        
        if(!consignee || !phone || !province || !city || !district || !address) {
            alert('请填写完整的收货信息');
            return false;
        }
        
        // 显示加载提示
        var submitBtn = this.querySelector('button[type="submit"]');
        submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 保存中...';
        submitBtn.disabled = true;
        
        return true;
    };
}