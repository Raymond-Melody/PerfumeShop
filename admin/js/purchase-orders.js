/**
 * V14.6 采购订单 - 页面脚本
 * 从 purchase_orders.asp 提取
 * 依赖: admin-common.js (AdminCommon 对象)
 * 注意: baseNotesData 数组在服务端内联脚本中注入
 */

// 渲染基香网格
function renderBaseNoteGrid(filterText) {
    var grid = document.getElementById('baseNoteGrid');
    filterText = (filterText || '').toLowerCase();
    var html = '';
    var count = 0;
    for (var i = 0; i < baseNotesData.length; i++) {
        var bn = baseNotesData[i];
        if (filterText && bn.name.toLowerCase().indexOf(filterText) === -1 && bn.desc.toLowerCase().indexOf(filterText) === -1) continue;
        count++;
        var priceHtml = bn.price > 0 ? '¥' + bn.price.toFixed(4) + '/ml' : '<span class="no-price">未设单价</span>';
        var descHtml = bn.desc ? bn.desc : '暂无描述';
        html += '<div class="base-note-card" onclick="selectBaseNote(' + bn.id + ')">';
        html += '<div class="bn-name">' + bn.name + '</div>';
        html += '<div class="bn-desc">' + descHtml + '</div>';
        html += '<div class="bn-price">' + priceHtml + '</div>';
        html += '</div>';
    }
    if (count === 0) {
        html = '<div style="grid-column:1/-1;text-align:center;color:#666;padding:20px;">暂无匹配的基香</div>';
    }
    grid.innerHTML = html;
}

// 搜索过滤基香
function filterBaseNotes() {
    var searchText = document.getElementById('baseNoteSearch').value;
    renderBaseNoteGrid(searchText);
}

// 打开基香选择模态框
function openBaseNoteModal() {
    document.getElementById('baseNoteModal').classList.add('active');
    document.getElementById('baseNoteSearch').value = '';
    renderBaseNoteGrid('');
    document.getElementById('baseNoteSearch').focus();
}

// 关闭基香选择模态框
function closeBaseNoteModal() {
    document.getElementById('baseNoteModal').classList.remove('active');
}

// 点击选择基香后自动填充明细行
function selectBaseNote(baseNoteId) {
    var bn = null;
    for (var i = 0; i < baseNotesData.length; i++) {
        if (baseNotesData[i].id === baseNoteId) { bn = baseNotesData[i]; break; }
    }
    if (!bn) return;

    // 调用addRow填充数据
    addRowWithData(bn.name, 'BN-' + bn.id, '', 'ml', 100000, bn.price);
    closeBaseNoteModal();

    // 自动滚动到底部
    var tbody = document.getElementById('detailsBody');
    tbody.lastElementChild.scrollIntoView({behavior:'smooth'});
}

// 带数据添加明细行
function addRowWithData(itemName, itemCode, spec, unit, qty, price) {
    var tbody = document.getElementById('detailsBody');
    var rowCount = tbody.children.length + 1;

    var row = document.createElement('tr');
    row.innerHTML =
        '<td><input type="text" name="item_name_' + rowCount + '" value="' + escapeHtml(itemName) + '" required></td>' +
        '<td><input type="text" name="item_code_' + rowCount + '" value="' + escapeHtml(itemCode) + '"></td>' +
        '<td><input type="text" name="spec_' + rowCount + '" value="' + escapeHtml(spec) + '"></td>' +
        '<td><input type="text" name="unit_' + rowCount + '" value="' + escapeHtml(unit) + '"></td>' +
        '<td><input type="number" name="qty_' + rowCount + '" class="num-input qty" value="' + qty + '" min="0" step="0.01" onchange="calculateRow(this)"></td>' +
        '<td><input type="number" name="price_' + rowCount + '" class="num-input price" value="' + price.toFixed(4) + '" min="0" step="0.0001" onchange="calculateRow(this)"></td>' +
        '<td class="row-total">¥' + (qty * price).toFixed(2) + '</td>' +
        '<td><button type="button" class="btn btn-danger btn-sm" onclick="removeRow(this)"><i class="fas fa-trash"></i></button></td>';

    tbody.appendChild(row);
    renumberRows();
}

// HTML转义
function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&#39;');
}

// ========== 采购分类与采购类型联动 ==========
var categoryToOrderType = {
    'RAW': 'RawMaterial',
    'BASE': 'RawMaterial',
    'PACK': 'Packaging',
    'MARKET': 'Packaging',
    'BOTTLE': 'Bottle',
    'PRINTING': 'Printing',
    'SPRAYHEAD': 'SprayHead'
};

var orderTypeToCategory = {
    'RawMaterial': 'RAW',
    'Packaging': 'PACK',
    'Bottle': 'BOTTLE',
    'Printing': 'PRINTING',
    'SprayHead': 'SPRAYHEAD'
};

// 根据采购分类自动选择采购类型
function updateOrderTypeByCategory() {
    var catSelect = document.querySelector('select[name="category_code"]');
    var otSelect = document.querySelector('select[name="order_type"]');
    if (!catSelect || !otSelect) return;
    var catCode = catSelect.value;
    if (catCode && categoryToOrderType[catCode]) {
        otSelect.value = categoryToOrderType[catCode];
    }
    updateBaseNoteVisibility();
}

// 根据采购类型自动选择采购分类
function updateCategoryByOrderType() {
    var catSelect = document.querySelector('select[name="category_code"]');
    var otSelect = document.querySelector('select[name="order_type"]');
    if (!catSelect || !otSelect) return;
    var otCode = otSelect.value;
    if (otCode && orderTypeToCategory[otCode]) {
        catSelect.value = orderTypeToCategory[otCode];
    }
    updateBaseNoteVisibility();
}

// 控制基香按钮可见性（联动时实时切换）
function updateBaseNoteVisibility() {
    var bnSection = document.getElementById('baseNoteSection');
    if (!bnSection) return;
    var catSelect = document.querySelector('select[name="category_code"]');
    var otSelect = document.querySelector('select[name="order_type"]');
    if (!catSelect || !otSelect) return;
    var show = (catSelect.value === 'BASE') || (otSelect.value === 'RawMaterial');
    bnSection.style.display = show ? 'flex' : 'none';
}

// 点击遮罩关闭模态框
document.addEventListener('DOMContentLoaded', function() {
    var modal = document.getElementById('baseNoteModal');
    if (modal) {
        modal.addEventListener('click', function(e) {
            if (e.target === modal) closeBaseNoteModal();
        });
    }
    // ESC关闭模态框
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') closeBaseNoteModal();
    });
});

// 添加明细行
function addRow() {
    var tbody = document.getElementById('detailsBody');
    var rowCount = tbody.children.length + 1;

    var row = document.createElement('tr');
    row.innerHTML =
        '<td><input type="text" name="item_name_' + rowCount + '" required></td>' +
        '<td><input type="text" name="item_code_' + rowCount + '"></td>' +
        '<td><input type="text" name="spec_' + rowCount + '"></td>' +
        '<td><input type="text" name="unit_' + rowCount + '"></td>' +
        '<td><input type="number" name="qty_' + rowCount + '" class="num-input qty" value="0" min="0" step="0.01" onchange="calculateRow(this)"></td>' +
        '<td><input type="number" name="price_' + rowCount + '" class="num-input price" value="0" min="0" step="0.01" onchange="calculateRow(this)"></td>' +
        '<td class="row-total">¥0.00</td>' +
        '<td><button type="button" class="btn btn-danger btn-sm" onclick="removeRow(this)"><i class="fas fa-trash"></i></button></td>';

    tbody.appendChild(row);
    renumberRows();
}

// 删除明细行
function removeRow(btn) {
    var tbody = document.getElementById('detailsBody');
    if (tbody.children.length <= 1) {
        alert('至少保留一行明细');
        return;
    }
    btn.closest('tr').remove();
    renumberRows();
}

// 重新编号行
function renumberRows() {
    var tbody = document.getElementById('detailsBody');
    var rows = tbody.children;

    for (var i = 0; i < rows.length; i++) {
        var rowNum = i + 1;
        var inputs = rows[i].querySelectorAll('input');
        inputs[0].name = 'item_name_' + rowNum;
        inputs[1].name = 'item_code_' + rowNum;
        inputs[2].name = 'spec_' + rowNum;
        inputs[3].name = 'unit_' + rowNum;
        inputs[4].name = 'qty_' + rowNum;
        inputs[5].name = 'price_' + rowNum;
    }

    document.getElementById('itemCount').value = rows.length;
}

// 计算行小计
function calculateRow(input) {
    var row = input.closest('tr');
    var qty = parseFloat(row.querySelector('.qty').value) || 0;
    var price = parseFloat(row.querySelector('.price').value) || 0;
    var total = qty * price;
    row.querySelector('.row-total').textContent = '¥' + total.toFixed(2);
}

// 页面加载时确保至少有一行
window.onload = function() {
    var tbody = document.getElementById('detailsBody');
    if (tbody && tbody.children.length === 0) {
        addRow();
    }
};

// ========== V11: 历史产品快速选择 ==========
var historyDataCache = {};

function openHistoryModal() {
    document.getElementById('historyModal').classList.add('active');
    document.getElementById('historySearch').value = '';
    loadHistoryData();
}

function closeHistoryModal() {
    document.getElementById('historyModal').classList.remove('active');
}

function getCurrentOrderType() {
    var otSelect = document.querySelector('select[name="order_type"]');
    return otSelect ? otSelect.value : 'RawMaterial';
}

function loadHistoryData() {
    var grid = document.getElementById('historyGrid');
    grid.innerHTML = '<div style="text-align:center;color:#666;padding:20px;"><i class="fas fa-spinner fa-pulse"></i> 加载中...</div>';

    var orderType = getCurrentOrderType();
    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'ajax_product_history.asp?ordertype=' + encodeURIComponent(orderType) + '&limit=100', true);
    xhr.onload = function() {
        if (xhr.status === 200) {
            try {
                var data = JSON.parse(xhr.responseText);
                historyDataCache[orderType] = data;
                renderHistoryGrid(data);
            } catch(e) {
                grid.innerHTML = '<div style="text-align:center;color:#F44336;padding:20px;">加载失败</div>';
            }
        }
    };
    xhr.onerror = function() {
        grid.innerHTML = '<div style="text-align:center;color:#F44336;padding:20px;">网络错误</div>';
    };
    xhr.send();
}

function searchHistory() {
    var orderType = getCurrentOrderType();
    var data = historyDataCache[orderType] || [];
    var searchText = (document.getElementById('historySearch').value || '').toLowerCase();
    var supplierFilter = document.getElementById('historySupplierFilter').value;

    if (data.length === 0) {
        loadHistoryData();
        return;
    }

    var filtered = data.filter(function(item) {
        if (searchText && item.itemname.toLowerCase().indexOf(searchText) === -1 && item.itemcode.toLowerCase().indexOf(searchText) === -1) return false;
        if (supplierFilter && item.lastsupplierid != supplierFilter) return false;
        return true;
    });
    renderHistoryGrid(filtered);
}

function renderHistoryGrid(items) {
    var grid = document.getElementById('historyGrid');
    if (!items || items.length === 0) {
        grid.innerHTML = '<div style="text-align:center;color:#666;padding:20px;">暂无匹配的历史采购记录</div>';
        return;
    }

    // Build supplier filter options
    var suppliers = {};
    items.forEach(function(item) {
        if (item.lastsupplier && item.lastsupplierid > 0) {
            suppliers[item.lastsupplierid] = item.lastsupplier;
        }
    });
    var supplierHtml = '<option value="">全部供应商</option>';
    for (var sid in suppliers) {
        supplierHtml += '<option value="' + sid + '">' + escapeHtml(suppliers[sid]) + '</option>';
    }
    document.getElementById('historySupplierFilter').innerHTML = supplierHtml;

    var html = '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<thead><tr style="background:rgba(255,255,255,0.03);"><th style="padding:8px;text-align:left;color:#aaa;">产品名称</th><th style="padding:8px;text-align:left;color:#aaa;">编码</th><th style="padding:8px;text-align:left;color:#aaa;">规格</th><th style="padding:8px;text-align:left;color:#aaa;">单位</th><th style="padding:8px;text-align:right;color:#aaa;">最近价格</th><th style="padding:8px;text-align:left;color:#aaa;">供应商</th><th style="padding:8px;text-align:center;color:#aaa;">采购次数</th><th style="padding:8px;"></th></tr></thead><tbody>';

    items.forEach(function(item) {
        html += '<tr style="border-bottom:1px solid rgba(255,255,255,0.04);cursor:pointer;" onmouseover="this.style.background=\'rgba(255,152,0,0.08)\'" onmouseout="this.style.background=\'\'">';
        html += '<td style="padding:8px;"><strong>' + escapeHtml(item.itemname) + '</strong></td>';
        html += '<td style="padding:8px;color:#888;">' + escapeHtml(item.itemcode) + '</td>';
        html += '<td style="padding:8px;color:#888;">' + escapeHtml(item.spec) + '</td>';
        html += '<td style="padding:8px;color:#888;">' + escapeHtml(item.unit) + '</td>';
        html += '<td style="padding:8px;text-align:right;color:#4CAF50;">¥' + parseFloat(item.lastprice).toFixed(4) + '</td>';
        html += '<td style="padding:8px;">' + escapeHtml(item.lastsupplier) + '</td>';
        html += '<td style="padding:8px;text-align:center;">' + item.purchasecount + '次</td>';
        html += '<td style="padding:8px;"><button type="button" class="btn-select-base" onclick="selectHistoryItem(\'' + escapeHtml(item.itemname) + '\',\'' + escapeHtml(item.itemcode) + '\',\'' + escapeHtml(item.spec) + '\',\'' + escapeHtml(item.unit) + '\',' + parseFloat(item.lastprice).toFixed(4) + ')">选择</button></td>';
        html += '</tr>';
    });
    html += '</tbody></table>';
    grid.innerHTML = html;
}

function selectHistoryItem(itemName, itemCode, spec, unit, price) {
    addRowWithData(itemName, itemCode, spec, unit, 0, price);
    closeHistoryModal();
}

// 点击遮罩关闭
document.addEventListener('DOMContentLoaded', function() {
    var hModal = document.getElementById('historyModal');
    if (hModal) {
        hModal.addEventListener('click', function(e) {
            if (e.target === hModal) closeHistoryModal();
        });
    }

    // V11: 退出确认 - 未保存内容时提示
    var formSection = document.querySelector('.form-section');
    if (formSection) {
        window.addEventListener('beforeunload', function(e) {
            var tbody = document.getElementById('detailsBody');
            if (tbody && tbody.children.length > 0) {
                var hasContent = false;
                var inputs = tbody.querySelectorAll('input[name^="item_name_"]');
                inputs.forEach(function(inp) {
                    if (inp.value.trim() !== '') hasContent = true;
                });
                if (hasContent) {
                    e.preventDefault();
                    e.returnValue = '您有未保存的采购明细，确定离开吗？';
                    return e.returnValue;
                }
            }
        });
    }
});

// ========== V11: 批量操作 ==========
function toggleSelectAll(cb) {
    var checks = document.querySelectorAll('.row-check');
    checks.forEach(function(c) { c.checked = cb.checked; });
    updateBatchCount();
}

function updateBatchCount() {
    var checked = document.querySelectorAll('.row-check:checked');
    var countEl = document.getElementById('batchCount');
    if (checked.length > 0) {
        countEl.textContent = '已选择 ' + checked.length + ' 个订单';
        countEl.style.color = '#FF9800';
    } else {
        countEl.textContent = '';
    }
}

function batchAction(action) {
    var checked = document.querySelectorAll('.row-check:checked');
    if (checked.length === 0) {
        alert('请至少选择一个订单');
        return;
    }

    var ids = [];
    checked.forEach(function(c) { ids.push(c.value); });

    var actionLabel = action === 'approve' ? '审批' : (action === 'order' ? '下单' : '收货');
    if (!confirm('确定要复制此订单吗？将创建一个新的草稿订单。')) return;

    // Create and submit form
    var form = document.createElement('form');
    form.method = 'POST';
    form.style.display = 'none';

    var inputAction = document.createElement('input');
    inputAction.name = 'batch_action';
    inputAction.value = action;
    form.appendChild(inputAction);

    var inputIds = document.createElement('input');
    inputIds.name = 'batch_ids';
    inputIds.value = ids.join(',');
    form.appendChild(inputIds);

    document.body.appendChild(form);
    form.submit();
}

function copyOrderFromView(orderId) {
    if (!confirm('确定要复制此订单吗？将创建一个新的草稿订单。')) return;
    var form = document.createElement('form');
    form.method = 'POST';
    form.style.display = 'none';
    var input1 = document.createElement('input');
    input1.name = 'action';
    input1.value = 'copy';
    form.appendChild(input1);
    var input2 = document.createElement('input');
    input2.name = 'purchase_id';
    input2.value = orderId;
    form.appendChild(input2);
    document.body.appendChild(form);
    form.submit();
}
