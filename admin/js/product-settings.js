/**
 * V14.6 产品设置 - 页面脚本
 * 从 product_settings.asp 提取
 * 依赖: admin-common.js (AdminCommon 对象)
 */

// 注意：minTopPercent / minMiddlePercent / minBasePercent 以及
// formulaData / recipeData 等服务端注入的变量仍由 product_settings.asp
// 通过内联 <script> 块输出，本文件依赖这些全局变量。

// 应用配方到表单
function applyFormula(formulaId) {
    if (!formulaId || !formulaData[formulaId]) {
        return;
    }

    // 先清空所有选择
    var allCheckboxes = document.querySelectorAll('input[name="noteCheckbox"]');
    for (var i = 0; i < allCheckboxes.length; i++) {
        allCheckboxes[i].checked = false;
        var percentInput = document.querySelector('input[name="notePercent_' + allCheckboxes[i].value + '"]');
        if (percentInput) {
            percentInput.value = '';
            percentInput.style.display = 'none';
        }
        allCheckboxes[i].closest('.checkbox-item').classList.remove('selected');
    }

    // 获取配方数据
    var notes = formulaData[formulaId];
    var selectedNoteIds = [];

    // 应用配方中的香调和百分比
    for (var i = 0; i < notes.length; i++) {
        var noteId = notes[i].noteId;
        var percentage = notes[i].percentage;

        var checkbox = document.querySelector('input[name="noteCheckbox"][value="' + noteId + '"]');
        if (checkbox) {
            checkbox.checked = true;
            checkbox.closest('.checkbox-item').classList.add('selected');
            selectedNoteIds.push(noteId);

            // 设置百分比
            var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
            if (percentInput) {
                percentInput.value = percentage;
                // 如果是KOL类型，显示百分比输入框
                var productType = document.getElementById('productType').value;
                if (productType === 'kol') {
                    percentInput.style.display = 'inline-block';
                }
            }
        }
    }

    // 更新选中的香调隐藏字段
    document.getElementById('selectedNotes').value = selectedNoteIds.join(',');

    // 更新配比统计
    var productType = document.getElementById('productType').value;
    if (productType === 'kol') {
        updateRatioSummary();
    }
}

// closest() polyfill for older browsers
if (!Element.prototype.closest) {
    Element.prototype.closest = function(s) {
        var el = this;
        do {
            if (el.matches(s)) return el;
            el = el.parentElement || el.parentNode;
        } while (el !== null && el.nodeType === 1);
        return null;
    };
}

// matches() polyfill for older browsers
if (!Element.prototype.matches) {
    Element.prototype.matches = Element.prototype.msMatchesSelector || Element.prototype.webkitMatchesSelector;
}

// 产品表单相关
function showAddProductForm() {
    document.getElementById('productModalTitle').textContent = '新增产品';
    document.getElementById('productFormAction').value = 'add_product';
    document.getElementById('editProductId').value = '';
    document.getElementById('productName').value = '';
    document.getElementById('productDescription').value = '';
    document.getElementById('productType').value = 'custom';
    document.getElementById('basePrice').value = '0';
    document.getElementById('baseIngredients').value = '';
    document.getElementById('reviewStatus').value = 'Pending';
    document.getElementById('imageURL_product').value = '/images/default-product.svg';
    document.getElementById('previewImg_product').style.display = 'none';
    document.getElementById('placeholder_product').style.display = 'flex';
    document.getElementById('fileInput_product').value = '';
    document.getElementById('productIsActive').value = '1';
    document.getElementById('kolId').value = '0';
    document.getElementById('engravable').checked = false;
    document.getElementById('engravingPrice').value = '0';
    document.getElementById('engravingPriceWrapper').style.display = 'none';

    // 清除所有复选框
    var checkboxes = document.querySelectorAll('input[type="checkbox"]');
    for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].checked = false;
        var checkboxItem = checkboxes[i].closest('.checkbox-item');
        if (checkboxItem) {
            checkboxItem.classList.remove('selected');
        }
    }
    // 隐藏所有瓶型价格输入框
    var bottlePriceWrappers = document.querySelectorAll('.bottle-price-wrapper');
    for (var i = 0; i < bottlePriceWrappers.length; i++) {
        bottlePriceWrappers[i].style.display = 'none';
    }
    document.getElementById('selectedNotes').value = '';
    document.getElementById('selectedVolumes').value = '';
    document.getElementById('selectedBottles').value = '';

    // 重置配方选择下拉框
    document.getElementById('formulaSelect').value = '';
    document.getElementById('recipeId').value = '';

    toggleProductFields();
    document.getElementById('productModal').style.display = 'block';
}

function showEditProductForm(button) {
    var id = button.getAttribute('data-id');
    var name = button.getAttribute('data-name');
    var desc = button.getAttribute('data-desc');
    var price = button.getAttribute('data-price');
    var type = button.getAttribute('data-type');
    var baseIng = button.getAttribute('data-baseing');
    var review = button.getAttribute('data-review');
    var active = button.getAttribute('data-active');
    var image = button.getAttribute('data-image');
    var kolId = button.getAttribute('data-kolid');
    var engravable = button.getAttribute('data-engravable');
    var engravingPrice = button.getAttribute('data-engravingprice');
    var recipeId = button.getAttribute('data-recipeid');
    var notesData = button.getAttribute('data-notes');
    var ratiosData = button.getAttribute('data-ratios');
    var volumesData = button.getAttribute('data-volumes');
    var bottlesData = button.getAttribute('data-bottles');

    document.getElementById('productModalTitle').textContent = '编辑产品';
    document.getElementById('productFormAction').value = 'edit_product';
    document.getElementById('editProductId').value = id;
    document.getElementById('productName').value = name;
    document.getElementById('productDescription').value = desc;
    document.getElementById('productType').value = type.toLowerCase();
    document.getElementById('basePrice').value = price;
    document.getElementById('baseIngredients').value = baseIng;
    // KOL类型强制待审核，其他类型使用传入的审核状态
    if (type.toLowerCase() === 'kol') {
        document.getElementById('reviewStatus').value = 'Pending';
    } else {
        document.getElementById('reviewStatus').value = review || 'Pending';
        document.getElementById('reviewStatusSelect').value = review || 'Pending';
    }
    var productImageVal = image || '/images/default-product.svg';
    document.getElementById('imageURL_product').value = productImageVal;
    if (productImageVal && productImageVal !== '/images/default-product.svg') {
        document.getElementById('previewImg_product').src = productImageVal;
        document.getElementById('previewImg_product').style.display = 'block';
        document.getElementById('placeholder_product').style.display = 'none';
    } else {
        document.getElementById('previewImg_product').style.display = 'none';
        document.getElementById('placeholder_product').style.display = 'flex';
    }
    document.getElementById('productIsActive').value = active || '1';
    document.getElementById('kolId').value = kolId || '0';
    document.getElementById('recipeId').value = recipeId || '';

    // 清除所有复选框
    var checkboxes = document.querySelectorAll('input[type="checkbox"]');
    for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].checked = false;
        if (checkboxes[i].closest('.checkbox-item')) {
            checkboxes[i].closest('.checkbox-item').classList.remove('selected');
        }
    }

    // 加载产品关联数据
    loadProductConfig(notesData, volumesData, type, ratiosData, bottlesData);

    // 设置刻字配置
    var engravableCheckbox = document.getElementById('engravable');
    var engravingPriceWrapper = document.getElementById('engravingPriceWrapper');
    var engravingPriceInput = document.getElementById('engravingPrice');
    if (engravableCheckbox) {
        engravableCheckbox.checked = (engravable === '1' || engravable === 'True' || engravable === '-1');
        if (engravingPriceWrapper) {
            engravingPriceWrapper.style.display = engravableCheckbox.checked ? 'block' : 'none';
        }
        if (engravingPriceInput) {
            engravingPriceInput.value = engravingPrice || '0';
        }
    }

    toggleProductFields();
    document.getElementById('productModal').style.display = 'block';
}

// 加载产品配置数据
function loadProductConfig(notesData, volumesData, productType, ratiosData, bottlesData) {
    // 重置配方选择下拉框（编辑时不清除已有选择，只是重置下拉框）
    document.getElementById('formulaSelect').value = '';

    // 解析配比数据为字典
    var ratiosDict = {};
    if (ratiosData) {
        var ratioPairs = ratiosData.split(',');
        for (var i = 0; i < ratioPairs.length; i++) {
            var pair = ratioPairs[i].trim();
            if (pair && pair.indexOf(':') > -1) {
                var parts = pair.split(':');
                ratiosDict[parts[0]] = parts[1];
            }
        }
    }

    // 设置香调
    if (notesData) {
        var noteIds = notesData.split(',');
        var selectedNotes = [];
        for (var i = 0; i < noteIds.length; i++) {
            var noteId = noteIds[i].trim();
            if (noteId) {
                var checkbox = document.querySelector('input[name="noteCheckbox"][value="' + noteId + '"]');
                if (checkbox) {
                    checkbox.checked = true;
                    checkbox.closest('.checkbox-item').classList.add('selected');
                    selectedNotes.push(noteId);
                    // 加载配比值
                    var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
                    if (percentInput && ratiosDict[noteId]) {
                        percentInput.value = ratiosDict[noteId];
                    }
                    // KOL类型时显示百分比输入框
                    if (percentInput && productType.toLowerCase() === 'kol') {
                        percentInput.style.display = 'inline-block';
                    }
                }
            }
        }
        document.getElementById('selectedNotes').value = selectedNotes.join(',');
    }

    // 设置容量
    if (volumesData) {
        var volIds = volumesData.split(',');
        var selectedVolumes = [];
        for (var i = 0; i < volIds.length; i++) {
            var volId = volIds[i].trim();
            if (volId) {
                var checkbox = document.querySelector('input[name="volumeCheckbox"][value="' + volId + '"]');
                if (checkbox) {
                    checkbox.checked = true;
                    checkbox.closest('.checkbox-item').classList.add('selected');
                    selectedVolumes.push(volId);
                }
            }
        }
        document.getElementById('selectedVolumes').value = selectedVolumes.join(',');
    }

    // 设置瓶型
    if (bottlesData) {
        try {
            // 将单引号替换为双引号以兼容JSON.parse
            var bottlesJson = bottlesData.replace(/'/g, '"');
            var bottlesArr = JSON.parse(bottlesJson);
            var selectedBottles = [];
            for (var i = 0; i < bottlesArr.length; i++) {
                var bottle = bottlesArr[i];
                if (bottle && bottle.bid) {
                    var checkbox = document.querySelector('input[name="bottleCheckbox"][value="' + bottle.bid + '"]');
                    if (checkbox) {
                        checkbox.checked = true;
                        checkbox.closest('.checkbox-item').classList.add('selected');
                        selectedBottles.push(bottle.bid);
                    }
                }
            }
            document.getElementById('selectedBottles').value = selectedBottles.join(',');
        } catch (e) {
            console.error('解析瓶型数据失败:', e);
        }
    }
}

// 切换香调百分比输入框显示/隐藏
function toggleNotePercentInput(checkbox) {
    var noteId = checkbox.value;
    var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
    if (percentInput) {
        percentInput.style.display = checkbox.checked ? 'inline-block' : 'none';
        if (!checkbox.checked) {
            percentInput.value = ''; // 取消勾选时清空值
        }
    }
    updateSelectedNotes();
}

function toggleProductFields() {
    var typeSelect = document.getElementById('productType');
    var selectedOption = typeSelect.options[typeSelect.selectedIndex];
    var typeCode = selectedOption.value.toLowerCase();
    var requiresReview = selectedOption.getAttribute('data-review') === 'True';
    var requiresRatio = selectedOption.getAttribute('data-ratio') === 'True';

    // 品牌定香(standard)类型提示横幅
    var fixedWarning = document.getElementById('fixedTypeWarning');
    var submitBtn = document.getElementById('submitProductBtn');
    var isBrand = (typeCode === 'standard');

    if (fixedWarning) {
        fixedWarning.style.display = isBrand ? 'block' : 'none';
    }
    if (submitBtn) {
        submitBtn.style.display = isBrand ? 'none' : 'inline-flex';
    }

    // standard类型显示基香成分字段
    document.getElementById('fixedFields').style.display = (typeCode === 'standard') ? 'block' : 'none';

    // KOL类型显示KOL选择字段
    document.getElementById('kolFields').style.display = (typeCode === 'kol') ? 'block' : 'none';

    // KOL类型强制待审核，不显示审核状态选择
    if (typeCode === 'kol') {
        document.getElementById('reviewStatus').value = 'Pending';
        document.getElementById('reviewFields').style.display = 'none';
    } else {
        // 其他需要审核的类型显示审核状态字段
        document.getElementById('reviewFields').style.display = requiresReview ? 'block' : 'none';
    }

    // Custom和KOL类型显示香调配置
    var isCustomOrKOL = (typeCode === 'custom' || typeCode === 'kol');
    document.getElementById('fragranceFields').style.display = isCustomOrKOL ? 'block' : 'none';

    // 配方导入区域仅KOL类型可见
    var formulaImportFields = document.getElementById('formulaImportFields');
    if (formulaImportFields) {
        formulaImportFields.style.display = (typeCode === 'kol') ? 'block' : 'none';
    }

    // Custom和KOL类型显示瓶型配置
    document.getElementById('bottleFields').style.display = isCustomOrKOL ? 'block' : 'none';

    // 所有类型都显示刻字配置
    document.getElementById('engravingFields').style.display = 'block';

    // KOL类型：根据复选框状态显示比例输入
    var percentInputs = document.querySelectorAll('.note-percent-input');
    for (var i = 0; i < percentInputs.length; i++) {
        var noteId = percentInputs[i].name.replace('notePercent_', '');
        var checkbox = document.querySelector('input[name="noteCheckbox"][value="' + noteId + '"]');
        if (typeCode === 'kol') {
            percentInputs[i].style.display = (checkbox && checkbox.checked) ? 'inline-block' : 'none';
        } else {
            percentInputs[i].style.display = 'none';
        }
    }

    // 显示/隐藏配比统计区域
    document.getElementById('ratioSummary').style.display = (typeCode === 'kol') ? 'block' : 'none';

    // 如果是KOL类型，更新配比统计
    if (typeCode === 'kol') {
        updateRatioSummary();
    }

    // 显示/隐藏关联配方字段（仅standard类型显示）
    var recipeFields = document.getElementById('recipeFields');
    var recipeSelect = document.getElementById('recipeId');
    var recipeLabel = document.getElementById('recipeLabel');
    var recipeHint = document.getElementById('recipeHint');
    if (recipeFields) {
        var showRecipe = (typeCode === 'standard');
        recipeFields.style.display = showRecipe ? 'block' : 'none';

        // 过滤配方选项（不区分大小写匹配）
        var options = recipeSelect.querySelectorAll('option[data-rtype]');
        for (var i = 0; i < options.length; i++) {
            var opt = options[i];
            if (opt.getAttribute('data-rtype').toLowerCase() === typeCode) {
                opt.style.display = '';
            } else {
                opt.style.display = 'none';
                // 如果当前选中的不是该类型的配方，清空选择
                if (recipeSelect.value === opt.value) {
                    recipeSelect.value = '';
                }
            }
        }

        if (typeCode === 'standard') {
            recipeLabel.innerHTML = '关联配方';
            recipeSelect.required = false;
            recipeHint.textContent = '可选：品牌定香产品不强制关联配方';
        } else if (typeCode === 'custom') {
            recipeLabel.innerHTML = '关联配方';
            recipeSelect.required = false;
            recipeHint.textContent = '可选：选择一个推荐配方';
        } else {
            recipeLabel.innerHTML = '关联配方';
            recipeSelect.required = false;
            recipeHint.textContent = '';
        }
    }
}

// 更新选中的香调
function updateSelectedNotes() {
    var checkboxes = document.querySelectorAll('input[name="noteCheckbox"]:checked');
    var selected = [];
    for (var i = 0; i < checkboxes.length; i++) {
        selected.push(checkboxes[i].value);
    }
    document.getElementById('selectedNotes').value = selected.join(',');

    // 更新配比统计
    var productType = document.getElementById('productType').value;
    if (productType === 'kol') {
        updateRatioSummary();
    }
}

// 配比统计和校验（使用容差0.01避免浮点数精度问题）
function updateRatioSummary() {
    var topTotal = 0, middleTotal = 0, baseTotal = 0;
    var checkboxes = document.querySelectorAll('input[name="noteCheckbox"]:checked');

    for (var i = 0; i < checkboxes.length; i++) {
        var noteType = checkboxes[i].getAttribute('data-type');
        var noteId = checkboxes[i].value;
        var percentInput = document.querySelector('input[name="notePercent_' + noteId + '"]');
        var percent = percentInput ? parseFloat(percentInput.value) || 0 : 0;

        if (noteType === 'top') {
            topTotal += percent;
        } else if (noteType === 'middle') {
            middleTotal += percent;
        } else if (noteType === 'base') {
            baseTotal += percent;
        }
    }

    var total = topTotal + middleTotal + baseTotal;

    // 更新显示 - 包含最小比例提示
    var ratioDetailText = '前调: ' + topTotal.toFixed(1) + '% (最低' + minTopPercent + '%) | ' +
                          '中调: ' + middleTotal.toFixed(1) + '% (最低' + minMiddlePercent + '%) | ' +
                          '后调: ' + baseTotal.toFixed(1) + '% (最低' + minBasePercent + '%)';
    document.getElementById('ratioDetail').textContent = ratioDetailText;
    var totalEl = document.getElementById('ratioTotal');
    totalEl.textContent = '总计: ' + total.toFixed(1) + '%';

    // 检查各项是否满足最小比例
    var topValid = topTotal >= minTopPercent - 0.01;
    var middleValid = middleTotal >= minMiddlePercent - 0.01;
    var baseValid = baseTotal >= minBasePercent - 0.01;
    var totalValid = Math.abs(total - 100) < 0.01;

    // 根据验证结果设置颜色
    if (totalValid && topValid && middleValid && baseValid) {
        totalEl.style.color = '#4caf50'; // 绿色 - 全部通过
    } else {
        totalEl.style.color = '#ff9800'; // 橙色 - 有错误
    }

    // 显示/隐藏错误提示
    var errorEl = document.getElementById('ratioError');
    var errorMsgs = [];

    if (!topValid) {
        errorMsgs.push('前调比例(' + topTotal.toFixed(1) + '%)不得低于最小值' + minTopPercent + '%');
    }
    if (!middleValid) {
        errorMsgs.push('中调比例(' + middleTotal.toFixed(1) + '%)不得低于最小值' + minMiddlePercent + '%');
    }
    if (!baseValid) {
        errorMsgs.push('后调比例(' + baseTotal.toFixed(1) + '%)不得低于最小值' + minBasePercent + '%');
    }
    if (!totalValid) {
        errorMsgs.push('配比总和必须等于100%，当前为' + total.toFixed(1) + '%');
    }

    if (errorMsgs.length > 0) {
        errorEl.innerHTML = errorMsgs.join('<br>');
        errorEl.style.display = 'block';
    } else {
        errorEl.style.display = 'none';
    }

    // 返回验证结果对象
    return {
        valid: totalValid && topValid && middleValid && baseValid,
        total: total,
        top: topTotal,
        middle: middleTotal,
        base: baseTotal,
        topValid: topValid,
        middleValid: middleValid,
        baseValid: baseValid,
        totalValid: totalValid
    };
}

// 更新选中的容量
function updateSelectedVolumes() {
    var checkboxes = document.querySelectorAll('input[name="volumeCheckbox"]:checked');
    var selected = [];
    for (var i = 0; i < checkboxes.length; i++) {
        selected.push(checkboxes[i].value);
    }
    document.getElementById('selectedVolumes').value = selected.join(',');
}

// 更新选中的瓶型
function updateSelectedBottles() {
    var checkboxes = document.querySelectorAll('input[name="bottleCheckbox"]:checked');
    var selected = [];
    for (var i = 0; i < checkboxes.length; i++) {
        selected.push(checkboxes[i].value);
    }
    document.getElementById('selectedBottles').value = selected.join(',');
}

// 绑定复选框事件
document.addEventListener('DOMContentLoaded', function() {
    // 香调复选框
    var noteCheckboxes = document.querySelectorAll('input[name="noteCheckbox"]');
    for (var i = 0; i < noteCheckboxes.length; i++) {
        noteCheckboxes[i].addEventListener('change', function() {
            updateSelectedNotes();
            // 切换选中样式
            if (this.checked) {
                this.closest('.checkbox-item').classList.add('selected');
            } else {
                this.closest('.checkbox-item').classList.remove('selected');
            }
        });
    }

    // 容量复选框
    var volumeCheckboxes = document.querySelectorAll('input[name="volumeCheckbox"]');
    for (var i = 0; i < volumeCheckboxes.length; i++) {
        volumeCheckboxes[i].addEventListener('change', function() {
            updateSelectedVolumes();
            // 切换选中样式
            if (this.checked) {
                this.closest('.checkbox-item').classList.add('selected');
            } else {
                this.closest('.checkbox-item').classList.remove('selected');
            }
        });
    }

    // 瓶型复选框
    var bottleCheckboxes = document.querySelectorAll('input[name="bottleCheckbox"]');
    for (var i = 0; i < bottleCheckboxes.length; i++) {
        bottleCheckboxes[i].addEventListener('change', function() {
            // 切换选中样式
            if (this.checked) {
                this.closest('.checkbox-item').classList.add('selected');
            } else {
                this.closest('.checkbox-item').classList.remove('selected');
            }
            // 更新选中列表
            updateSelectedBottles();
        });
    }

    // 配比输入框事件监听
    var percentInputs = document.querySelectorAll('.note-percent-input');
    for (var i = 0; i < percentInputs.length; i++) {
        percentInputs[i].addEventListener('input', function() {
            updateRatioSummary();
        });
    }

    // 表单提交前校验
    var productForm = document.getElementById('productForm');
    if (productForm) {
        productForm.addEventListener('submit', function(e) {
            var productType = document.getElementById('productType').value;
            if (productType === 'kol') {
                var result = updateRatioSummary();
                if (!result.valid) {
                    e.preventDefault();
                    var errorMsgs = [];
                    if (!result.topValid) {
                        errorMsgs.push('前调比例(' + result.top.toFixed(1) + '%)不得低于最小值' + minTopPercent + '%');
                    }
                    if (!result.middleValid) {
                        errorMsgs.push('中调比例(' + result.middle.toFixed(1) + '%)不得低于最小值' + minMiddlePercent + '%');
                    }
                    if (!result.baseValid) {
                        errorMsgs.push('后调比例(' + result.base.toFixed(1) + '%)不得低于最小值' + minBasePercent + '%');
                    }
                    if (!result.totalValid) {
                        errorMsgs.push('配比总和必须等于100%，当前为' + result.total.toFixed(1) + '%');
                    }
                    alert('配比校验失败：\n' + errorMsgs.join('\n'));
                    return false;
                }
            }
        });
    }

    // 刻字开关
    var engravableCheckbox = document.getElementById('engravable');
    if (engravableCheckbox) {
        engravableCheckbox.addEventListener('change', function() {
            document.getElementById('engravingPriceWrapper').style.display = this.checked ? 'block' : 'none';
        });
    }
});

function closeProductModal() {
    document.getElementById('productModal').style.display = 'none';
}

// 类型表单相关
function showEditTypeForm(button) {
    var id = button.getAttribute('data-id');
    var code = button.getAttribute('data-code');
    var display = button.getAttribute('data-display');
    var nav = button.getAttribute('data-nav');
    var desc = button.getAttribute('data-desc');
    var icon = button.getAttribute('data-icon');
    var review = button.getAttribute('data-review') === 'True';
    var ratio = button.getAttribute('data-ratio') === 'True';
    var order = button.getAttribute('data-order');
    var active = button.getAttribute('data-active') === 'True';

    document.getElementById('typeConfigId').value = id;
    document.getElementById('typeCodeDisplay').value = code;
    document.getElementById('typeDisplayName').value = display;
    document.getElementById('typeNavName').value = nav;
    document.getElementById('typeDescription').value = desc;
    document.getElementById('typeIcon').value = icon;
    document.getElementById('typeDisplayOrder').value = order;
    document.getElementById('typeRequiresReview').checked = review;
    document.getElementById('typeRequiresRatio').checked = ratio;
    document.getElementById('typeIsActive').checked = active;

    document.getElementById('typeModal').style.display = 'block';
}

function closeTypeModal() {
    document.getElementById('typeModal').style.display = 'none';
}

// 点击模态框外部关闭
window.onclick = function(event) {
    var productModal = document.getElementById('productModal');
    var typeModal = document.getElementById('typeModal');
    if (event.target == productModal) {
        productModal.style.display = 'none';
    }
    if (event.target == typeModal) {
        typeModal.style.display = 'none';
    }
}

// 图片压缩函数 - product
function compressImage_product(file, maxSizeKB, callback) {
    // SVG 不压缩，直接返回
    if (file.type === 'image/svg+xml') {
        callback(file, false);
        return;
    }
    var maxSize = maxSizeKB * 1024;
    // 文件已经足够小，直接返回
    if (file.size <= maxSize) {
        callback(file, false);
        return;
    }
    var reader = new FileReader();
    reader.onload = function(e) {
        var img = new Image();
        img.onload = function() {
            var canvas = document.createElement('canvas');
            var ctx = canvas.getContext('2d');
            var maxDim = 1200;
            var width = img.width;
            var height = img.height;
            if (width > maxDim || height > maxDim) {
                if (width > height) {
                    height = Math.round(height * maxDim / width);
                    width = maxDim;
                } else {
                    width = Math.round(width * maxDim / height);
                    height = maxDim;
                }
            }
            canvas.width = width;
            canvas.height = height;
            ctx.drawImage(img, 0, 0, width, height);
            var quality = 0.8;
            var tryCompress = function() {
                canvas.toBlob(function(blob) {
                    if (blob.size > maxSize && quality > 0.1) {
                        quality -= 0.1;
                        tryCompress();
                    } else {
                        var compressedFile = new File([blob], file.name.replace(/\.[^.]+$/, '.jpg'), {
                            type: 'image/jpeg',
                            lastModified: Date.now()
                        });
                        callback(compressedFile, true);
                    }
                }, 'image/jpeg', quality);
            };
            tryCompress();
        };
        img.src = e.target.result;
    };
    reader.readAsDataURL(file);
}

// 图片上传 - product
document.getElementById('fileInput_product').addEventListener('change', function(e) {
    var file = e.target.files[0];
    if (!file) return;
    var maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) { alert('文件大小不能超过5MB'); return; }
    var allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/svg+xml'];
    if (allowedTypes.indexOf(file.type) === -1) { alert('仅支持 JPG/PNG/GIF/WebP/SVG 格式'); return; }
    var reader = new FileReader();
    reader.onload = function(ev) {
        document.getElementById('previewImg_product').src = ev.target.result;
        document.getElementById('previewImg_product').style.display = 'block';
        document.getElementById('placeholder_product').style.display = 'none';
    };
    reader.readAsDataURL(file);
    compressImage_product(file, 180, function(fileToUpload, wasCompressed) {
        uploadImage_product(fileToUpload, 'product', wasCompressed);
    });
});

function uploadImage_product(file, uploadType, wasCompressed) {
    var formData = new FormData();
    formData.append('file', file);
    formData.append('type', uploadType);
    var csrfInput = document.querySelector('input[name="csrf_token"]');
    if (csrfInput) formData.append('csrf_token', csrfInput.value);
    var progressDiv = document.getElementById('uploadProgress_product');
    var progressBar = document.getElementById('progressBar_product');
    var progressText = document.getElementById('progressText_product');
    progressDiv.style.display = 'block';
    progressBar.style.width = '0%';
    progressText.textContent = '上传中...';
    var xhr = new XMLHttpRequest();
    xhr.upload.addEventListener('progress', function(e) {
        if (e.lengthComputable) {
            var pct = Math.round(e.loaded / e.total * 100);
            progressBar.style.width = pct + '%';
            progressText.textContent = pct + '%';
        }
    });
    xhr.addEventListener('load', function() {
        try {
            var resp = JSON.parse(xhr.responseText);
            if (resp.success) {
                document.getElementById('imageURL_product').value = resp.url;
                progressBar.style.width = '100%';
                progressText.textContent = wasCompressed ? '上传成功（图片已自动压缩）' : '上传成功';
                setTimeout(function() { progressDiv.style.display = 'none'; }, 2000);
            } else {
                alert('上传失败: ' + (resp.error || '未知错误'));
                progressDiv.style.display = 'none';
            }
        } catch(ex) {
            alert('上传响应解析失败');
            progressDiv.style.display = 'none';
        }
    });
    xhr.addEventListener('error', function() {
        alert('上传请求失败，请检查网络');
        progressDiv.style.display = 'none';
    });
    xhr.open('POST', '/api/upload.asp', true);
    xhr.send(formData);
}

(function() {
    var wrapper = document.getElementById('imagePreview_product').parentElement;
    wrapper.addEventListener('dragover', function(e) { e.preventDefault(); wrapper.classList.add('dragover'); });
    wrapper.addEventListener('dragleave', function() { wrapper.classList.remove('dragover'); });
    wrapper.addEventListener('drop', function(e) {
        e.preventDefault();
        wrapper.classList.remove('dragover');
        var file = e.dataTransfer.files[0];
        if (file) {
            document.getElementById('fileInput_product').files = e.dataTransfer.files;
            document.getElementById('fileInput_product').dispatchEvent(new Event('change'));
        }
    });
    document.getElementById('imagePreview_product').addEventListener('click', function() {
        document.getElementById('fileInput_product').click();
    });
})();

function toggleUrlInput_product() {
    var el = document.getElementById('urlInputWrapper_product');
    el.style.display = el.style.display === 'none' ? 'block' : 'none';
}

function applyManualUrl_product() {
    var url = document.getElementById('manualUrl_product').value.trim();
    if (url) {
        document.getElementById('imageURL_product').value = url;
        document.getElementById('previewImg_product').src = url;
        document.getElementById('previewImg_product').style.display = 'block';
        document.getElementById('placeholder_product').style.display = 'none';
        document.getElementById('urlInputWrapper_product').style.display = 'none';
    }
}
