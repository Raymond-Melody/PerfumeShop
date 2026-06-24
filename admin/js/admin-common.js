/**
 * V14.6 管理后台公共脚本
 * 从 product_settings.asp 和 purchase_orders.asp 提取
 * 使用方式: <script src="/admin/js/admin-common.js"></script>
 */

(function(window, document) {
    'use strict';

    // ========== 1. Polyfill ==========

    // closest() polyfill
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

    // matches() polyfill
    if (!Element.prototype.matches) {
        Element.prototype.matches = Element.prototype.msMatchesSelector || Element.prototype.webkitMatchesSelector;
    }

    // ========== 2. HTML转义 ==========
    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/'/g, '&#39;');
    }

    // ========== 3. 模态框管理 ==========

    /**
     * 打开模态框 (admin-modal 模式)
     * @param {string} modalId - 模态框元素ID
     */
    function openModal(modalId) {
        var modal = document.getElementById(modalId);
        if (modal) {
            modal.style.display = 'block';
            document.body.style.overflow = 'hidden';
        }
    }

    /**
     * 关闭模态框 (admin-modal 模式)
     * @param {string} modalId - 模态框元素ID
     */
    function closeModal(modalId) {
        var modal = document.getElementById(modalId);
        if (modal) {
            modal.style.display = 'none';
            document.body.style.overflow = '';
        }
    }

    /**
     * 打开模态框 (modal-overlay 模式)
     * @param {string} overlayId - overlay元素ID
     */
    function openOverlay(overlayId) {
        var overlay = document.getElementById(overlayId);
        if (overlay) {
            overlay.classList.add('active');
            document.body.style.overflow = 'hidden';
        }
    }

    /**
     * 关闭模态框 (modal-overlay 模式)
     * @param {string} overlayId - overlay元素ID
     */
    function closeOverlay(overlayId) {
        var overlay = document.getElementById(overlayId);
        if (overlay) {
            overlay.classList.remove('active');
            document.body.style.overflow = '';
        }
    }

    /**
     * 绑定模态框外部点击关闭和ESC关闭
     * @param {string} modalId - 模态框元素ID
     * @param {string} mode - 'modal' (admin-modal) 或 'overlay' (modal-overlay)
     */
    function bindModalClose(modalId, mode) {
        var modal = document.getElementById(modalId);
        if (!modal) return;

        // 点击遮罩关闭
        modal.addEventListener('click', function(e) {
            if (e.target === modal) {
                if (mode === 'overlay') {
                    closeOverlay(modalId);
                } else {
                    closeModal(modalId);
                }
            }
        });
    }

    // ESC键关闭所有模态框
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            // 关闭 admin-modal 类型
            var modals = document.querySelectorAll('.admin-modal');
            for (var i = 0; i < modals.length; i++) {
                if (modals[i].style.display === 'block') {
                    modals[i].style.display = 'none';
                    document.body.style.overflow = '';
                }
            }
            // 关闭 modal-overlay 类型
            var overlays = document.querySelectorAll('.modal-overlay.active');
            for (var j = 0; j < overlays.length; j++) {
                overlays[j].classList.remove('active');
                document.body.style.overflow = '';
            }
        }
    });

    // ========== 4. 图片压缩 ==========

    /**
     * 压缩图片
     * @param {File} file - 图片文件
     * @param {number} maxSizeKB - 最大文件大小(KB)
     * @param {Function} callback - 回调函数(file, wasCompressed)
     * @param {number} [maxDim=1200] - 最大尺寸(px)
     */
    function compressImage(file, maxSizeKB, callback, maxDim) {
        maxDim = maxDim || 1200;

        // SVG 不压缩
        if (file.type === 'image/svg+xml') {
            callback(file, false);
            return;
        }

        var maxSize = maxSizeKB * 1024;

        // 文件已经足够小
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
                var width = img.width;
                var height = img.height;

                // 缩放
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
                            var compressedFile = new File(
                                [blob],
                                file.name.replace(/\.[^.]+$/, '.jpg'),
                                { type: 'image/jpeg', lastModified: Date.now() }
                            );
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

    // ========== 5. 图片上传 ==========

    /**
     * 通用图片上传
     * @param {File} file - 要上传的文件
     * @param {string} uploadType - 上传类型
     * @param {boolean} wasCompressed - 是否已压缩
     * @param {object} options - 配置项
     * @param {string} options.imageURLId - 存储URL的hidden input ID
     * @param {string} options.previewImgId - 预览img元素ID
     * @param {string} options.placeholderId - 占位符元素ID
     * @param {string} options.progressBarId - 进度条元素ID
     * @param {string} options.progressDivId - 进度容器元素ID
     * @param {string} options.progressTextId - 进度文本元素ID
     * @param {Function} [options.onSuccess] - 上传成功回调
     */
    function uploadImage(file, uploadType, wasCompressed, options) {
        var formData = new FormData();
        formData.append('file', file);
        formData.append('type', uploadType);

        // 添加CSRF令牌
        var csrfInput = document.querySelector('input[name="csrf_token"]');
        if (csrfInput) {
            formData.append('csrf_token', csrfInput.value);
        } else if (typeof window.csrfToken !== 'undefined') {
            formData.append('csrf_token', window.csrfToken);
        }

        var progressDiv = document.getElementById(options.progressDivId);
        var progressBar = document.getElementById(options.progressBarId);
        var progressText = document.getElementById(options.progressTextId);

        if (progressDiv) progressDiv.style.display = 'block';
        if (progressBar) progressBar.style.width = '0%';
        if (progressText) progressText.textContent = '上传中...';

        var xhr = new XMLHttpRequest();

        xhr.upload.addEventListener('progress', function(e) {
            if (e.lengthComputable) {
                var pct = Math.round(e.loaded / e.total * 100);
                if (progressBar) progressBar.style.width = pct + '%';
                if (progressText) progressText.textContent = pct + '%';
            }
        });

        xhr.addEventListener('load', function() {
            try {
                var resp = JSON.parse(xhr.responseText);
                if (resp.success) {
                    // 更新URL字段
                    if (options.imageURLId) {
                        document.getElementById(options.imageURLId).value = resp.url;
                    }
                    if (progressBar) progressBar.style.width = '100%';
                    if (progressText) {
                        progressText.textContent = wasCompressed ? '上传成功（图片已自动压缩）' : '上传成功';
                    }
                    setTimeout(function() {
                        if (progressDiv) progressDiv.style.display = 'none';
                    }, 2000);
                    if (options.onSuccess) options.onSuccess(resp.url);
                } else {
                    alert('上传失败: ' + (resp.error || '未知错误'));
                    if (progressDiv) progressDiv.style.display = 'none';
                }
            } catch (ex) {
                alert('上传响应解析失败');
                if (progressDiv) progressDiv.style.display = 'none';
            }
        });

        xhr.addEventListener('error', function() {
            alert('上传请求失败，请检查网络');
            if (progressDiv) progressDiv.style.display = 'none';
        });

        xhr.open('POST', '/api/upload.asp', true);
        xhr.send(formData);
    }

    // ========== 6. 拖拽上传绑定 ==========

    /**
     * 为图片上传区域绑定拖拽事件
     * @param {string} wrapperSelector - 拖拽区域的CSS选择器或元素ID
     * @param {string} fileInputId - file input元素ID
     */
    function bindDragDrop(wrapperSelector, fileInputId) {
        var wrapper;
        if (wrapperSelector.charAt(0) === '#') {
            wrapper = document.getElementById(wrapperSelector.substring(1));
        } else {
            wrapper = document.querySelector(wrapperSelector);
        }
        var fileInput = document.getElementById(fileInputId);

        if (!wrapper || !fileInput) return;

        wrapper.addEventListener('dragover', function(e) {
            e.preventDefault();
            wrapper.classList.add('dragover');
        });
        wrapper.addEventListener('dragleave', function() {
            wrapper.classList.remove('dragover');
        });
        wrapper.addEventListener('drop', function(e) {
            e.preventDefault();
            wrapper.classList.remove('dragover');
            var file = e.dataTransfer.files[0];
            if (file) {
                fileInput.files = e.dataTransfer.files;
                fileInput.dispatchEvent(new Event('change'));
            }
        });
    }

    // ========== 7. 复选框网格管理 ==========

    /**
     * 更新选中复选框的值到隐藏字段
     * @param {string} checkboxName - 复选框name属性
     * @param {string} hiddenFieldId - 隐藏字段ID
     */
    function updateCheckedValues(checkboxName, hiddenFieldId) {
        var checkboxes = document.querySelectorAll('input[name="' + checkboxName + '"]:checked');
        var selected = [];
        for (var i = 0; i < checkboxes.length; i++) {
            selected.push(checkboxes[i].value);
        }
        var hiddenField = document.getElementById(hiddenFieldId);
        if (hiddenField) {
            hiddenField.value = selected.join(',');
        }
    }

    /**
     * 绑定复选框选中样式切换
     * @param {string} checkboxName - 复选框name属性
     * @param {Function} [onChange] - 变更回调
     */
    function bindCheckboxStyle(checkboxName, onChange) {
        var checkboxes = document.querySelectorAll('input[name="' + checkboxName + '"]');
        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].addEventListener('change', function() {
                var item = this.closest('.checkbox-item');
                if (item) {
                    if (this.checked) {
                        item.classList.add('selected');
                    } else {
                        item.classList.remove('selected');
                    }
                }
                if (onChange) onChange(this);
            });
        }
    }

    // ========== 8. 文件验证 ==========

    /**
     * 验证图片文件
     * @param {File} file - 文件对象
     * @param {number} [maxSizeMB=5] - 最大文件大小(MB)
     * @returns {boolean}
     */
    function validateImageFile(file, maxSizeMB) {
        maxSizeMB = maxSizeMB || 5;
        if (!file) return false;

        var maxSize = maxSizeMB * 1024 * 1024;
        if (file.size > maxSize) {
            alert('文件大小不能超过' + maxSizeMB + 'MB');
            return false;
        }

        var allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/svg+xml'];
        if (allowedTypes.indexOf(file.type) === -1) {
            alert('仅支持 JPG/PNG/GIF/WebP/SVG 格式');
            return false;
        }

        return true;
    }

    // ========== 9. 预览图片 ==========

    /**
     * 设置图片预览
     * @param {string} url - 图片URL
     * @param {string} previewImgId - 预览img元素ID
     * @param {string} placeholderId - 占位符元素ID
     */
    function setImagePreview(url, previewImgId, placeholderId) {
        var previewImg = document.getElementById(previewImgId);
        var placeholder = document.getElementById(placeholderId);

        if (url && url !== '/images/default-product.svg') {
            if (previewImg) {
                previewImg.src = url;
                previewImg.style.display = 'block';
            }
            if (placeholder) {
                placeholder.style.display = 'none';
            }
        } else {
            if (previewImg) previewImg.style.display = 'none';
            if (placeholder) placeholder.style.display = 'flex';
        }
    }

    // ========== 10. 数字格式化 ==========

    /**
     * 格式化价格
     * @param {number} num - 数字
     * @param {number} [decimals=2] - 小数位数
     * @returns {string}
     */
    function formatPrice(num, decimals) {
        decimals = decimals || 2;
        num = parseFloat(num) || 0;
        return '¥' + num.toFixed(decimals);
    }

    // ========== 暴露公共API ==========
    window.AdminCommon = {
        escapeHtml: escapeHtml,
        openModal: openModal,
        closeModal: closeModal,
        openOverlay: openOverlay,
        closeOverlay: closeOverlay,
        bindModalClose: bindModalClose,
        compressImage: compressImage,
        uploadImage: uploadImage,
        bindDragDrop: bindDragDrop,
        updateCheckedValues: updateCheckedValues,
        bindCheckboxStyle: bindCheckboxStyle,
        validateImageFile: validateImageFile,
        setImagePreview: setImagePreview,
        formatPrice: formatPrice
    };

})(window, document);
