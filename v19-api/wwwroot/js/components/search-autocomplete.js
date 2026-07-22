/**
 * V18.0 <search-autocomplete> Web Component
 * 搜索自动补全组件，支持 debounce 输入 + 下拉建议
 *
 * 属性:
 *   data-api       - 搜索建议 API 端点 (默认 /api/search_suggestions.asp)
 *   data-debounce  - 防抖延迟 ms (默认 300)
 *   data-min-chars - 最小触发字符数 (默认 2)
 *   data-max-results - 最大结果数 (默认 8)
 *   data-placeholder - 输入框占位文字
 *   data-name      - 表单字段名 (默认 "keyword")
 *   data-value     - 初始值
 *   data-show-icon - 是否显示搜索图标 (默认 true)
 *
 * 事件:
 *   search - 用户提交搜索，detail: { keyword: string }
 *   select - 用户选择建议项，detail: { item: object }
 *
 * 用法:
 *   <search-autocomplete data-placeholder="搜索香水..."
 *       data-api="/api/search_suggestions.asp"></search-autocomplete>
 */
class SearchAutocomplete extends HTMLElement {
    constructor() {
        super();
        this.attachShadow({ mode: 'open' });
        this._debounceTimer = null;
        this._abortController = null;
        this._selectedIndex = -1;
        this._results = [];
        this._isOpen = false;
    }

    static get observedAttributes() {
        return ['data-value', 'data-placeholder'];
    }

    connectedCallback() {
        this.render();
        this.setupEvents();
    }

    attributeChangedCallback(name, oldVal, newVal) {
        if (oldVal !== newVal && name === 'data-value') {
            const input = this.shadowRoot.querySelector('input');
            if (input && input.value !== newVal) {
                input.value = newVal || '';
            }
        }
    }

    get value() {
        const input = this.shadowRoot.querySelector('input');
        return input ? input.value : '';
    }

    set value(val) {
        const input = this.shadowRoot.querySelector('input');
        if (input) input.value = val || '';
    }

    setupEvents() {
        const input = this.shadowRoot.querySelector('input');
        const dropdown = this.shadowRoot.querySelector('.autocomplete-dropdown');
        const form = this.shadowRoot.querySelector('form');

        if (!input) return;

        // 输入事件
        input.addEventListener('input', (e) => {
            const query = e.target.value.trim();
            this._selectedIndex = -1;
            
            if (query.length >= this.minChars) {
                this.debounceSearch(query);
            } else {
                this.closeDropdown();
            }
        });

        // 键盘导航
        input.addEventListener('keydown', (e) => {
            if (!this._isOpen || this._results.length === 0) return;
            
            const items = this.shadowRoot.querySelectorAll('.suggestion-item');
            
            switch (e.key) {
                case 'ArrowDown':
                    e.preventDefault();
                    this._selectedIndex = Math.min(this._selectedIndex + 1, items.length - 1);
                    this.updateSelection(items);
                    break;
                case 'ArrowUp':
                    e.preventDefault();
                    this._selectedIndex = Math.max(this._selectedIndex - 1, -1);
                    this.updateSelection(items);
                    break;
                case 'Enter':
                    if (this._selectedIndex >= 0 && this._selectedIndex < this._results.length) {
                        e.preventDefault();
                        this.selectItem(this._results[this._selectedIndex]);
                    }
                    break;
                case 'Escape':
                    this.closeDropdown();
                    break;
            }
        });

        // 点击外部关闭
        document.addEventListener('click', (e) => {
            if (!this.contains(e.target)) {
                this.closeDropdown();
            }
        });

        // 表单提交
        if (form) {
            form.addEventListener('submit', (e) => {
                e.preventDefault();
                const keyword = input.value.trim();
                if (keyword) {
                    this.dispatchEvent(new CustomEvent('search', {
                        detail: { keyword },
                        bubbles: true,
                        composed: true
                    }));
                    // 默认行为: 跳转到产品搜索页
                    const searchUrl = '/products.asp?keyword=' + encodeURIComponent(keyword);
                    window.location.href = searchUrl;
                }
            });
        }

        // 聚焦时重新打开
        input.addEventListener('focus', () => {
            const query = input.value.trim();
            if (query.length >= this.minChars) {
                this.debounceSearch(query);
            }
        });
    }

    updateSelection(items) {
        items.forEach((item, i) => {
            item.classList.toggle('active', i === this._selectedIndex);
        });
        
        // 滚动到可见
        if (this._selectedIndex >= 0 && items[this._selectedIndex]) {
            items[this._selectedIndex].scrollIntoView({ block: 'nearest' });
        }
    }

    debounceSearch(query) {
        clearTimeout(this._debounceTimer);
        this._debounceTimer = setTimeout(() => {
            this.fetchSuggestions(query);
        }, this.debounce);
    }

    async fetchSuggestions(query) {
        // 取消上一次请求
        if (this._abortController) {
            this._abortController.abort();
        }
        this._abortController = new AbortController();

        try {
            const url = this.api + '?q=' + encodeURIComponent(query) + '&limit=' + this.maxResults;
            const response = await fetch(url, {
                signal: this._abortController.signal,
                headers: { 'Accept': 'application/json' }
            });

            if (!response.ok) return;

            const data = await response.json();
            
            // 兼容不同 API 响应格式
            let results = [];
            if (Array.isArray(data)) {
                results = data;
            } else if (data.data && Array.isArray(data.data)) {
                results = data.data;
            } else if (data.suggestions && Array.isArray(data.suggestions)) {
                results = data.suggestions;
            }

            if (results.length > 0) {
                this._results = results;
                this.showDropdown(results);
            } else {
                this.closeDropdown();
            }
        } catch (err) {
            if (err.name !== 'AbortError') {
                console.warn('SearchAutocomplete: fetch error', err.message);
            }
        }
    }

    showDropdown(results) {
        const dropdown = this.shadowRoot.querySelector('.autocomplete-dropdown');
        if (!dropdown) return;

        this._isOpen = true;
        this._selectedIndex = -1;

        let html = '';
        results.forEach((item, index) => {
            const name = item.name || item.ProductName || item.title || '';
            const price = item.price || item.BasePrice || '';
            const image = item.image || item.ImageURL || '';
            const category = item.category || item.Category || '';
            const type = item.type || item.ProductType || '';

            html += `
                <div class="suggestion-item" data-index="${index}" role="option">
                    ${image ? '<div class="suggestion-image"><img src="' + this.escapeHTML(image) + '" alt="" onerror="this.style.display=\'none\'"></div>' : ''}
                    <div class="suggestion-content">
                        <div class="suggestion-name">${this.escapeHTML(name)}</div>
                        <div class="suggestion-meta">
                            ${category ? '<span class="suggestion-category">' + this.escapeHTML(category) + '</span>' : ''}
                            ${price ? '<span class="suggestion-price">¥' + parseFloat(price).toFixed(2) + '</span>' : ''}
                        </div>
                    </div>
                </div>
            `;
        });

        dropdown.innerHTML = html;
        dropdown.classList.add('open');

        // 绑定点击事件
        dropdown.querySelectorAll('.suggestion-item').forEach((el, i) => {
            el.addEventListener('click', () => {
                this.selectItem(results[i]);
            });
            
            el.addEventListener('mouseenter', () => {
                this._selectedIndex = i;
                this.updateSelection(dropdown.querySelectorAll('.suggestion-item'));
            });
        });

        // 显示"查看全部"链接
        const input = this.shadowRoot.querySelector('input');
        if (input && input.value.trim()) {
            const viewAll = document.createElement('div');
            viewAll.className = 'suggestion-view-all';
            viewAll.textContent = '查看全部结果 →';
            viewAll.addEventListener('click', () => {
                const keyword = input.value.trim();
                window.location.href = '/products.asp?keyword=' + encodeURIComponent(keyword);
            });
            dropdown.appendChild(viewAll);
        }
    }

    selectItem(item) {
        const input = this.shadowRoot.querySelector('input');
        if (input) {
            input.value = item.name || item.ProductName || item.title || '';
        }
        
        this.dispatchEvent(new CustomEvent('select', {
            detail: { item },
            bubbles: true,
            composed: true
        }));

        this.closeDropdown();

        // 如果有产品ID，跳转到产品详情
        const id = item.id || item.ProductID;
        if (id) {
            window.location.href = '/product.asp?id=' + id;
        }
    }

    closeDropdown() {
        const dropdown = this.shadowRoot.querySelector('.autocomplete-dropdown');
        if (dropdown) {
            dropdown.classList.remove('open');
            dropdown.innerHTML = '';
        }
        this._isOpen = false;
        this._selectedIndex = -1;
        this._results = [];
    }

    escapeHTML(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    get api() {
        return this.dataset.api || '/api/search_suggestions.asp';
    }

    get debounce() {
        return parseInt(this.dataset.debounce) || 300;
    }

    get minChars() {
        return parseInt(this.dataset.minChars) || 2;
    }

    get maxResults() {
        return parseInt(this.dataset.maxResults) || 8;
    }

    render() {
        const placeholder = this.dataset.placeholder || '搜索香水、品牌、香调...';
        const name = this.dataset.name || 'keyword';
        const value = this.dataset.value || '';
        const showIcon = this.dataset.showIcon !== 'false';

        this.shadowRoot.innerHTML = `
            <style>
                :host {
                    display: block;
                    position: relative;
                }
                .search-form {
                    display: flex;
                    align-items: center;
                    position: relative;
                }
                .search-input-wrapper {
                    position: relative;
                    flex: 1;
                    display: flex;
                    align-items: center;
                }
                .search-icon {
                    position: absolute;
                    left: 12px;
                    color: #999;
                    font-size: 16px;
                    pointer-events: none;
                    z-index: 1;
                }
                .search-input {
                    width: 100%;
                    padding: 10px 40px 10px ${showIcon ? '38px' : '14px'};
                    border: 2px solid #e0e0e0;
                    border-radius: 24px;
                    font-size: 14px;
                    outline: none;
                    transition: border-color 0.3s, box-shadow 0.3s;
                    background: #fff;
                    box-sizing: border-box;
                }
                .search-input:focus {
                    border-color: #667eea;
                    box-shadow: 0 0 0 3px rgba(102,126,234,0.15);
                }
                .search-input::placeholder {
                    color: #bbb;
                }
                .search-btn {
                    display: none;
                    position: absolute;
                    right: 4px;
                    top: 50%;
                    transform: translateY(-50%);
                    background: #667eea;
                    border: none;
                    color: #fff;
                    width: 32px;
                    height: 32px;
                    border-radius: 50%;
                    cursor: pointer;
                    font-size: 14px;
                    transition: background 0.2s;
                }
                .search-input:focus ~ .search-btn,
                .search-input:not(:placeholder-shown) ~ .search-btn {
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
                .search-btn:hover {
                    background: #5a6fd6;
                }
                .autocomplete-dropdown {
                    position: absolute;
                    top: calc(100% + 4px);
                    left: 0;
                    right: 0;
                    background: #fff;
                    border-radius: 12px;
                    box-shadow: 0 8px 30px rgba(0,0,0,0.12);
                    z-index: 1000;
                    max-height: 360px;
                    overflow-y: auto;
                    display: none;
                }
                .autocomplete-dropdown.open {
                    display: block;
                }
                .suggestion-item {
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    padding: 10px 14px;
                    cursor: pointer;
                    transition: background 0.15s;
                }
                .suggestion-item:first-child {
                    border-radius: 12px 12px 0 0;
                }
                .suggestion-item.active,
                .suggestion-item:hover {
                    background: #f5f7ff;
                }
                .suggestion-image {
                    width: 40px;
                    height: 40px;
                    border-radius: 6px;
                    overflow: hidden;
                    flex-shrink: 0;
                    background: #f5f5f5;
                }
                .suggestion-image img {
                    width: 100%;
                    height: 100%;
                    object-fit: cover;
                }
                .suggestion-content {
                    flex: 1;
                    min-width: 0;
                }
                .suggestion-name {
                    font-size: 14px;
                    font-weight: 500;
                    color: #333;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }
                .suggestion-meta {
                    display: flex;
                    gap: 8px;
                    margin-top: 2px;
                }
                .suggestion-category {
                    font-size: 11px;
                    color: #999;
                }
                .suggestion-price {
                    font-size: 12px;
                    color: #e74c3c;
                    font-weight: 600;
                }
                .suggestion-view-all {
                    padding: 10px 14px;
                    text-align: center;
                    color: #667eea;
                    font-size: 13px;
                    font-weight: 600;
                    cursor: pointer;
                    border-top: 1px solid #f0f0f0;
                    transition: background 0.15s;
                }
                .suggestion-view-all:hover {
                    background: #f5f7ff;
                    border-radius: 0 0 12px 12px;
                }
                @media (max-width: 768px) {
                    .search-input {
                        font-size: 16px;
                        padding: 12px 40px 12px ${showIcon ? '38px' : '14px'};
                    }
                    .autocomplete-dropdown {
                        max-height: 280px;
                    }
                }
            </style>
            <form class="search-form" role="search" autocomplete="off">
                <div class="search-input-wrapper">
                    ${showIcon ? '<span class="search-icon">🔍</span>' : ''}
                    <input type="search" class="search-input" name="${name}" 
                        placeholder="${this.escapeHTML(placeholder)}" 
                        value="${this.escapeHTML(value)}"
                        aria-label="${this.escapeHTML(placeholder)}"
                        aria-autocomplete="list" role="combobox" aria-expanded="false"
                        autocomplete="off">
                    <button type="submit" class="search-btn" aria-label="搜索">→</button>
                </div>
                <div class="autocomplete-dropdown" role="listbox"></div>
            </form>
        `;
    }
}

customElements.define('search-autocomplete', SearchAutocomplete);
