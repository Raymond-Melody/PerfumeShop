/**
 * V18.0 <review-stars> Web Component
 * 星级评分交互组件，支持只读和可读写模式
 * 
 * 属性:
 *   data-value    - 当前评分值 (0-5)
 *   data-max      - 最大星数 (默认 5)
 *   data-readonly - 是否只读 (默认 false)
 *   data-size     - 星星大小: sm/md/lg (默认 md)
 *   data-color    - 自定义颜色 (默认 #f0a500)
 *   data-name     - 表单字段名 (默认 "rating")
 *   data-show-value - 是否显示数值 (默认 true)
 *
 * 事件:
 *   change - 评分改变时触发，detail: { value: number }
 *
 * 用法:
 *   <!-- 只读模式 -->
 *   <review-stars data-value="4.5" data-readonly="true"></review-stars>
 *   
 *   <!-- 交互模式 -->
 *   <review-stars data-value="0" data-name="rating"></review-stars>
 */
class ReviewStars extends HTMLElement {
    constructor() {
        super();
        this.attachShadow({ mode: 'open' });
        this._value = 0;
        this._hoverValue = 0;
    }

    static get observedAttributes() {
        return ['data-value', 'data-readonly', 'data-size', 'data-color', 'data-max'];
    }

    connectedCallback() {
        this._value = parseFloat(this.dataset.value) || 0;
        this._max = parseInt(this.dataset.max) || 5;
        this.render();
        if (this.dataset.readonly !== 'true') {
            this.setupInteraction();
        }
    }

    attributeChangedCallback(name, oldVal, newVal) {
        if (oldVal !== newVal && this.shadowRoot.innerHTML) {
            if (name === 'data-value') {
                this._value = parseFloat(newVal) || 0;
            }
            this.render();
        }
    }

    get value() {
        return this._value;
    }

    set value(val) {
        this._value = parseFloat(val) || 0;
        this.dataset.value = this._value;
        this.render();
    }

    setupInteraction() {
        const stars = this.shadowRoot.querySelectorAll('.star-interactive');
        stars.forEach((star, index) => {
            const starValue = index + 1;
            
            star.addEventListener('mouseenter', () => {
                this._hoverValue = starValue;
                this.highlightStars(starValue);
            });
            
            star.addEventListener('mouseleave', () => {
                this._hoverValue = 0;
                this.highlightStars(this._value);
            });
            
            star.addEventListener('click', () => {
                // 如果点击的是当前已选中的星，且是整星，取消选择
                if (this._value === starValue) {
                    this._value = 0;
                } else {
                    this._value = starValue;
                }
                this.dataset.value = this._value;
                this.highlightStars(this._value);
                this.updateValueDisplay();
                
                // 触发 change 事件
                this.dispatchEvent(new CustomEvent('change', {
                    detail: { value: this._value },
                    bubbles: true,
                    composed: true
                }));
            });
        });

        // 支持半星：按住 Alt 点击为半星
        stars.forEach((star, index) => {
            star.addEventListener('click', (e) => {
                if (e.altKey) {
                    e.preventDefault();
                    e.stopPropagation();
                    const halfValue = index + 0.5;
                    this._value = halfValue;
                    this.dataset.value = this._value;
                    this.highlightStars(this._value);
                    this.updateValueDisplay();
                    
                    this.dispatchEvent(new CustomEvent('change', {
                        detail: { value: this._value },
                        bubbles: true,
                        composed: true
                    }));
                }
            });
        });
    }

    highlightStars(value) {
        const stars = this.shadowRoot.querySelectorAll('.star-interactive');
        stars.forEach((star, index) => {
            const starValue = index + 1;
            star.classList.remove('full', 'half');
            if (starValue <= Math.floor(value)) {
                star.classList.add('full');
            } else if (starValue === Math.ceil(value) && value % 1 >= 0.5) {
                star.classList.add('half');
            }
        });
    }

    updateValueDisplay() {
        const display = this.shadowRoot.querySelector('.rating-value');
        if (display) {
            display.textContent = this._value > 0 ? this._value.toFixed(1) : '';
        }
    }

    getStarsHTML(readonly) {
        const size = this.dataset.size || 'md';
        const color = this.dataset.color || '#f0a500';
        const value = this._value;
        
        let html = '';
        for (let i = 1; i <= this._max; i++) {
            const filled = i <= Math.floor(value);
            const half = !filled && i === Math.ceil(value) && value % 1 >= 0.5;
            
            if (readonly) {
                if (filled) {
                    html += `<span class="star star-${size} star-filled">★</span>`;
                } else if (half) {
                    html += `<span class="star star-${size} star-half">★</span>`;
                } else {
                    html += `<span class="star star-${size} star-empty">☆</span>`;
                }
            } else {
                html += `<span class="star star-${size} star-interactive" data-index="${i}">☆</span>`;
            }
        }
        return html;
    }

    render() {
        const readonly = this.dataset.readonly === 'true';
        const size = this.dataset.size || 'md';
        const color = this.dataset.color || '#f0a500';
        const name = this.dataset.name || 'rating';
        const showValue = this.dataset.showValue !== 'false';
        
        // Size dimensions
        const sizeMap = { sm: '16px', md: '24px', lg: '32px' };
        const fontSize = sizeMap[size] || '24px';
        
        this.shadowRoot.innerHTML = `
            <style>
                :host {
                    display: inline-flex;
                    align-items: center;
                    gap: 6px;
                    --star-color: ${color};
                }
                .stars-container {
                    display: inline-flex;
                    align-items: center;
                    gap: 2px;
                    line-height: 1;
                }
                .star {
                    font-size: ${fontSize};
                    transition: transform 0.15s ease, color 0.15s ease;
                    cursor: default;
                    user-select: none;
                }
                .star-interactive {
                    cursor: pointer;
                    color: #ddd;
                }
                .star-interactive:hover {
                    transform: scale(1.2);
                }
                .star-interactive.full {
                    color: var(--star-color);
                }
                .star-interactive.half {
                    color: var(--star-color);
                    position: relative;
                }
                .star-interactive.half::after {
                    content: '★';
                    position: absolute;
                    left: 0;
                    top: 0;
                    width: 50%;
                    overflow: hidden;
                    color: var(--star-color);
                }
                .star-filled {
                    color: var(--star-color);
                }
                .star-half {
                    color: #ddd;
                    position: relative;
                    display: inline-block;
                }
                .star-half::before {
                    content: '★';
                    position: absolute;
                    left: 0;
                    top: 0;
                    width: 50%;
                    overflow: hidden;
                    color: var(--star-color);
                    z-index: 1;
                }
                .star-empty {
                    color: #ddd;
                }
                .rating-value {
                    font-size: 13px;
                    color: #888;
                    font-weight: 600;
                    min-width: 28px;
                }
                input[type="hidden"] {
                    display: none;
                }
                @media (max-width: 768px) {
                    .star-sm { font-size: 14px; }
                    .star-md { font-size: 20px; }
                    .star-lg { font-size: 28px; }
                }
            </style>
            <div class="stars-container">
                ${this.getStarsHTML(readonly)}
            </div>
            ${showValue && readonly && this._value > 0 ? '<span class="rating-value">' + this._value.toFixed(1) + '</span>' : ''}
            ${showValue && !readonly ? '<span class="rating-value">' + (this._value > 0 ? this._value.toFixed(1) : '') + '</span>' : ''}
            ${!readonly ? '<input type="hidden" name="' + name + '" value="' + this._value + '">' : ''}
        `;
    }
}

customElements.define('review-stars', ReviewStars);
