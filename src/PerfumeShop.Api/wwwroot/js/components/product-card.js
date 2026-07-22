/**
 * V18.0 <product-card> Web Component
 * 封装产品卡片：图片/名称/价格/评分/徽章
 * 
 * 属性:
 *   data-id       - 产品ID
 *   data-name     - 产品名称
 *   data-price    - 价格
 *   data-image    - 图片URL
 *   data-category - 分类
 *   data-desc     - 描述
 *   data-type     - 产品类型 (custom/standard/kol)
 *   data-rating   - 评分 (0-5)
 *   data-badge    - 自定义徽章文字
 *   data-link     - 链接URL (默认 /product.asp?id=xxx)
 *   data-lazy     - 是否懒加载图片 (默认 true)
 *
 * 用法:
 *   <product-card data-id="1" data-name="玫瑰香水" data-price="299.00"
 *       data-image="/images/perfume1.jpg" data-category="花香调"
 *       data-rating="4.5" data-type="standard"></product-card>
 */
class ProductCard extends HTMLElement {
    constructor() {
        super();
        this.attachShadow({ mode: 'open' });
    }

    static get observedAttributes() {
        return ['data-id', 'data-name', 'data-price', 'data-image', 
                'data-category', 'data-desc', 'data-type', 'data-rating', 
                'data-badge', 'data-link', 'data-lazy'];
    }

    connectedCallback() {
        this.render();
        this.setupLazyLoad();
    }

    attributeChangedCallback(name, oldVal, newVal) {
        if (oldVal !== newVal && this.shadowRoot.innerHTML) {
            this.render();
        }
    }

    setupLazyLoad() {
        if (this.dataset.lazy !== 'false') {
            const img = this.shadowRoot.querySelector('.card-image img');
            if (img && 'IntersectionObserver' in window) {
                const observer = new IntersectionObserver((entries) => {
                    entries.forEach(entry => {
                        if (entry.isIntersecting) {
                            const el = entry.target;
                            if (el.dataset.src) {
                                el.src = el.dataset.src;
                                el.removeAttribute('data-src');
                                el.classList.add('loaded');
                            }
                            observer.unobserve(el);
                        }
                    });
                }, { rootMargin: '100px' });
                if (img.dataset.src) {
                    observer.observe(img);
                }
            }
        }
    }

    getStarsHTML(rating) {
        if (!rating || rating === '0' || rating === '') return '';
        const r = parseFloat(rating);
        if (isNaN(r)) return '';
        const fullStars = Math.floor(r);
        const halfStar = r % 1 >= 0.5;
        const emptyStars = 5 - fullStars - (halfStar ? 1 : 0);
        let html = '<span class="product-stars" aria-label="评分 ' + r.toFixed(1) + ' 星">';
        html += '<i class="fas fa-star"></i>'.repeat(fullStars);
        if (halfStar) html += '<i class="fas fa-star-half-alt"></i>';
        html += '<i class="far fa-star"></i>'.repeat(emptyStars);
        html += '<span class="rating-text">' + r.toFixed(1) + '</span>';
        html += '</span>';
        return html;
    }

    render() {
        const id = this.dataset.id || '';
        const name = this.dataset.name || '未命名产品';
        const price = this.dataset.price || '0';
        const image = this.dataset.image || '/images/default-product.jpg';
        const category = this.dataset.category || '';
        const desc = this.dataset.desc || '';
        const type = this.dataset.type || 'custom';
        const rating = this.dataset.rating || '';
        const badge = this.dataset.badge || '';
        const link = this.dataset.link || ('/product.asp?id=' + id);
        const lazy = this.dataset.lazy !== 'false';
        
        // 类型徽章
        let typeBadge = '';
        if (type === 'standard') {
            typeBadge = '<span class="badge badge-fixed">品牌定香</span>';
        } else if (type === 'kol') {
            typeBadge = '<span class="badge badge-kol">KOL推荐</span>';
        } else if (type === 'custom') {
            typeBadge = '<span class="badge badge-custom">定制</span>';
        }
        
        const customBadge = badge ? '<span class="badge badge-custom-badge">' + this.escapeHTML(badge) + '</span>' : '';
        
        // 按钮文字
        const btnText = type === 'custom' ? '开始定制' : '查看详情';
        
        this.shadowRoot.innerHTML = `
            <style>
                :host {
                    display: block;
                    contain: content;
                }
                .product-card {
                    background: #fff;
                    border-radius: 12px;
                    overflow: hidden;
                    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
                    transition: transform 0.3s ease, box-shadow 0.3s ease;
                    height: 100%;
                    display: flex;
                    flex-direction: column;
                }
                .product-card:hover {
                    transform: translateY(-4px);
                    box-shadow: 0 8px 24px rgba(0,0,0,0.12);
                }
                .card-image {
                    position: relative;
                    padding-top: 100%;
                    overflow: hidden;
                    background: #f5f5f5;
                }
                .card-image img {
                    position: absolute;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    object-fit: cover;
                    transition: transform 0.5s ease, opacity 0.3s ease;
                }
                .card-image img[data-src] {
                    opacity: 0.6;
                }
                .card-image img.loaded {
                    opacity: 1;
                }
                .product-card:hover .card-image img {
                    transform: scale(1.05);
                }
                .product-badges {
                    position: absolute;
                    top: 8px;
                    left: 8px;
                    display: flex;
                    flex-direction: column;
                    gap: 4px;
                    z-index: 2;
                }
                .badge {
                    display: inline-block;
                    padding: 4px 10px;
                    border-radius: 4px;
                    font-size: 12px;
                    font-weight: 600;
                    color: #fff;
                    white-space: nowrap;
                }
                .badge-fixed { background: linear-gradient(135deg, #667eea, #764ba2); }
                .badge-kol { background: linear-gradient(135deg, #f093fb, #f5576c); }
                .badge-custom { background: linear-gradient(135deg, #4facfe, #00f2fe); }
                .badge-custom-badge { background: #ff6b6b; }
                .card-overlay {
                    position: absolute;
                    bottom: 0;
                    left: 0;
                    right: 0;
                    background: linear-gradient(transparent, rgba(0,0,0,0.6));
                    padding: 20px 12px 12px;
                    opacity: 0;
                    transform: translateY(10px);
                    transition: opacity 0.3s ease, transform 0.3s ease;
                    display: flex;
                    justify-content: center;
                }
                .product-card:hover .card-overlay {
                    opacity: 1;
                    transform: translateY(0);
                }
                .btn-view {
                    display: inline-block;
                    padding: 8px 20px;
                    background: #fff;
                    color: #333;
                    border-radius: 20px;
                    text-decoration: none;
                    font-size: 13px;
                    font-weight: 600;
                    transition: background 0.2s;
                }
                .btn-view:hover {
                    background: #f0f0f0;
                }
                .card-info {
                    padding: 12px;
                    flex: 1;
                    display: flex;
                    flex-direction: column;
                    gap: 4px;
                }
                .card-category {
                    font-size: 11px;
                    color: #999;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                .card-name {
                    margin: 0;
                    font-size: 15px;
                    font-weight: 600;
                    line-height: 1.3;
                }
                .card-name a {
                    color: #333;
                    text-decoration: none;
                    transition: color 0.2s;
                    display: -webkit-box;
                    -webkit-line-clamp: 2;
                    -webkit-box-orient: vertical;
                    overflow: hidden;
                }
                .card-name a:hover {
                    color: #667eea;
                }
                .card-desc {
                    font-size: 12px;
                    color: #888;
                    line-height: 1.4;
                    display: -webkit-box;
                    -webkit-line-clamp: 2;
                    -webkit-box-orient: vertical;
                    overflow: hidden;
                    flex: 1;
                }
                .product-stars {
                    display: flex;
                    align-items: center;
                    gap: 2px;
                    font-size: 12px;
                    color: #f0a500;
                    margin: 2px 0;
                }
                .product-stars .rating-text {
                    color: #888;
                    margin-left: 4px;
                    font-size: 11px;
                }
                .card-price {
                    display: flex;
                    align-items: baseline;
                    gap: 4px;
                    margin-top: auto;
                }
                .price {
                    font-size: 18px;
                    font-weight: 700;
                    color: #e74c3c;
                }
                .price-label {
                    font-size: 12px;
                    color: #999;
                }
                @media (max-width: 768px) {
                    .card-info { padding: 10px; }
                    .card-name { font-size: 14px; }
                    .price { font-size: 16px; }
                }
            </style>
            <div class="product-card">
                <div class="card-image">
                    <${lazy ? 'img data-src="' + this.escapeHTML(image) + '" src="data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22%3E%3Crect fill=%22%23f0f0f0%22 width=%22100%22 height=%22100%22/%3E%3C/svg%3E"' : 'img src="' + this.escapeHTML(image) + '"'} alt="${this.escapeHTML(name)}" onerror="this.src='/images/default-product.jpg'" loading="lazy">
                    <div class="product-badges">
                        ${typeBadge}
                        ${customBadge}
                    </div>
                    <div class="card-overlay">
                        <a href="${this.escapeHTML(link)}" class="btn-view">${btnText}</a>
                    </div>
                </div>
                <div class="card-info">
                    ${category ? '<span class="card-category">' + this.escapeHTML(category) + '</span>' : ''}
                    <h3 class="card-name"><a href="${this.escapeHTML(link)}">${this.escapeHTML(name)}</a></h3>
                    ${desc ? '<p class="card-desc">' + this.escapeHTML(desc) + '</p>' : ''}
                    ${this.getStarsHTML(rating)}
                    <div class="card-price">
                        <span class="price">¥${parseFloat(price).toFixed(2)}</span>
                        <span class="price-label">起</span>
                    </div>
                </div>
            </div>
        `;
    }

    escapeHTML(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }
}

customElements.define('product-card', ProductCard);
