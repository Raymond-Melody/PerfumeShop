# V16.0 API 接口文档

> **版本**: V16.0 | **更新**: 2026-06-24 | **基础URL**: `https://你的域名/api/`

---

## 目录

1. [概述](#概述)
2. [通用规范](#通用规范)
3. [购物车 API](#购物车-api)
4. [收藏 API](#收藏-api)
5. [订单 API](#订单-api)
6. [地址 API](#地址-api)
7. [通知 API](#通知-api)
8. [管理工具 API](#管理工具-api)
9. [追踪 API](#追踪-api)
10. [文件上传 API](#文件上传-api)
11. [风控 API](#风控-api)
12. [备份状态 API](#备份状态-api)
13. [错误码参考](#错误码参考)

---

## 概述

V16.0 提供 17 个 API 端点，覆盖购物车、收藏、订单、地址、通知、管理工具等核心业务。所有 API 遵循统一的 JSON 响应格式。

### 响应格式

```json
{
  "success": true,
  "code": 0,
  "message": "操作成功",
  "data": { },
  "requestId": "2026062415304500A3F2"
}
```

### 认证方式

- **Session 认证**: 前端请求自动携带 ASP Session Cookie
- **CSRF 保护**: 写操作（POST/PUT/DELETE）需携带 CSRF Token，通过隐藏域 `csrf_token` 提交

---

## 通用规范

### 请求方法

| 操作类型 | HTTP 方法 | CSRF 验证 |
|---------|----------|----------|
| 读取 | GET | 不需要 |
| 写入 | POST | **需要** |
| 删除 | POST | **需要** |

### 通用响应头

| Header | 值 | 说明 |
|--------|---|------|
| Content-Type | application/json; charset=UTF-8 | UTF-8 编码 JSON |
| Cache-Control | no-cache | 禁止缓存（数据接口） |
| ETag | `"SHA256-HASH"` | 条件请求支持（cart_count, favorites） |

---

## 购物车 API

### 1. GET /api/cart_count.asp — 获取购物车数量

**认证**: 不需要（基于 Session）

**请求参数**: 无

**成功响应**:
```json
{
  "success": true, "code": 0, "message": "success",
  "data": { "count": 5 }
}
```

**缓存**: 支持 ETag，客户端发送 `If-None-Match` 可返回 `304 Not Modified`

---

### 2. POST /api/cart_add.asp — 添加到购物车

**认证**: 不需要 | **CSRF**: 需要

**请求参数** (application/x-www-form-urlencoded):

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| productId | int | ✅ | 产品ID |
| volume | int | ✅ | 容量ID |
| bottle | int | 否 | 瓶身ID |
| topNote | string | 定制 | 前调香调ID列表(逗号分隔) |
| middleNote | string | 定制 | 中调香调ID列表(逗号分隔) |
| baseNote | string | 定制 | 后调香调ID列表(逗号分隔) |
| percent_top_{id} | float | 定制 | 前调百分比 |
| percent_mid_{id} | float | 定制 | 中调百分比 |
| percent_base_{id} | float | 定制 | 后调百分比 |
| customLabel | string | 否 | 刻字内容 |
| quantity | int | 否 | 数量(默认1) |
| buyNow | string | 否 | 设为任意值触发"立即购买"跳转 |

**成功响应**:
```json
{
  "success": true, "code": 0, "message": "已添加到购物车！",
  "data": {
    "cartId": 128,
    "cartCount": 5,
    "unitPrice": 298.00,
    "quantity": 1,
    "redirect": ""
  }
}
```

**业务规则**:
- 定制香水配比总和必须 = 100%，每种调性最少 10%
- 库存不足时返回 `code: 3004`

---

### 3. POST /api/cart_update.asp — 更新购物车数量

**认证**: 不需要 | **CSRF**: 需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| cartId | int | ✅ | 购物车项ID |
| delta | int | ✅ | 变化量(+1或-1) |

**成功响应**:
```json
{
  "success": true, "code": 0, "message": "success",
  "data": { "quantity": 3, "stock": 50 }
}
```

**V16 新增**: 实时库存检查，不足时返回 `code: 3004`

---

### 4. POST /api/cart_remove.asp — 移除购物车项

**认证**: 不需要 | **CSRF**: 需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| cartId | int | ✅ | 购物车项ID |

**成功响应**:
```json
{ "success": true, "code": 0, "message": "已移除", "data": null }
```

---

### 5. POST /api/cart_clear.asp — 清空购物车

**认证**: 不需要 | **CSRF**: 需要

**请求参数**: 无

**成功响应**:
```json
{ "success": true, "code": 0, "message": "购物车已清空", "data": null }
```

---

## 收藏 API

### 6. GET/POST /api/favorites.asp — 收藏管理

**认证**: 需要登录 | **CSRF**: add/remove 需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| action | enum | ✅ | `add` / `remove` / `check` |
| productId | int | ✅ | 产品ID |

**action=add 成功响应**:
```json
{
  "success": true, "code": 0, "message": "收藏成功",
  "data": { "action": "added", "isFavorite": true }
}
```

**action=check 成功响应**:
```json
{
  "success": true, "code": 0, "message": "success",
  "data": { "isFavorite": true }
}
```

**action=check 缓存**: 支持 ETag 条件请求

---

## 订单 API

### 7. POST /api/order_cancel.asp — 取消订单

**认证**: 需要登录 | **CSRF**: 需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| orderId | int | ✅ | 订单ID |

**成功响应**:
```json
{ "success": true, "code": 0, "message": "订单已取消", "data": null }
```

**业务规则**: 只能取消状态为 `Pending` 的订单

---

### 8. POST /api/order_confirm.asp — 确认收货

**认证**: 需要登录 | **CSRF**: 需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| orderId | int | ✅ | 订单ID |

**成功响应**:
```json
{ "success": true, "code": 0, "message": "收货确认成功", "data": null }
```

**业务规则**: 只能确认状态为 `Shipped` 的订单

---

## 地址 API

### 9. GET /api/get_areas.asp — 获取地区列表

**认证**: 不需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| parent_id | int | 否 | 父级地区ID(默认0=顶级) |
| parent_name | string | 否 | 父级地区名称(替代parent_id) |
| level | int | 否 | 层级(1=省,2=市,3=区) |

**成功响应**:
```json
[
  { "AreaID": 110000, "AreaName": "北京市" },
  { "AreaID": 310000, "AreaName": "上海市" }
]
```

---

### 10. GET /api/get_area_name.asp — 获取地区名称

**认证**: 不需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| id | int | ✅ | 地区ID |

**成功响应** (text/plain):
```
朝阳区
```

---

## 通知 API

### 11. GET /api/notifications_sse.asp — SSE 实时通知

**认证**: 需要 Admin 登录 | **Content-Type**: `text/event-stream`

**请求参数**: 无

**响应格式** (Server-Sent Events):
```
event: notification
data: {"type":"new_order","title":"新订单","message":"订单 #20240601 已创建","timestamp":"..."}

event: heartbeat
data: ping
```

**超时**: 600 秒（10 分钟长连接）

**V16 新增**: SSE 实时推送，前端使用 `new EventSource('/api/notifications_sse.asp')` 订阅

---

## 管理工具 API

### 12. POST /api/batch_operations.asp — 批量操作

**认证**: 需要 Admin 登录 | **CSRF**: 需要

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| action | enum | ✅ | `batch_ship` / `batch_cancel` / `batch_list` / `batch_unlist` / `batch_delete_cart` |
| ids | string | ✅ | ID列表(逗号分隔)，如 `"1,2,3,4,5"` |
| tracking_no | string | batch_ship时需要 | 快递单号 |

**成功响应**:
```json
{
  "success": true, "code": 0, "message": "批量操作完成：成功5条，失败0条",
  "data": { "action": "batch_ship", "total": 5, "successCount": 5, "failCount": 0 }
}
```

**V16 新增**: 操作自动记录审计日志

---

### 13. GET /api/export_data.asp — 数据导出

**认证**: 需要 Admin 登录 | **Content-Type**: `text/csv; charset=UTF-8`

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| type | enum | ✅ | `orders` / `revenue` / `customers` / `products` |
| from | date | 否 | 开始日期(默认30天前) |
| to | date | 否 | 结束日期(默认今天) |

**使用示例**:
```
/api/export_data.asp?type=orders&from=2026-06-01&to=2026-06-24
```

**导出格式**:
- **orders**: 订单号,客户,邮箱,金额,状态,支付方式,创建时间,发货时间
- **revenue**: 日期,订单数,总营收,平均客单价
- **customers**: 用户名,邮箱,姓名,手机,注册时间,订单数,累计消费
- **products**: 产品ID,产品名称,类型,基础价格,是否活跃,库存,创建时间

**V16 新增**: 导出操作自动记录审计日志；UTF-8 BOM 确保 Excel 正确显示中文

---

## 追踪 API

### 14. GET /api/track.asp — 用户行为追踪

**认证**: 不需要 | **Content-Type**: `image/gif` (1x1透明像素)

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| action | enum | ✅ | `view` / `search` / `cart` / `fav` |
| target | int | view/cart/fav时需要 | 目标ID(产品ID) |
| keyword | string | search时需要 | 搜索关键词 |
| qty | int | cart时可选 | 数量(默认1) |

**前端使用**:
```html
<img src="/api/track.asp?action=view&target=123" width="1" height="1" alt="">
```

**响应**: 1x1 透明 GIF（行为记录在响应后异步完成，不阻塞页面）

---

## 文件上传 API

### 15. POST /api/upload.asp — 文件上传

**认证**: 需要 Admin 登录 | **CSRF**: 需要 | **Content-Type**: `multipart/form-data`

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| type | enum | ✅ | `product` / `note` / `bottle` / `avatar` |
| file | file | ✅ | 上传文件 |
| csrf_token | string | ✅ | CSRF Token（表单域） |

**支持格式**: JPG, JPEG, PNG, GIF, SVG, WebP, BMP

**大小限制**: 由 config.asp 中 `MAX_UPLOAD_SIZE` 配置

**成功响应**:
```json
{
  "success": true,
  "url": "/images/products/20260624153045_a3f2.jpg",
  "fileName": "perfume.jpg",
  "fileSize": 245760
}
```

**安全特性**:
- 文件魔数（Magic Bytes）验证（非 SVG 文件）
- 随机文件名生成，防止路径遍历
- MIME 类型白名单校验

---

## 风控 API

### 16. POST /api/risk_check.asp — 下单风险检查

**认证**: 不需要 | **Content-Type**: `application/json`

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| userId | int | 否 | 用户ID(未登录传0) |
| orderTotal | float | ✅ | 订单总金额 |
| productIds | string | ✅ | 产品ID列表 |
| shippingAddress | string | ✅ | 收货地址 |
| shippingPhone | string | ✅ | 收货手机号 |

**成功响应**:
```json
{
  "riskLevel": "low",
  "risks": {},
  "riskCount": 0,
  "canProceed": true,
  "maxRisk": "low",
  "description": "风险等级: 低 — 可以继续下单"
}
```

**风险等级**: `low` → `medium` → `high` → `blocked`

**检查维度**: 用户信用（退货率/取消率/未付款订单）、金额异常、地址重复

---

## 备份状态 API

### 17. GET /api/backup_status.asp — 备份系统状态

**认证**: 不需要 | **Content-Type**: `application/json`

**请求参数**: 无

**成功响应**:
```json
{
  "status": "ok",
  "version": "V16.0",
  "lastBackup": {
    "fileName": "PerfumeShop_20260624_020000.bak",
    "sizeBytes": 15728640,
    "sizeMB": 15.00,
    "time": "2026-06-24 02:00:15",
    "verified": true
  },
  "totals": {
    "totalBackups": 14,
    "recent30Days": 7,
    "databaseSizeMB": 45.50
  },
  "schedule": {
    "frequency": "daily",
    "time": "02:00",
    "nextRun": "2026-06-25T02:00:00"
  },
  "generatedAt": "2026-06-24 15:30:45"
}
```

---

## 错误码参考

### 错误响应格式

```json
{
  "success": false,
  "code": 2001,
  "message": "缺少必填参数",
  "data": null
}
```

### 错误码表

| 范围 | 错误码 | 常量名 | 说明 |
|------|--------|--------|------|
| 成功 | 0 | API_ERR_SUCCESS | 操作成功 |
| 认证 | 1001 | API_ERR_AUTH_REQUIRED | 需要登录 |
| 认证 | 1002 | API_ERR_AUTH_EXPIRED | 登录过期 |
| 认证 | 1003 | API_ERR_CSRF_INVALID | CSRF令牌无效 |
| 认证 | 1004 | API_ERR_FORBIDDEN | 权限不足 |
| 参数 | 2001 | API_ERR_PARAM_MISSING | 缺少必填参数 |
| 参数 | 2002 | API_ERR_PARAM_INVALID | 参数格式无效 |
| 参数 | 2003 | API_ERR_PARAM_TYPE | 参数类型错误 |
| 业务 | 3001 | API_ERR_NOT_FOUND | 资源不存在 |
| 业务 | 3002 | API_ERR_DUPLICATE | 资源重复 |
| 业务 | 3003 | API_ERR_LIMIT_EXCEEDED | 超出限制 |
| 业务 | 3004 | API_ERR_BUSINESS_RULE | 业务规则限制 |
| 数据库 | 4001 | API_ERR_DB_ERROR | 数据库错误 |
| 数据库 | 4002 | API_ERR_DB_TIMEOUT | 数据库超时 |
| 数据库 | 4003 | API_ERR_DB_DEADLOCK | 数据库死锁 |
| 文件 | 5001 | API_ERR_FILE_UPLOAD | 文件上传失败 |
| 文件 | 5002 | API_ERR_FILE_TYPE | 文件类型不支持 |
| 文件 | 5003 | API_ERR_FILE_SIZE | 文件大小超限 |
| 服务器 | 6001 | API_ERR_SERVER_ERROR | 服务器内部错误 |
| 服务器 | 6002 | API_ERR_MAINTENANCE | 系统维护中 |

---

## 前端调用示例

### JavaScript (fetch)

```javascript
// 添加到购物车
async function addToCart(productId, volume, quantity) {
  const formData = new FormData();
  formData.append('productId', productId);
  formData.append('volume', volume);
  formData.append('quantity', quantity);
  // CSRF Token 从页面隐藏域获取
  formData.append('csrf_token', document.querySelector('[name=csrf_token]').value);

  const res = await fetch('/api/cart_add.asp', { method: 'POST', body: formData });
  const data = await res.json();
  
  if (data.success) {
    console.log(`已添加，购物车共 ${data.data.cartCount} 件`);
    if (data.data.redirect) location.href = data.data.redirect;
  } else {
    alert(data.message);
  }
}

// 获取购物车数量（带缓存）
async function getCartCount() {
  const res = await fetch('/api/cart_count.asp', {
    headers: { 'If-None-Match': localStorage.getItem('cart_etag') || '' }
  });
  if (res.status === 304) return parseInt(localStorage.getItem('cart_count') || '0');
  const data = await res.json();
  localStorage.setItem('cart_etag', res.headers.get('ETag') || '');
  localStorage.setItem('cart_count', data.data.count);
  return data.data.count;
}

// SSE 实时通知
const es = new EventSource('/api/notifications_sse.asp');
es.addEventListener('notification', (e) => {
  const notif = JSON.parse(e.data);
  showToast(notif.title, notif.message);
});
```

### jQuery

```javascript
// 更新购物车数量
$.post('/api/cart_update.asp', {
  cartId: 128,
  delta: 1,
  csrf_token: $('[name=csrf_token]').val()
}, function(res) {
  if (res.success) {
    $('#qty-' + 128).text(res.data.quantity);
  }
}, 'json');
```

---

> **注意**: 所有 POST 请求必须携带 `csrf_token` 参数。CSRF Token 通过 ASP Session 生成，可从页面隐藏域 `<input type="hidden" name="csrf_token" value="<%=Session("CSRFToken")%>">` 获取。
