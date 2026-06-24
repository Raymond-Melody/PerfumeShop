/**
 * PerfumeShop PWA Service Worker
 * V16.0 - 完整离线缓存 + Stale-While-Revalidate + 缓存过期 + 推送通知 + API NetworkFirst
 */

// ── 缓存配置 ──
var CACHE_VERSION = 'perfumeshop-v16';
var CACHE_STATIC = CACHE_VERSION + '-static';
var CACHE_DYNAMIC = CACHE_VERSION + '-dynamic';
var CACHE_IMAGES = CACHE_VERSION + '-images';
var OFFLINE_PAGE = '/offline.html';

// 缓存过期时间（毫秒）
var MAX_AGE_STATIC = 365 * 24 * 60 * 60 * 1000;  // 1 年
var MAX_AGE_DYNAMIC = 7 * 24 * 60 * 60 * 1000;   // 7 天
var MAX_AGE_IMAGES = 30 * 24 * 60 * 60 * 1000;    // 30 天

// 静态核心资源（安装时预缓存）
var STATIC_ASSETS = [
  '/',
  '/index.asp',
  '/products.asp',
  '/about.asp',
  '/contact.asp',
  '/css/design-tokens.css?v=16.0',
  '/css/style.css?v=16.0',
  '/css/pages.css?v=16.0',
  '/css/buttons.css?v=16.0',
  '/css/responsive.css?v=16.0',
  '/css/lazy-load.css?v=16.0',
  '/css/cart-animation.css?v=16.0',
  '/css/filter-optimization.css?v=16.0',
  '/css/theme.css?v=16.0',
  '/css/skeleton.css?v=16.0',
  '/js/main.js?v=16.0',
  '/js/lazy-load.js?v=16.0',
  '/js/theme-toggle.js?v=16.0',
  '/js/skeleton-loader.js?v=16.0',
  '/images/default-product.svg',
  '/images/default-avatar.svg',
  OFFLINE_PAGE,
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css',
  'https://code.jquery.com/jquery-3.6.0.min.js'
];

// 跨域白名单
var CDN_HOSTS = [
  'cdnjs.cloudflare.com',
  'code.jquery.com',
  'cdn.ampproject.org'
];

// ── 工具函数 ──

// 检查 URL 是否为跨域 CDN
function isCDNHost(url) {
  for (var i = 0; i < CDN_HOSTS.length; i++) {
    if (url.hostname === CDN_HOSTS[i]) return true;
  }
  return false;
}

// 判断请求类型
function getRequestType(url, request) {
  var accept = request.headers.get('Accept') || '';
  if (url.pathname.match(/\.(png|jpg|jpeg|gif|svg|webp|ico)$/i)) return 'image';
  if (url.pathname.match(/\.(css)$/i)) return 'style';
  if (url.pathname.match(/\.(js)$/i)) return 'script';
  if (url.pathname.match(/\.(woff2?|ttf|eot)$/i)) return 'font';
  if (accept.indexOf('text/html') >= 0) return 'html';
  return 'other';
}

// 带过期时间的缓存写入
function cacheWithTimestamp(cacheName, request, response) {
  var body = response.body;
  var headers = new Headers(response.headers);
  headers.set('sw-cached-at', Date.now().toString());

  var cachedResponse = new Response(body, {
    status: response.status,
    statusText: response.statusText,
    headers: headers
  });

  return caches.open(cacheName).then(function(cache) {
    return cache.put(request, cachedResponse);
  });
}

// 检查缓存是否过期
function isCacheExpired(response, maxAge) {
  var cachedAt = response.headers.get('sw-cached-at');
  if (!cachedAt) return false; // 无时间戳视为未过期
  var age = Date.now() - parseInt(cachedAt, 10);
  return age > maxAge;
}

// 清理过期缓存条目
function purgeExpiredCache(cacheName, maxAge) {
  return caches.open(cacheName).then(function(cache) {
    return cache.keys().then(function(requests) {
      var promises = requests.map(function(request) {
        return cache.match(request).then(function(response) {
          if (response && isCacheExpired(response, maxAge)) {
            return cache.delete(request);
          }
          return false;
        });
      });
      return Promise.all(promises);
    });
  });
}

// ── 安装事件 ──
self.addEventListener('install', function(event) {
  console.log('[SW V16] Installing...');
  event.waitUntil(
    caches.open(CACHE_STATIC)
      .then(function(cache) {
        console.log('[SW V16] Caching ' + STATIC_ASSETS.length + ' static assets...');
        // 逐个缓存，允许单个失败不影响整体
        var promises = STATIC_ASSETS.map(function(url) {
          return cache.add(url).catch(function(err) {
            console.warn('[SW V16] Failed to cache: ' + url, err.message);
          });
        });
        return Promise.all(promises);
      })
      .then(function() {
        return self.skipWaiting();
      })
  );
});

// ── 激活事件 ──
self.addEventListener('activate', function(event) {
  console.log('[SW V16] Activating...');
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      // 清理所有旧版本缓存
      var deletions = cacheNames.filter(function(name) {
        return name.indexOf('perfumeshop') === 0 && name !== CACHE_STATIC && name !== CACHE_DYNAMIC && name !== CACHE_IMAGES;
      }).map(function(name) {
        console.log('[SW V16] Removing old cache: ' + name);
        return caches.delete(name);
      });
      return Promise.all(deletions);
    })
    .then(function() {
      // 清理过期缓存
      return Promise.all([
        purgeExpiredCache(CACHE_STATIC, MAX_AGE_STATIC),
        purgeExpiredCache(CACHE_DYNAMIC, MAX_AGE_DYNAMIC),
        purgeExpiredCache(CACHE_IMAGES, MAX_AGE_IMAGES)
      ]);
    })
    .then(function() {
      return self.clients.claim();
    })
  );
});

// ── 请求拦截 ──
self.addEventListener('fetch', function(event) {
  var request = event.request;
  var url;

  try {
    url = new URL(request.url);
  } catch (e) {
    return; // 无效 URL 直接跳过
  }

  // 跳过非 GET 请求
  if (request.method !== 'GET') return;

  // 确定请求类型
  var reqType = getRequestType(url, request);
  var isLocal = url.origin === location.origin;
  var isCDN = isCDNHost(url);

  // HTML 页面：Stale-While-Revalidate 策略
  if (reqType === 'html' && isLocal) {
    event.respondWith(staleWhileRevalidate(request, CACHE_DYNAMIC, MAX_AGE_DYNAMIC));
    return;
  }

  // 图片：Cache-First 策略（独立缓存桶）
  if (reqType === 'image') {
    event.respondWith(cacheFirst(request, CACHE_IMAGES, MAX_AGE_IMAGES, isCDN));
    return;
  }

  // CSS/JS/字体：Cache-First 策略（静态缓存）
  if (reqType === 'style' || reqType === 'script' || reqType === 'font') {
    if (isLocal || isCDN) {
      event.respondWith(cacheFirst(request, CACHE_STATIC, MAX_AGE_STATIC, isCDN));
      return;
    }
    return; // 非白名单外部资源不拦截
  }

  // 其他资源：Network-First 策略
  if (isLocal) {
    event.respondWith(networkFirst(request, CACHE_DYNAMIC, MAX_AGE_DYNAMIC));
  }
});

// ── 缓存策略实现 ──

// Stale-While-Revalidate：先返回缓存，同时后台更新
function staleWhileRevalidate(request, cacheName, maxAge) {
  return caches.open(cacheName).then(function(cache) {
    return cache.match(request).then(function(cachedResponse) {
      // 发起网络请求（后台更新）
      var fetchPromise = fetch(request).then(function(networkResponse) {
        if (networkResponse && networkResponse.status === 200) {
          cacheWithTimestamp(cacheName, request, networkResponse.clone());
        }
        return networkResponse;
      }).catch(function() {
        return null;
      });

      if (cachedResponse && !isCacheExpired(cachedResponse, maxAge)) {
        // 有有效缓存：立即返回，后台静默更新
        fetchPromise; // 触发后台更新但不等待
        return cachedResponse;
      }

      // 无缓存或已过期：等待网络，如失败返回离线页
      return fetchPromise.then(function(response) {
        if (response) return response;
        // 尝试返回过期缓存（比离线页好）
        if (cachedResponse) return cachedResponse;
        // HTML 请求返回离线页
        if (request.headers.get('Accept') && request.headers.get('Accept').indexOf('text/html') >= 0) {
          return caches.match(OFFLINE_PAGE);
        }
        return new Response('Offline - PerfumeShop', {
          status: 503,
          statusText: 'Service Unavailable',
          headers: { 'Content-Type': 'text/plain' }
        });
      });
    });
  });
}

// Cache-First：优先缓存，失败则网络
function cacheFirst(request, cacheName, maxAge, allowCORS) {
  return caches.match(request).then(function(cachedResponse) {
    if (cachedResponse && !isCacheExpired(cachedResponse, maxAge)) {
      return cachedResponse;
    }

    return fetch(request).then(function(networkResponse) {
      if (!networkResponse || networkResponse.status !== 200) {
        return networkResponse;
      }
      // 缓存响应（CORS 响应也可缓存）
      if (allowCORS || networkResponse.type === 'basic') {
        cacheWithTimestamp(cacheName, request, networkResponse.clone());
      }
      return networkResponse;
    }).catch(function() {
      // 离线 + 无缓存：返回过期缓存或空响应
      if (cachedResponse) return cachedResponse;
      return new Response('', { status: 503 });
    });
  });
}

// Network-First：优先网络，失败则缓存
function networkFirst(request, cacheName, maxAge) {
  return fetch(request).then(function(networkResponse) {
    if (networkResponse && networkResponse.status === 200) {
      cacheWithTimestamp(cacheName, request, networkResponse.clone());
    }
    return networkResponse;
  }).catch(function() {
    return caches.match(request).then(function(cachedResponse) {
      return cachedResponse || new Response('Offline', {
        status: 503,
        statusText: 'Service Unavailable'
      });
    });
  });
}

// ── 推送通知 ──
self.addEventListener('push', function(event) {
  if (!event.data) return;

  var data;
  try {
    data = event.data.json();
  } catch (e) {
    data = { title: '香氛定制', body: event.data.text() };
  }

  var options = {
    body: data.body || '',
    icon: '/images/icons/icon-192x192.png',
    badge: '/images/icons/icon-72x72.png',
    vibrate: [100, 50, 100],
    data: {
      url: data.url || '/index.asp',
      dateOfArrival: Date.now(),
      primaryKey: data.id || 1
    },
    actions: [
      { action: 'open', title: '查看详情' },
      { action: 'close', title: '关闭' }
    ],
    tag: data.tag || 'default',
    renotify: !!data.renotify
  };

  event.waitUntil(
    self.registration.showNotification(data.title || '香氛定制', options)
  );
});

// ── 通知点击 ──
self.addEventListener('notificationclick', function(event) {
  event.notification.close();

  var targetUrl = (event.notification.data && event.notification.data.url) || '/index.asp';

  if (event.action === 'close') return;

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(clientList) {
      // 如果已有打开的窗口，聚焦并导航
      for (var i = 0; i < clientList.length; i++) {
        var client = clientList[i];
        if (client.url.indexOf(self.location.origin) === 0 && 'focus' in client) {
          client.navigate(targetUrl);
          return client.focus();
        }
      }
      // 否则打开新窗口
      return self.clients.openWindow(targetUrl);
    })
  );
});

// ── 后台同步 ──
self.addEventListener('sync', function(event) {
  if (event.tag === 'sync-cart') {
    event.waitUntil(syncCartData());
  }
});

// 购物车数据同步
function syncCartData() {
  return self.clients.matchAll({ type: 'window' }).then(function(clients) {
    return Promise.all(clients.map(function(client) {
      return client.postMessage({
        type: 'SYNC_CART',
        timestamp: Date.now()
      });
    }));
  });
}

// ── 消息通信 ──
self.addEventListener('message', function(event) {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  if (event.data && event.data.type === 'CACHE_URLS') {
    var urls = event.data.urls || [];
    event.waitUntil(
      caches.open(CACHE_DYNAMIC).then(function(cache) {
        return Promise.all(urls.map(function(url) {
          return fetch(url).then(function(response) {
            if (response && response.status === 200) {
              return cacheWithTimestamp(CACHE_DYNAMIC, new Request(url), response);
            }
          }).catch(function() {});
        }));
      })
    );
  }
  if (event.data && event.data.type === 'CLEAR_CACHE') {
    event.waitUntil(
      caches.keys().then(function(names) {
        return Promise.all(names.filter(function(name) {
          return name.indexOf(CACHE_VERSION) === 0;
        }).map(function(name) {
          return caches.delete(name);
        }));
      }).then(function() {
        event.ports[0].postMessage({ cleared: true });
      })
    );
  }
});

console.log('[SW V16] Service Worker loaded');
