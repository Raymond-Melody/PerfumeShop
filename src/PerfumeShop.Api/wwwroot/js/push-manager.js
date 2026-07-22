/**
 * PerfumeShop V18.0 - 推送通知管理器
 * 前端推送订阅管理：subscribe/unsubscribe、Permission状态管理、VAPID配置
 */
(function() {
    'use strict';

    // VAPID 公钥（由后端生成，此处为占位，部署时替换）
    var VAPID_PUBLIC_KEY = 'BPLACEHOLDER_VAPID_PUBLIC_KEY_REPLACE_WITH_REAL_ONE';

    var PushManager = {
        isSupported: false,
        subscription: null,
        permission: 'default',

        /**
         * 初始化推送管理器
         */
        init: function() {
            var self = this;
            if (!('serviceWorker' in navigator) || !('PushManager' in window)) {
                console.log('[PushManager] Push API not supported');
                return;
            }

            this.isSupported = true;
            this.permission = Notification.permission;

            // 检查已有订阅
            navigator.serviceWorker.ready.then(function(registration) {
                return registration.pushManager.getSubscription();
            }).then(function(subscription) {
                self.subscription = subscription;
                self._updateUI();
                console.log('[PushManager] Initialized, subscribed:', !!subscription);
            });
        },

        /**
         * 订阅推送
         */
        subscribe: function() {
            var self = this;
            if (!this.isSupported) {
                console.warn('[PushManager] Push not supported');
                return Promise.reject(new Error('浏览器不支持推送通知'));
            }

            return navigator.serviceWorker.ready.then(function(registration) {
                return registration.pushManager.subscribe({
                    userVisibleOnly: true,
                    applicationServerKey: self._urlB64ToUint8Array(VAPID_PUBLIC_KEY)
                });
            }).then(function(subscription) {
                self.subscription = subscription;
                self.permission = 'granted';
                self._updateUI();
                // 发送订阅信息到后端
                return self._sendSubscriptionToServer(subscription, 'subscribe');
            }).then(function() {
                console.log('[PushManager] Subscribed successfully');
                return true;
            }).catch(function(err) {
                console.error('[PushManager] Subscribe failed:', err);
                throw err;
            });
        },

        /**
         * 取消订阅
         */
        unsubscribe: function() {
            var self = this;
            if (!this.subscription) {
                return Promise.resolve(true);
            }

            return this.subscription.unsubscribe().then(function() {
                // 通知后端删除订阅
                return self._sendSubscriptionToServer(self.subscription, 'unsubscribe');
            }).then(function() {
                self.subscription = null;
                self._updateUI();
                console.log('[PushManager] Unsubscribed');
                return true;
            }).catch(function(err) {
                console.error('[PushManager] Unsubscribe failed:', err);
                throw err;
            });
        },

        /**
         * 切换订阅状态
         */
        toggle: function() {
            if (this.subscription) {
                return this.unsubscribe();
            } else {
                return this.subscribe();
            }
        },

        /**
         * 发送订阅到后端
         */
        _sendSubscriptionToServer: function(subscription, action) {
            var data = {
                action: action,
                endpoint: subscription.endpoint,
                p256dh: subscription.getKey ? btoa(String.fromCharCode.apply(null, new Uint8Array(subscription.getKey('p256dh')))) : '',
                auth: subscription.getKey ? btoa(String.fromCharCode.apply(null, new Uint8Array(subscription.getKey('auth')))) : ''
            };

            return fetch('/api/notifications_sse.asp', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: 'push_action=' + encodeURIComponent(action) +
                      '&endpoint=' + encodeURIComponent(data.endpoint) +
                      '&p256dh=' + encodeURIComponent(data.p256dh) +
                      '&auth=' + encodeURIComponent(data.auth)
            });
        },

        /**
         * 更新 UI 按钮状态
         */
        _updateUI: function() {
            var buttons = document.querySelectorAll('.push-toggle-btn');
            buttons.forEach(function(btn) {
                if (this.subscription) {
                    btn.classList.add('subscribed');
                    btn.textContent = btn.dataset.labelOn || '已订阅';
                } else {
                    btn.classList.remove('subscribed');
                    btn.textContent = btn.dataset.labelOff || '订阅通知';
                }
                // 权限被拒绝时隐藏按钮
                if (this.permission === 'denied') {
                    btn.style.display = 'none';
                } else {
                    btn.style.display = '';
                }
            }.bind(this));
        },

        /**
         * Base64 URL 转 Uint8Array（VAPID 公钥解码）
         */
        _urlB64ToUint8Array: function(base64String) {
            var padding = '='.repeat((4 - base64String.length % 4) % 4);
            var base64 = (base64String + padding).replace(/\-/g, '+').replace(/_/g, '/');
            var rawData = window.atob(base64);
            var outputArray = new Uint8Array(rawData.length);
            for (var i = 0; i < rawData.length; ++i) {
                outputArray[i] = rawData.charCodeAt(i);
            }
            return outputArray;
        }
    };

    // 暴露到全局
    window.PushManager = PushManager;

    // 自动初始化
    document.addEventListener('DOMContentLoaded', function() {
        PushManager.init();
    });
})();
