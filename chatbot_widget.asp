<%
' ============================================
' V18.0 智能客服悬浮组件 (Chatbot Widget)
' 纯JS/CSS实现，通过AJAX调用 /api/chatbot.asp
' 用法: <!--#include file="chatbot_widget.asp"-->
' 在 footer.asp 或任意页面底部引入
' ============================================

' 仅当 V18 智能客服功能开启时才渲染
If Not FEATURE_AI_CHATBOT Then
    ' 功能关闭，不输出任何内容
Else
%>
<!-- V18 智能客服组件 -->
<style nonce="<%= Session("csp_nonce") %>">
/* ================================================
   Chatbot Widget Styles - V18
   ================================================ */
.chatbot-widget {
    position: fixed;
    bottom: 24px;
    right: 24px;
    z-index: 9998;
    font-family: 'Microsoft YaHei','PingFang SC',-apple-system,sans-serif;
}

/* 悬浮按钮 */
.chatbot-toggle {
    width: 56px;
    height: 56px;
    border-radius: 50%;
    background: linear-gradient(135deg, #8B4513, #A0522D);
    color: #fff;
    border: none;
    cursor: pointer;
    box-shadow: 0 4px 16px rgba(139,69,19,0.4);
    font-size: 24px;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: transform 0.2s, box-shadow 0.2s;
    position: relative;
}
.chatbot-toggle:hover {
    transform: scale(1.08);
    box-shadow: 0 6px 24px rgba(139,69,19,0.55);
}
.chatbot-toggle:active {
    transform: scale(0.95);
}
.chatbot-toggle .unread-dot {
    position: absolute;
    top: 4px;
    right: 4px;
    width: 12px;
    height: 12px;
    background: #ff4444;
    border-radius: 50%;
    border: 2px solid #fff;
    display: none;
}
.chatbot-toggle.open {
    background: linear-gradient(135deg, #666, #555);
    box-shadow: 0 4px 12px rgba(0,0,0,0.3);
}

/* 对话窗口 */
.chatbot-window {
    position: absolute;
    bottom: 72px;
    right: 0;
    width: 380px;
    max-height: 550px;
    background: #fff;
    border-radius: 16px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.18);
    display: none;
    flex-direction: column;
    overflow: hidden;
    animation: chatbotSlideUp 0.3s ease;
}
@keyframes chatbotSlideUp {
    from { opacity: 0; transform: translateY(20px); }
    to   { opacity: 1; transform: translateY(0); }
}
.chatbot-window.open {
    display: flex;
}

/* 头部 */
.chatbot-header {
    background: linear-gradient(135deg, #8B4513, #A0522D);
    color: #fff;
    padding: 14px 18px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-shrink: 0;
}
.chatbot-header-left {
    display: flex;
    align-items: center;
    gap: 10px;
}
.chatbot-header-icon {
    width: 36px;
    height: 36px;
    border-radius: 50%;
    background: rgba(255,255,255,0.2);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 16px;
}
.chatbot-header-text h4 {
    margin: 0;
    font-size: 15px;
    font-weight: 600;
}
.chatbot-header-text span {
    font-size: 11px;
    opacity: 0.85;
}
.chatbot-header-actions {
    display: flex;
    gap: 4px;
}
.chatbot-header-actions button {
    background: rgba(255,255,255,0.15);
    border: none;
    color: #fff;
    width: 30px;
    height: 30px;
    border-radius: 50%;
    cursor: pointer;
    font-size: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background 0.2s;
}
.chatbot-header-actions button:hover {
    background: rgba(255,255,255,0.3);
}

/* 消息区域 */
.chatbot-messages {
    flex: 1;
    overflow-y: auto;
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    min-height: 200px;
    max-height: 380px;
    background: #faf9f7;
}
.chatbot-messages::-webkit-scrollbar {
    width: 5px;
}
.chatbot-messages::-webkit-scrollbar-thumb {
    background: #d5d0ca;
    border-radius: 3px;
}

/* 消息气泡 */
.chatbot-message {
    display: flex;
    gap: 8px;
    animation: msgFadeIn 0.25s ease;
}
@keyframes msgFadeIn {
    from { opacity: 0; transform: translateY(8px); }
    to   { opacity: 1; transform: translateY(0); }
}
.chatbot-message.bot {
    align-items: flex-start;
}
.chatbot-message.user {
    flex-direction: row-reverse;
}
.chatbot-avatar {
    width: 32px;
    height: 32px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 13px;
    flex-shrink: 0;
}
.chatbot-message.bot .chatbot-avatar {
    background: #f0e6da;
    color: #8B4513;
}
.chatbot-message.user .chatbot-avatar {
    background: #8B4513;
    color: #fff;
}
.chatbot-bubble {
    max-width: 75%;
    padding: 10px 14px;
    border-radius: 14px;
    font-size: 13px;
    line-height: 1.6;
    word-break: break-word;
}
.chatbot-message.bot .chatbot-bubble {
    background: #fff;
    color: #333;
    border: 1px solid #e8e3dc;
    border-top-left-radius: 4px;
}
.chatbot-message.user .chatbot-bubble {
    background: #8B4513;
    color: #fff;
    border-top-right-radius: 4px;
}
.chatbot-bubble a {
    color: #8B4513;
    text-decoration: underline;
}
.chatbot-message.user .chatbot-bubble a {
    color: #ffd699;
}
.chatbot-time {
    font-size: 10px;
    color: #999;
    margin-top: 4px;
    padding: 0 4px;
}
.chatbot-message.user .chatbot-time {
    text-align: right;
}

/* 消息来源标签 */
.chatbot-source-tag {
    display: inline-block;
    font-size: 10px;
    padding: 1px 6px;
    border-radius: 8px;
    margin-left: 6px;
    background: #e8f5e9;
    color: #2e7d32;
}
.chatbot-source-tag.ai {
    background: #e3f2fd;
    color: #1565c0;
}
.chatbot-source-tag.handoff {
    background: #fff3e0;
    color: #e65100;
}

/* 打字指示器 */
.chatbot-typing {
    display: none;
    align-items: center;
    gap: 8px;
    padding: 4px 16px;
}
.chatbot-typing.show {
    display: flex;
}
.chatbot-typing .typing-dots {
    display: flex;
    gap: 4px;
    padding: 10px 16px;
    background: #fff;
    border: 1px solid #e8e3dc;
    border-radius: 14px;
    border-top-left-radius: 4px;
}
.chatbot-typing .typing-dots span {
    width: 7px;
    height: 7px;
    border-radius: 50%;
    background: #c5bdb5;
    animation: typingBounce 1.4s infinite ease-in-out;
}
.chatbot-typing .typing-dots span:nth-child(2) { animation-delay: 0.2s; }
.chatbot-typing .typing-dots span:nth-child(3) { animation-delay: 0.4s; }
@keyframes typingBounce {
    0%, 60%, 100% { transform: translateY(0); }
    30% { transform: translateY(-6px); }
}

/* 快捷回复 */
.chatbot-quick-replies {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    padding: 8px 16px;
    border-top: 1px solid #f0ede8;
    background: #fff;
}
.chatbot-quick-reply {
    padding: 6px 14px;
    border-radius: 16px;
    border: 1px solid #d5d0ca;
    background: #fff;
    color: #8B4513;
    font-size: 12px;
    cursor: pointer;
    transition: all 0.2s;
    white-space: nowrap;
}
.chatbot-quick-reply:hover {
    background: #8B4513;
    color: #fff;
    border-color: #8B4513;
}

/* 输入区域 */
.chatbot-input-area {
    display: flex;
    gap: 8px;
    padding: 12px 16px;
    border-top: 1px solid #e8e3dc;
    background: #fff;
    border-radius: 0 0 16px 16px;
}
.chatbot-input-area input {
    flex: 1;
    padding: 10px 14px;
    border: 1px solid #d5d0ca;
    border-radius: 20px;
    font-size: 13px;
    outline: none;
    transition: border-color 0.2s;
    font-family: inherit;
}
.chatbot-input-area input:focus {
    border-color: #8B4513;
}
.chatbot-input-area input::placeholder {
    color: #bbb;
}
.chatbot-send-btn {
    width: 40px;
    height: 40px;
    border-radius: 50%;
    background: #8B4513;
    color: #fff;
    border: none;
    cursor: pointer;
    font-size: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background 0.2s, transform 0.15s;
    flex-shrink: 0;
}
.chatbot-send-btn:hover {
    background: #A0522D;
}
.chatbot-send-btn:active {
    transform: scale(0.92);
}
.chatbot-send-btn:disabled {
    background: #ccc;
    cursor: not-allowed;
}

/* 手风琴面板 */
.chatbot-handoff-banner {
    margin: 8px 16px;
    padding: 10px 14px;
    background: #fff8e1;
    border: 1px solid #ffcc02;
    border-radius: 10px;
    font-size: 12px;
    color: #795548;
    display: none;
    align-items: center;
    gap: 8px;
}
.chatbot-handoff-banner.show {
    display: flex;
}
.chatbot-handoff-banner i {
    font-size: 16px;
    color: #ff8f00;
}

/* 移动端适配 */
@media (max-width: 480px) {
    .chatbot-widget {
        bottom: 12px;
        right: 12px;
    }
    .chatbot-toggle {
        width: 48px;
        height: 48px;
        font-size: 20px;
    }
    .chatbot-window {
        width: calc(100vw - 24px);
        right: -6px;
        bottom: 64px;
        max-height: 65vh;
        border-radius: 12px;
    }
    .chatbot-messages {
        max-height: 45vh;
    }
    .chatbot-bubble {
        max-width: 82%;
    }
}
</style>

<div class="chatbot-widget" id="chatbotWidget">
    <!-- 悬浮按钮 -->
    <button class="chatbot-toggle" id="chatbotToggle" title="智能客服">
        <i class="fas fa-comment-dots"></i>
        <span class="unread-dot" id="chatbotUnread"></span>
    </button>

    <!-- 对话窗口 -->
    <div class="chatbot-window" id="chatbotWindow">
        <!-- 头部 -->
        <div class="chatbot-header">
            <div class="chatbot-header-left">
                <div class="chatbot-header-icon"><i class="fas fa-robot"></i></div>
                <div class="chatbot-header-text">
                    <h4>智能客服</h4>
                    <span>AI 在线 · 7×24h</span>
                </div>
            </div>
            <div class="chatbot-header-actions">
                <button onclick="chatbotToggleWindow()" title="关闭"><i class="fas fa-times"></i></button>
            </div>
        </div>

        <!-- 消息区域 -->
        <div class="chatbot-messages" id="chatbotMessages">
            <!-- 欢迎消息由JS动态插入 -->
        </div>

        <!-- 打字指示器 -->
        <div class="chatbot-typing" id="chatbotTyping">
            <div class="chatbot-avatar" style="background:#f0e6da;color:#8B4513;width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;">
                <i class="fas fa-robot"></i>
            </div>
            <div class="typing-dots">
                <span></span><span></span><span></span>
            </div>
        </div>

        <!-- 快捷回复 -->
        <div class="chatbot-quick-replies" id="chatbotQuickReplies"></div>

        <!-- 人工转接提示 -->
        <div class="chatbot-handoff-banner" id="chatbotHandoffBanner">
            <i class="fas fa-headset"></i>
            <span>问题未解决？回复「<b>转人工</b>」联系人工客服</span>
        </div>

        <!-- 输入区域 -->
        <div class="chatbot-input-area">
            <input type="text" id="chatbotInput" placeholder="输入您的问题..." maxlength="300" autocomplete="off">
            <button class="chatbot-send-btn" id="chatbotSendBtn" title="发送">
                <i class="fas fa-paper-plane"></i>
            </button>
        </div>
    </div>
</div>

<script nonce="<%= Session("csp_nonce") %>">
(function() {
    'use strict';
    
    var isOpen = false;
    var sessionId = '';
    var isWaiting = false;
    var quickRepliesVisible = false;
    
    var widget = document.getElementById('chatbotWidget');
    var toggle = document.getElementById('chatbotToggle');
    var windowEl = document.getElementById('chatbotWindow');
    var messagesEl = document.getElementById('chatbotMessages');
    var typingEl = document.getElementById('chatbotTyping');
    var quickRepliesEl = document.getElementById('chatbotQuickReplies');
    var handoffBanner = document.getElementById('chatbotHandoffBanner');
    var inputEl = document.getElementById('chatbotInput');
    var sendBtn = document.getElementById('chatbotSendBtn');
    
    // 生成会话ID
    function generateSessionId() {
        return 'cb_' + Math.random().toString(36).substring(2, 10) + Date.now().toString(36);
    }
    
    // 初始化
    function init() {
        sessionId = sessionStorage.getItem('chatbot_session_id') || generateSessionId();
        sessionStorage.setItem('chatbot_session_id', sessionId);
        
        // 欢迎消息
        addWelcomeMessage();
        
        // 事件绑定
        toggle.addEventListener('click', function(e) {
            e.preventDefault();
            toggleWindow();
        });
        
        sendBtn.addEventListener('click', function(e) {
            e.preventDefault();
            sendMessage();
        });
        
        inputEl.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
            }
        });
        
        // 点击窗口外关闭（仅桌面端点击遮罩区域）
        document.addEventListener('click', function(e) {
            if (isOpen && !widget.contains(e.target)) {
                // 桌面端不自动关闭，保持对话
            }
        });
        
        // 初始快捷回复
        showDefaultQuickReplies();
    }
    
    // 切换窗口
    window.toggleWindow = function() {
        if (isOpen) {
            closeWindow();
        } else {
            openWindow();
        }
    };
    
    function openWindow() {
        isOpen = true;
        windowEl.classList.add('open');
        toggle.classList.add('open');
        toggle.innerHTML = '<i class="fas fa-times"></i>';
        inputEl.focus();
        scrollToBottom();
    }
    
    function closeWindow() {
        isOpen = false;
        windowEl.classList.remove('open');
        toggle.classList.remove('open');
        toggle.innerHTML = '<i class="fas fa-comment-dots"></i><span class="unread-dot" id="chatbotUnread"></span>';
    }
    
    // 欢迎消息
    function addWelcomeMessage() {
        var existingMsgs = messagesEl.querySelectorAll('.chatbot-message.bot');
        if (existingMsgs.length === 0) {
            addBotMessage(
                '👋 您好！我是智能客服小香。' + '\n\n' +
                '我可以帮您解答：\n' +
                '🔹 退换货政策\n' +
                '🔹 配送物流信息\n' +
                '🔹 定制香水流程\n' +
                '🔹 香调与浓度选择\n' +
                '🔹 会员权益与积分\n\n' +
                '请随时向我提问，或点击下方快捷按钮快速获取帮助！'
            );
        }
    }
    
    // 添加消息
    function addBotMessage(text, source, handoff) {
        var msgDiv = document.createElement('div');
        msgDiv.className = 'chatbot-message bot';
        
        var avatar = document.createElement('div');
        avatar.className = 'chatbot-avatar';
        avatar.innerHTML = '<i class="fas fa-robot"></i>';
        
        var bubble = document.createElement('div');
        bubble.className = 'chatbot-bubble';
        
        // 处理换行和链接
        var formattedText = text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/\n/g, '<br>')
            .replace(/&lt;a href=/g, '<a href=')
            .replace(/&lt;\/a&gt;/g, '</a>')
            .replace(/&lt;b&gt;/g, '<b>')
            .replace(/&lt;\/b&gt;/g, '</b>');
        
        bubble.innerHTML = formattedText;
        
        // 来源标签
        if (source && source !== 'faq') {
            var tag = document.createElement('span');
            tag.className = 'chatbot-source-tag ' + source;
            if (source === 'ai') {
                tag.textContent = 'AI';
            } else if (source === 'handoff' || handoff) {
                tag.textContent = '待人工';
                tag.className = 'chatbot-source-tag handoff';
            }
            bubble.appendChild(document.createTextNode(' '));
            bubble.appendChild(tag);
        }
        
        var time = document.createElement('div');
        time.className = 'chatbot-time';
        time.textContent = getTimeString();
        
        msgDiv.appendChild(avatar);
        var wrapper = document.createElement('div');
        wrapper.appendChild(bubble);
        wrapper.appendChild(time);
        msgDiv.appendChild(wrapper);
        
        messagesEl.appendChild(msgDiv);
        scrollToBottom();
        
        if (handoff) {
            showHandoffBanner();
        } else {
            hideHandoffBanner();
        }
    }
    
    function addUserMessage(text) {
        var msgDiv = document.createElement('div');
        msgDiv.className = 'chatbot-message user';
        
        var avatar = document.createElement('div');
        avatar.className = 'chatbot-avatar';
        avatar.innerHTML = '<i class="fas fa-user"></i>';
        
        var bubble = document.createElement('div');
        bubble.className = 'chatbot-bubble';
        bubble.textContent = text;
        
        var time = document.createElement('div');
        time.className = 'chatbot-time';
        time.textContent = getTimeString();
        
        msgDiv.appendChild(avatar);
        var wrapper = document.createElement('div');
        wrapper.appendChild(bubble);
        wrapper.appendChild(time);
        msgDiv.appendChild(wrapper);
        
        messagesEl.appendChild(msgDiv);
        scrollToBottom();
    }
    
    function showTyping() {
        typingEl.classList.add('show');
        hideQuickReplies();
        scrollToBottom();
    }
    
    function hideTyping() {
        typingEl.classList.remove('show');
    }
    
    function showQuickReplies(suggestions) {
        quickRepliesEl.innerHTML = '';
        if (suggestions && suggestions.length > 0) {
            suggestions.forEach(function(s) {
                var btn = document.createElement('button');
                btn.className = 'chatbot-quick-reply';
                btn.textContent = s;
                btn.addEventListener('click', function() {
                    inputEl.value = s;
                    sendMessage();
                });
                quickRepliesEl.appendChild(btn);
            });
            quickRepliesVisible = true;
        }
    }
    
    function showDefaultQuickReplies() {
        var defaults = ['如何退换货？', '配送多久能到？', '怎么定制香水？', '有哪些香调？', '转人工客服'];
        showQuickReplies(defaults);
    }
    
    function hideQuickReplies() {
        quickRepliesEl.innerHTML = '';
        quickRepliesVisible = false;
    }
    
    function showHandoffBanner() {
        handoffBanner.classList.add('show');
    }
    
    function hideHandoffBanner() {
        handoffBanner.classList.remove('show');
    }
    
    function scrollToBottom() {
        setTimeout(function() {
            messagesEl.scrollTop = messagesEl.scrollHeight;
        }, 50);
    }
    
    function getTimeString() {
        var now = new Date();
        return ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);
    }
    
    // 发送消息
    function sendMessage() {
        var text = inputEl.value.trim();
        if (text === '' || isWaiting) return;
        
        // 显示用户消息
        addUserMessage(text);
        inputEl.value = '';
        isWaiting = true;
        sendBtn.disabled = true;
        
        // 显示打字动画
        showTyping();
        
        // AJAX 请求
        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/api/chatbot.asp', true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
        
        xhr.onload = function() {
            hideTyping();
            isWaiting = false;
            sendBtn.disabled = false;
            inputEl.focus();
            
            if (xhr.status === 200) {
                try {
                    var res = JSON.parse(xhr.responseText);
                    if (res.code === 0 && res.data) {
                        var data = res.data;
                        addBotMessage(data.reply, data.source, data.handoff);
                        
                        // 更新会话ID
                        if (data.session_id) {
                            sessionId = data.session_id;
                            sessionStorage.setItem('chatbot_session_id', sessionId);
                        }
                        
                        // 显示追问建议
                        if (data.suggestions && data.suggestions.length > 0) {
                            showQuickReplies(data.suggestions);
                        } else {
                            showDefaultQuickReplies();
                        }
                    } else {
                        addBotMessage('抱歉，我暂时无法处理您的请求。请稍后再试或拨打客服热线 400-888-8888。', 'error', true);
                        showDefaultQuickReplies();
                    }
                } catch(e) {
                    addBotMessage('抱歉，系统出现了错误。请拨打客服热线 400-888-8888 获取帮助。', 'error', true);
                    showDefaultQuickReplies();
                }
            } else if (xhr.status === 429) {
                addBotMessage('您发送消息太快了，请稍等片刻再试 😊', 'error');
                showDefaultQuickReplies();
            } else {
                addBotMessage('网络连接异常，请检查网络后重试。客服热线：400-888-8888', 'error', true);
                showDefaultQuickReplies();
            }
        };
        
        xhr.onerror = function() {
            hideTyping();
            isWaiting = false;
            sendBtn.disabled = false;
            addBotMessage('网络连接失败，请检查网络后重试。客服热线：400-888-8888', 'error', true);
            showDefaultQuickReplies();
        };
        
        // 构建请求体
        var params = 'message=' + encodeURIComponent(text) + '&session_id=' + encodeURIComponent(sessionId);
        xhr.send(params);
    }
    
    // 初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
    
    // 暴露函数到全局
    window.chatbotToggleWindow = toggleWindow;
    window.chatbotSendMessage = sendMessage;
})();
</script>
<%
End If
%>
