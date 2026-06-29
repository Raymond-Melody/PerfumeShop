<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()
%>
<!--#include file="includes/header.asp"-->

<!-- 香氛测试横幅 -->
<section class="quiz-hero">
    <div class="container">
        <div class="hero-content">
            <h1><i class="fas fa-magic"></i> 智能香氛匹配</h1>
            <p>回答几个简单问题，让 AI 为你找到最适合的香氛配方</p>
        </div>
    </div>
</section>

<div class="container">
    <div class="quiz-page">
        <!-- 进度条 -->
        <div class="quiz-progress">
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill" style="width: 0%"></div>
            </div>
            <div class="progress-steps">
                <span class="step active" data-step="1">风格</span>
                <span class="step" data-step="2">场合</span>
                <span class="step" data-step="3">季节</span>
                <span class="step" data-step="4">类型</span>
                <span class="step" data-step="5">浓度</span>
                <span class="step" data-step="6">预算</span>
            </div>
        </div>

        <!-- 题目区域 -->
        <div class="quiz-content" id="quizContent">
            <!-- 第1题：风格 -->
            <div class="quiz-question active" data-question="1" id="q1">
                <h2 class="question-title">Q1. 你更喜欢哪种风格的香水？</h2>
                <p class="question-desc">选择最吸引你的香调风格</p>
                <div class="option-grid">
                    <div class="option-card" data-value="floral" onclick="selectOption(this, 'q1')">
                        <div class="option-icon"><i class="fas fa-seedling"></i></div>
                        <div class="option-label">花香调</div>
                        <div class="option-desc">浪漫优雅</div>
                    </div>
                    <div class="option-card" data-value="fresh" onclick="selectOption(this, 'q1')">
                        <div class="option-icon"><i class="fas fa-water"></i></div>
                        <div class="option-label">清新调</div>
                        <div class="option-desc">干净自然</div>
                    </div>
                    <div class="option-card" data-value="woody" onclick="selectOption(this, 'q1')">
                        <div class="option-icon"><i class="fas fa-tree"></i></div>
                        <div class="option-label">木质调</div>
                        <div class="option-desc">沉稳大气</div>
                    </div>
                    <div class="option-card" data-value="oriental" onclick="selectOption(this, 'q1')">
                        <div class="option-icon"><i class="fas fa-moon"></i></div>
                        <div class="option-label">东方调</div>
                        <div class="option-desc">神秘性感</div>
                    </div>
                    <div class="option-card" data-value="citrus" onclick="selectOption(this, 'q1')">
                        <div class="option-icon"><i class="fas fa-sun"></i></div>
                        <div class="option-label">柑橘调</div>
                        <div class="option-desc">活力阳光</div>
                    </div>
                </div>
            </div>

            <!-- 第2题：场合 -->
            <div class="quiz-question" data-question="2" id="q2">
                <h2 class="question-title">Q2. 你主要在什么场合使用香水？</h2>
                <p class="question-desc">不同场合适合不同风格的香氛</p>
                <div class="option-grid">
                    <div class="option-card" data-value="daily" onclick="selectOption(this, 'q2')">
                        <div class="option-icon"><i class="fas fa-calendar-day"></i></div>
                        <div class="option-label">日常通勤</div>
                        <div class="option-desc">每天随身</div>
                    </div>
                    <div class="option-card" data-value="work" onclick="selectOption(this, 'q2')">
                        <div class="option-icon"><i class="fas fa-briefcase"></i></div>
                        <div class="option-label">工作办公</div>
                        <div class="option-desc">专业优雅</div>
                    </div>
                    <div class="option-card" data-value="date" onclick="selectOption(this, 'q2')">
                        <div class="option-icon"><i class="fas fa-heart"></i></div>
                        <div class="option-label">约会聚会</div>
                        <div class="option-desc">迷人魅力</div>
                    </div>
                    <div class="option-card" data-value="party" onclick="selectOption(this, 'q2')">
                        <div class="option-icon"><i class="fas fa-glass-cheers"></i></div>
                        <div class="option-label">派对晚宴</div>
                        <div class="option-desc">闪耀出众</div>
                    </div>
                    <div class="option-card" data-value="sport" onclick="selectOption(this, 'q2')">
                        <div class="option-icon"><i class="fas fa-running"></i></div>
                        <div class="option-label">运动休闲</div>
                        <div class="option-desc">活力清新</div>
                    </div>
                </div>
            </div>

            <!-- 第3题：季节 -->
            <div class="quiz-question" data-question="3" id="q3">
                <h2 class="question-title">Q3. 你最喜欢哪个季节？</h2>
                <p class="question-desc">季节偏好会影响香调选择</p>
                <div class="option-grid option-grid-4">
                    <div class="option-card" data-value="spring" onclick="selectOption(this, 'q3')">
                        <div class="option-icon"><i class="fas fa-flower"></i></div>
                        <div class="option-label">春天</div>
                        <div class="option-desc">万物复苏</div>
                    </div>
                    <div class="option-card" data-value="summer" onclick="selectOption(this, 'q3')">
                        <div class="option-icon"><i class="fas fa-umbrella-beach"></i></div>
                        <div class="option-label">夏天</div>
                        <div class="option-desc">热情洋溢</div>
                    </div>
                    <div class="option-card" data-value="autumn" onclick="selectOption(this, 'q3')">
                        <div class="option-icon"><i class="fas fa-leaf"></i></div>
                        <div class="option-label">秋天</div>
                        <div class="option-desc">温暖醇厚</div>
                    </div>
                    <div class="option-card" data-value="winter" onclick="selectOption(this, 'q3')">
                        <div class="option-icon"><i class="fas fa-snowflake"></i></div>
                        <div class="option-label">冬天</div>
                        <div class="option-desc">深邃温暖</div>
                    </div>
                </div>
            </div>

            <!-- 第4题：性别偏好 -->
            <div class="quiz-question" data-question="4" id="q4">
                <h2 class="question-title">Q4. 你偏好的香水类型？</h2>
                <p class="question-desc">帮助 AI 了解你的性别偏好</p>
                <div class="option-grid option-grid-3">
                    <div class="option-card" data-value="female" onclick="selectOption(this, 'q4')">
                        <div class="option-icon"><i class="fas fa-venus"></i></div>
                        <div class="option-label">女士香水</div>
                        <div class="option-desc">柔美花香</div>
                    </div>
                    <div class="option-card" data-value="male" onclick="selectOption(this, 'q4')">
                        <div class="option-icon"><i class="fas fa-mars"></i></div>
                        <div class="option-label">男士香水</div>
                        <div class="option-desc">沉稳大气</div>
                    </div>
                    <div class="option-card" data-value="unisex" onclick="selectOption(this, 'q4')">
                        <div class="option-icon"><i class="fas fa-genderless"></i></div>
                        <div class="option-label">中性香水</div>
                        <div class="option-desc">自由不拘</div>
                    </div>
                </div>
            </div>

            <!-- 第5题：浓度 -->
            <div class="quiz-question" data-question="5" id="q5">
                <h2 class="question-title">Q5. 你喜欢多浓的香水？</h2>
                <p class="question-desc">影响香水持久度和扩散力</p>
                <div class="option-grid option-grid-3">
                    <div class="option-card" data-value="light" onclick="selectOption(this, 'q5')">
                        <div class="option-icon"><i class="fas fa-feather-alt"></i></div>
                        <div class="option-label">淡雅清新</div>
                        <div class="option-desc">EDT 淡香水</div>
                    </div>
                    <div class="option-card" data-value="medium" onclick="selectOption(this, 'q5')">
                        <div class="option-icon"><i class="fas fa-balance-scale"></i></div>
                        <div class="option-label">适中持久</div>
                        <div class="option-desc">EDP 淡香精</div>
                    </div>
                    <div class="option-card" data-value="strong" onclick="selectOption(this, 'q5')">
                        <div class="option-icon"><i class="fas fa-fire"></i></div>
                        <div class="option-label">浓郁持久</div>
                        <div class="option-desc">Parfum 浓香精</div>
                    </div>
                </div>
            </div>

            <!-- 第6题：预算 -->
            <div class="quiz-question" data-question="6" id="q6">
                <h2 class="question-title">Q6. 你的预算范围？</h2>
                <p class="question-desc">帮助推荐合适价位的产品</p>
                <div class="option-grid option-grid-3">
                    <div class="option-card" data-value="entry" onclick="selectOption(this, 'q6')">
                        <div class="option-icon"><i class="fas fa-coins"></i></div>
                        <div class="option-label">入门级</div>
                        <div class="option-desc">200-500元</div>
                    </div>
                    <div class="option-card" data-value="mid" onclick="selectOption(this, 'q6')">
                        <div class="option-icon"><i class="fas fa-wallet"></i></div>
                        <div class="option-label">中端</div>
                        <div class="option-desc">500-1500元</div>
                    </div>
                    <div class="option-card" data-value="premium" onclick="selectOption(this, 'q6')">
                        <div class="option-icon"><i class="fas fa-gem"></i></div>
                        <div class="option-label">高端</div>
                        <div class="option-desc">1500元以上</div>
                    </div>
                </div>
            </div>

            <!-- 加载状态 -->
            <div class="quiz-loading" id="quizLoading" style="display:none;">
                <div class="loading-spinner">
                    <i class="fas fa-spinner fa-spin"></i>
                </div>
                <h3>AI 正在为你分析香氛偏好...</h3>
                <p>请稍候，正在匹配最适合你的香调组合</p>
            </div>

            <!-- 结果展示 -->
            <div class="quiz-result" id="quizResult" style="display:none;">
                <div class="result-header">
                    <i class="fas fa-check-circle"></i>
                    <h2>你的专属香氛画像</h2>
                </div>

                <!-- 匹配的香调家族 -->
                <div class="result-section" id="resultFamilies">
                    <h3><i class="fas fa-layer-group"></i> 匹配香调家族</h3>
                    <div class="family-bars" id="familyBars"></div>
                </div>

                <!-- 推荐香调组合 -->
                <div class="result-section">
                    <h3><i class="fas fa-flask"></i> 推荐香调组合</h3>
                    <div class="notes-pyramid" id="notesPyramid">
                        <div class="pyramid-layer pyramid-top">
                            <div class="layer-label">前调 Top</div>
                            <div class="layer-notes" id="topNotes"></div>
                        </div>
                        <div class="pyramid-layer pyramid-middle">
                            <div class="layer-label">中调 Heart</div>
                            <div class="layer-notes" id="middleNotes"></div>
                        </div>
                        <div class="pyramid-layer pyramid-base">
                            <div class="layer-label">后调 Base</div>
                            <div class="layer-notes" id="baseNotes"></div>
                        </div>
                    </div>
                </div>

                <!-- 强度建议 -->
                <div class="result-section">
                    <h3><i class="fas fa-lightbulb"></i> 专业建议</h3>
                    <div class="advice-box" id="intensityAdvice"></div>
                </div>

                <!-- CTA -->
                <div class="result-cta">
                    <a href="/customize.asp" class="btn btn-primary btn-lg">
                        <i class="fas fa-spray-can"></i> 开始定制你的香水
                    </a>
                    <button class="btn btn-outline btn-lg" onclick="resetQuiz()">
                        <i class="fas fa-redo"></i> 重新测试
                    </button>
                </div>
            </div>
        </div>
    </div>
</div>

<style>
.quiz-hero {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff;
    padding: 60px 0;
    text-align: center;
}
.quiz-hero h1 {
    font-size: 36px;
    margin-bottom: 12px;
}
.quiz-hero p {
    font-size: 16px;
    opacity: 0.9;
}
.quiz-page {
    padding: 40px 0 60px;
    max-width: 800px;
    margin: 0 auto;
}
.quiz-progress {
    margin-bottom: 40px;
}
.progress-bar {
    height: 6px;
    background: #e9ecef;
    border-radius: 3px;
    overflow: hidden;
    margin-bottom: 16px;
}
.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #667eea, #764ba2);
    border-radius: 3px;
    transition: width 0.4s ease;
}
.progress-steps {
    display: flex;
    justify-content: space-between;
}
.progress-steps .step {
    font-size: 13px;
    color: #adb5bd;
    transition: color 0.3s;
    position: relative;
}
.progress-steps .step.active {
    color: #667eea;
    font-weight: 600;
}
.progress-steps .step.done {
    color: #28a745;
}
.quiz-question {
    display: none;
}
.quiz-question.active {
    display: block;
    animation: quizFadeIn 0.4s ease;
}
@keyframes quizFadeIn {
    from { opacity: 0; transform: translateY(20px); }
    to { opacity: 1; transform: translateY(0); }
}
.question-title {
    font-size: 24px;
    margin-bottom: 8px;
    color: #2d3748;
}
.question-desc {
    color: #718096;
    margin-bottom: 30px;
    font-size: 15px;
}
.option-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
}
.option-grid-4 {
    grid-template-columns: repeat(4, 1fr);
}
.option-grid-3 {
    grid-template-columns: repeat(3, 1fr);
}
.option-card {
    background: #fff;
    border: 2px solid #e2e8f0;
    border-radius: 16px;
    padding: 24px 16px;
    text-align: center;
    cursor: pointer;
    transition: all 0.3s ease;
}
.option-card:hover {
    border-color: #667eea;
    transform: translateY(-3px);
    box-shadow: 0 8px 25px rgba(102,126,234,0.15);
}
.option-card.selected {
    border-color: #667eea;
    background: linear-gradient(135deg, #f0f0ff 0%, #e8e8ff 100%);
    box-shadow: 0 4px 15px rgba(102,126,234,0.2);
}
.option-icon {
    font-size: 36px;
    color: #667eea;
    margin-bottom: 12px;
}
.option-label {
    font-size: 16px;
    font-weight: 600;
    color: #2d3748;
    margin-bottom: 4px;
}
.option-desc {
    font-size: 13px;
    color: #a0aec0;
}
.quiz-loading {
    text-align: center;
    padding: 60px 0;
}
.loading-spinner {
    font-size: 48px;
    color: #667eea;
    margin-bottom: 20px;
}
.quiz-loading h3 {
    color: #2d3748;
    margin-bottom: 8px;
}
.quiz-loading p {
    color: #718096;
}
.quiz-result {
    animation: quizFadeIn 0.6s ease;
}
.result-header {
    text-align: center;
    margin-bottom: 30px;
}
.result-header i {
    font-size: 48px;
    color: #28a745;
    margin-bottom: 12px;
}
.result-header h2 {
    font-size: 26px;
    color: #2d3748;
}
.result-section {
    background: #fff;
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.result-section h3 {
    font-size: 18px;
    color: #2d3748;
    margin-bottom: 16px;
}
.result-section h3 i {
    color: #667eea;
    margin-right: 8px;
}
.family-bar-item {
    display: flex;
    align-items: center;
    margin-bottom: 12px;
}
.family-bar-item:last-child {
    margin-bottom: 0;
}
.family-bar-label {
    width: 80px;
    font-size: 14px;
    font-weight: 600;
    color: #4a5568;
    flex-shrink: 0;
}
.family-bar-track {
    flex: 1;
    height: 24px;
    background: #e9ecef;
    border-radius: 12px;
    overflow: hidden;
    position: relative;
}
.family-bar-fill {
    height: 100%;
    border-radius: 12px;
    background: linear-gradient(90deg, #667eea, #764ba2);
    transition: width 0.8s ease;
    display: flex;
    align-items: center;
    padding-left: 10px;
}
.family-bar-score {
    font-size: 12px;
    color: #fff;
    font-weight: 600;
}
.family-bar-keywords {
    margin-left: 12px;
    font-size: 12px;
    color: #a0aec0;
    flex-shrink: 0;
}
.notes-pyramid {
    display: flex;
    flex-direction: column;
    gap: 12px;
    align-items: center;
}
.pyramid-layer {
    text-align: center;
    padding: 16px 20px;
    border-radius: 12px;
    border-left: 4px solid;
}
.pyramid-top {
    background: #fff9e6;
    border-color: #ffd700;
    width: 70%;
}
.pyramid-middle {
    background: #fff0f5;
    border-color: #ff69b4;
    width: 85%;
}
.pyramid-base {
    background: #f5f0eb;
    border-color: #8b4513;
    width: 100%;
}
.layer-label {
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 8px;
}
.pyramid-top .layer-label { color: #b8860b; }
.pyramid-middle .layer-label { color: #c71585; }
.pyramid-base .layer-label { color: #654321; }
.layer-notes {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    justify-content: center;
}
.layer-notes .note-tag {
    display: inline-block;
    padding: 4px 14px;
    background: rgba(255,255,255,0.8);
    border-radius: 20px;
    font-size: 14px;
    color: #4a5568;
}
.advice-box {
    background: linear-gradient(135deg, #f0f0ff 0%, #faf5ff 100%);
    border-radius: 12px;
    padding: 20px;
    font-size: 15px;
    color: #4a5568;
    line-height: 1.8;
}
.result-cta {
    text-align: center;
    margin-top: 30px;
    display: flex;
    gap: 16px;
    justify-content: center;
    flex-wrap: wrap;
}
.result-cta .btn-lg {
    padding: 14px 32px;
    font-size: 16px;
    border-radius: 12px;
}
.result-cta .btn-outline {
    background: #fff;
    border: 2px solid #667eea;
    color: #667eea;
}
.result-cta .btn-outline:hover {
    background: #667eea;
    color: #fff;
}
@media (max-width: 768px) {
    .quiz-page { padding: 20px 16px 40px; }
    .option-grid { grid-template-columns: repeat(2, 1fr); gap: 12px; }
    .option-grid-4 { grid-template-columns: repeat(2, 1fr); }
    .option-grid-3 { grid-template-columns: repeat(2, 1fr); }
    .option-card { padding: 18px 12px; }
    .option-icon { font-size: 28px; }
    .question-title { font-size: 20px; }
    .pyramid-top { width: 85%; }
    .pyramid-middle { width: 92%; }
    .family-bar-keywords { display: none; }
}
@media (max-width: 480px) {
    .option-grid, .option-grid-3, .option-grid-4 { grid-template-columns: 1fr 1fr; }
    .progress-steps .step { font-size: 11px; }
}
</style>

<script>
var quizAnswers = {};
var currentStep = 1;
var totalSteps = 6;

function selectOption(card, questionId) {
    // 移除同题的选中状态
    var qEl = document.getElementById(questionId);
    var cards = qEl.querySelectorAll('.option-card');
    cards.forEach(function(c) { c.classList.remove('selected'); });
    
    // 标记选中
    card.classList.add('selected');
    
    // 保存答案
    var qNum = qEl.getAttribute('data-question');
    quizAnswers[getQuestionKey(qNum)] = card.getAttribute('data-value');
    
    // 自动进入下一题
    setTimeout(function() {
        if (parseInt(qNum) < totalSteps) {
            goToStep(parseInt(qNum) + 1);
        } else {
            submitQuiz();
        }
    }, 400);
}

function getQuestionKey(stepNum) {
    var keys = ['style', 'occasion', 'season', 'gender', 'intensity', 'budget'];
    return keys[parseInt(stepNum) - 1];
}

function goToStep(step) {
    // 隐藏所有问题
    var allQuestions = document.querySelectorAll('.quiz-question');
    allQuestions.forEach(function(q) { q.classList.remove('active'); });
    
    // 显示当前问题
    var targetQ = document.getElementById('q' + step);
    if (targetQ) {
        targetQ.classList.add('active');
        targetQ.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
    
    // 更新进度条
    currentStep = step;
    updateProgress();
}

function updateProgress() {
    var pct = Math.round((currentStep - 1) / totalSteps * 100);
    document.getElementById('progressFill').style.width = pct + '%';
    
    // 更新步骤指示器
    var steps = document.querySelectorAll('.progress-steps .step');
    steps.forEach(function(s) {
        var sNum = parseInt(s.getAttribute('data-step'));
        s.classList.remove('active', 'done');
        if (sNum === currentStep) {
            s.classList.add('active');
        } else if (sNum < currentStep) {
            s.classList.add('done');
        }
    });
}

function submitQuiz() {
    // 检查是否全部答完
    if (Object.keys(quizAnswers).length < totalSteps) {
        return;
    }
    
    // 隐藏所有问题，显示加载状态
    var allQuestions = document.querySelectorAll('.quiz-question');
    allQuestions.forEach(function(q) { q.classList.remove('active'); });
    document.getElementById('quizLoading').style.display = 'block';
    document.getElementById('progressFill').style.width = '100%';
    document.getElementById('quizLoading').scrollIntoView({ behavior: 'smooth' });
    
    // 发送 AJAX 请求
    $.ajax({
        url: '/api/fragrance_match.asp',
        method: 'POST',
        contentType: 'application/json',
        data: JSON.stringify({ answers: quizAnswers }),
        success: function(res) {
            if (res.code === 0 && res.data) {
                showResult(res.data);
            } else {
                showError(res.message || '匹配失败，请重试');
            }
        },
        error: function() {
            showError('网络错误，请检查连接后重试');
        }
    });
}

function showResult(data) {
    document.getElementById('quizLoading').style.display = 'none';
    document.getElementById('quizResult').style.display = 'block';
    
    // 渲染匹配家族
    var familiesHtml = '';
    if (data.matched_families) {
        var families = Array.isArray(data.matched_families) 
            ? data.matched_families 
            : Object.values(data.matched_families);
        
        families.forEach(function(fam) {
            var pct = Math.min(Math.round(fam.score * 30), 100);
            var famNames = {
                'floral': '花香调', 'citrus': '柑橘调', 'woody': '木质调',
                'oriental': '东方调', 'fresh': '清新调', 'fruity': '果香调', 'green': '绿叶调'
            };
            var keywords = Array.isArray(fam.keywords) ? fam.keywords.slice(0,3).join('·') : '';
            familiesHtml += 
                '<div class="family-bar-item">' +
                '<span class="family-bar-label">' + (famNames[fam.family] || fam.family) + '</span>' +
                '<div class="family-bar-track">' +
                '<div class="family-bar-fill" style="width:' + pct + '%">' +
                '<span class="family-bar-score">' + fam.score.toFixed(1) + '</span>' +
                '</div></div>' +
                '<span class="family-bar-keywords">' + keywords + '</span>' +
                '</div>';
        });
    }
    document.getElementById('familyBars').innerHTML = familiesHtml;
    
    // 渲染推荐香调
    if (data.recommended_notes) {
        var notes = data.recommended_notes;
        renderNotes('topNotes', notes.top);
        renderNotes('middleNotes', notes.middle);
        renderNotes('baseNotes', notes.base);
    }
    
    // 渲染强度建议
    if (data.intensity_advice) {
        document.getElementById('intensityAdvice').textContent = data.intensity_advice;
    }
    
    // 滚动到结果
    document.getElementById('quizResult').scrollIntoView({ behavior: 'smooth' });
}

function renderNotes(containerId, notesArr) {
    var container = document.getElementById(containerId);
    if (!container) return;
    
    var notes = Array.isArray(notesArr) ? notesArr : (notesArr ? Object.values(notesArr) : []);
    if (notes.length === 0) {
        container.innerHTML = '<span class="note-tag">—</span>';
        return;
    }
    
    var html = '';
    notes.forEach(function(n) {
        if (n) html += '<span class="note-tag">' + n + '</span>';
    });
    container.innerHTML = html;
}

function showError(msg) {
    document.getElementById('quizLoading').style.display = 'none';
    document.getElementById('quizResult').style.display = 'block';
    document.getElementById('familyBars').innerHTML = 
        '<div style="text-align:center;padding:30px;color:#e53e3e;">' +
        '<i class="fas fa-exclamation-triangle" style="font-size:40px;margin-bottom:12px;"></i>' +
        '<p>' + msg + '</p>' +
        '<button class="btn btn-outline" onclick="resetQuiz()" style="margin-top:16px;">重新测试</button>' +
        '</div>';
}

function resetQuiz() {
    quizAnswers = {};
    currentStep = 1;
    updateProgress();
    document.querySelectorAll('.option-card').forEach(function(c) { c.classList.remove('selected'); });
    document.getElementById('quizResult').style.display = 'none';
    goToStep(1);
}

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    goToStep(1);
});
</script>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
