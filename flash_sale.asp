<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<%
If Not FEATURE_FLASH_SALE Then Response.Redirect "/index.asp"

Call OpenConnection()

Dim page, pageSize, pageInfo, rsFlash
page = CLng(GetQueryParam("page", "1"))
pageSize = 12
If page < 1 Then page = 1

' 获取当前正在进行的秒杀活动
Dim sqlFlash, sqlCount, totalCount, totalPages
sqlFlash = "SELECT fs.FlashSaleID, fs.ProductID, fs.FlashPrice, fs.Stock, fs.SoldCount, " & _
           "fs.LimitPerUser, fs.StartTime, fs.EndTime, fs.IsActive, " & _
           "p.ProductName, p.ImageURL, p.BasePrice, p.Category, p.Description " & _
           "FROM FlashSale fs " & _
           "INNER JOIN Products p ON fs.ProductID = p.ProductID " & _
           "WHERE fs.IsActive = 1 AND GETDATE() >= fs.StartTime AND GETDATE() <= fs.EndTime AND fs.Stock > fs.SoldCount " & _
           "ORDER BY fs.SortOrder ASC, fs.EndTime ASC"

' 分页
sqlCount = "SELECT COUNT(*) FROM (" & sqlFlash & ") AS cnt"
totalCount = CLng(DAL_GetScalar(sqlCount, Null, 0))
totalPages = Int((totalCount + pageSize - 1) / pageSize)
If totalPages < 1 Then totalPages = 1
If page > totalPages Then page = totalPages

Dim offset : offset = (page - 1) * pageSize
sqlFlash = sqlFlash & " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"

Set rsFlash = DAL_GetList(sqlFlash, Null)

' 即将开始的秒杀
Dim sqlUpcoming, rsUpcoming
sqlUpcoming = "SELECT TOP 6 fs.FlashSaleID, fs.FlashPrice, fs.Stock, fs.StartTime, fs.EndTime, " & _
              "p.ProductName, p.ImageURL, p.BasePrice " & _
              "FROM FlashSale fs " & _
              "INNER JOIN Products p ON fs.ProductID = p.ProductID " & _
              "WHERE fs.IsActive = 1 AND GETDATE() < fs.StartTime AND fs.Stock > 0 " & _
              "ORDER BY fs.StartTime ASC"
Set rsUpcoming = DAL_GetList(sqlUpcoming, Null)

Function GetQueryParam(name, defaultVal)
    Dim v : v = Request.QueryString(name)
    If v = "" Or IsNull(v) Then v = defaultVal
    GetQueryParam = v
End Function
%>
<!--#include file="includes/header.asp"-->

<section class="page-hero flash-sale-hero">
    <div class="container">
        <div class="hero-content text-center">
            <div class="flash-icon"><i class="fas fa-bolt"></i></div>
            <h1>限时秒杀</h1>
            <p>超值好物限时抢购，手慢无！</p>
        </div>
    </div>
</section>

<%
' 进行中的秒杀
%>
<section class="flash-sale-section">
    <div class="container">
        <div class="section-header">
            <h2><i class="fas fa-fire"></i> 进行中</h2>
            <span class="section-count"><%= totalCount %> 件秒杀商品</span>
        </div>
        
        <% If rsFlash Is Nothing Or rsFlash.EOF Then %>
        <div class="empty-state">
            <i class="fas fa-clock"></i>
            <p>当前没有正在进行的秒杀活动</p>
            <p class="text-muted">看看即将开始的秒杀吧</p>
        </div>
        <% Else %>
        <div class="flash-grid">
            <%
            Do While Not rsFlash.EOF
                Dim fsID, fsProductID, fsPrice, fsStock, fsSold, fsLimit, fsEnd, fsProductName, fsImage, fsBasePrice
                fsID = rsFlash("FlashSaleID")
                fsProductID = rsFlash("ProductID")
                fsPrice = rsFlash("FlashPrice")
                fsStock = rsFlash("Stock")
                fsSold = rsFlash("SoldCount")
                fsLimit = rsFlash("LimitPerUser")
                fsEnd = rsFlash("EndTime")
                fsProductName = rsFlash("ProductName")
                fsImage = rsFlash("ImageURL")
                fsBasePrice = rsFlash("BasePrice")
                
                If IsNull(fsImage) Or fsImage = "" Then fsImage = DEFAULT_PRODUCT_IMAGE
                
                Dim remainStock : remainStock = fsStock - fsSold
                If remainStock < 0 Then remainStock = 0
                Dim progressPercent : progressPercent = Int((fsSold / fsStock) * 100)
                If progressPercent > 100 Then progressPercent = 100
                
                Dim discountPercent : discountPercent = 0
                If CDbl(fsBasePrice) > 0 Then
                    discountPercent = Int((1 - CDbl(fsPrice) / CDbl(fsBasePrice)) * 100)
                End If
            %>
            <div class="flash-card" data-endtime="<%= FormatDateTime(fsEnd, 0) %>">
                <a href="/product.asp?id=<%= fsProductID %>&flash=<%= fsID %>" class="flash-card-link">
                    <div class="flash-card-image">
                        <img src="<%= fsImage %>" alt="<%= Server.HTMLEncode(fsProductName) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                        <div class="flash-badge">秒杀</div>
                        <% If discountPercent >= 30 Then %>
                        <div class="discount-tag">-<%= discountPercent %>%</div>
                        <% End If %>
                    </div>
                    <div class="flash-card-body">
                        <h3 class="flash-card-title"><%= Server.HTMLEncode(fsProductName) %></h3>
                        <div class="flash-price-row">
                            <span class="flash-price">&yen;<%= FormatNumber(fsPrice, 2) %></span>
                            <span class="flash-original-price">&yen;<%= FormatNumber(fsBasePrice, 2) %></span>
                        </div>
                        <div class="flash-progress">
                            <div class="progress-bar">
                                <div class="progress-fill" style="width:<%= progressPercent %>%"></div>
                            </div>
                            <div class="progress-text">
                                <span>已抢 <%= progressPercent %>%</span>
                                <span>剩余 <%= remainStock %> 件</span>
                            </div>
                        </div>
                        <div class="flash-countdown" data-end="<%= FormatDateTime(fsEnd, 0) %>">
                            <span class="countdown-label">距离结束</span>
                            <span class="countdown-timer" id="timer_<%= fsID %>">
                                <span class="t-d">00</span>:<span class="t-h">00</span>:<span class="t-m">00</span>:<span class="t-s">00</span>
                            </span>
                        </div>
                        <button class="btn btn-danger btn-block flash-btn">立即抢购</button>
                    </div>
                </a>
            </div>
            <%
                rsFlash.MoveNext
            Loop
            %>
        </div>
        
        <% ' 分页 %>
        <% If totalPages > 1 Then %>
        <div class="pagination">
            <% If page > 1 Then %><a href="?page=<%= page - 1 %>" class="page-link">&laquo; 上一页</a><% End If %>
            <span class="page-info"><%= page %> / <%= totalPages %></span>
            <% If page < totalPages Then %><a href="?page=<%= page + 1 %>" class="page-link">下一页 &raquo;</a><% End If %>
        </div>
        <% End If %>
        <% End If %>
    </div>
</section>

<% ' 即将开始的秒杀 %>
<% If Not rsUpcoming Is Nothing And Not rsUpcoming.EOF Then %>
<section class="upcoming-section">
    <div class="container">
        <div class="section-header">
            <h2><i class="fas fa-clock"></i> 即将开始</h2>
        </div>
        <div class="upcoming-grid">
            <%
            Do While Not rsUpcoming.EOF
                Dim uID, uPrice, uStock, uStart, uEnd, uName, uImage, uBase
                uID = rsUpcoming("FlashSaleID")
                uPrice = rsUpcoming("FlashPrice")
                uStock = rsUpcoming("Stock")
                uStart = rsUpcoming("StartTime")
                uEnd = rsUpcoming("EndTime")
                uName = rsUpcoming("ProductName")
                uImage = rsUpcoming("ImageURL")
                uBase = rsUpcoming("BasePrice")
                If IsNull(uImage) Or uImage = "" Then uImage = DEFAULT_PRODUCT_IMAGE
            %>
            <div class="upcoming-card">
                <div class="upcoming-image">
                    <img src="<%= uImage %>" alt="<%= Server.HTMLEncode(uName) %>" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
                </div>
                <div class="upcoming-info">
                    <h4><%= Server.HTMLEncode(uName) %></h4>
                    <div class="upcoming-price">
                        <span class="flash-price">&yen;<%= FormatNumber(uPrice, 2) %></span>
                        <span class="flash-original-price">&yen;<%= FormatNumber(uBase, 2) %></span>
                    </div>
                    <div class="upcoming-start">
                        <i class="fas fa-clock"></i>
                        <span><%= FormatDateTime(uStart, 0) %> 开始</span>
                    </div>
                </div>
            </div>
            <%
                rsUpcoming.MoveNext
            Loop
            %>
        </div>
    </div>
</section>
<%
If Not rsUpcoming Is Nothing Then rsUpcoming.Close: Set rsUpcoming = Nothing
End If
%>

<style nonce="<%= Session("csp_nonce") %>">
.flash-sale-hero {
    background: linear-gradient(135deg, #ff416c 0%, #ff4b2b 100%);
    color: #fff;
    padding: 60px 0 40px;
    text-align: center;
}
.flash-sale-hero .flash-icon {
    font-size: 3rem;
    margin-bottom: 10px;
    animation: flashPulse 1s ease-in-out infinite;
}
@keyframes flashPulse {
    0%, 100% { opacity: 1; transform: scale(1); }
    50% { opacity: 0.6; transform: scale(1.1); }
}
.flash-sale-hero h1 { font-size: 2.5rem; margin: 10px 0; }
.flash-sale-hero p { font-size: 1.1rem; opacity: 0.9; }

.flash-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(270px, 1fr));
    gap: 20px;
}
.flash-card {
    background: #fff;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    transition: transform 0.2s, box-shadow 0.2s;
    position: relative;
}
.flash-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.15);
}
.flash-card-link { text-decoration: none; color: inherit; display: block; }
.flash-card-image {
    position: relative;
    padding-top: 100%;
    overflow: hidden;
    background: #f9f9f9;
}
.flash-card-image img {
    position: absolute; top: 0; left: 0;
    width: 100%; height: 100%;
    object-fit: cover;
    transition: transform 0.3s;
}
.flash-card:hover .flash-card-image img { transform: scale(1.05); }
.flash-badge {
    position: absolute; top: 10px; left: 10px;
    background: linear-gradient(135deg, #ff416c, #ff4b2b);
    color: #fff; padding: 4px 10px; border-radius: 4px;
    font-size: 12px; font-weight: 700;
}
.discount-tag {
    position: absolute; top: 10px; right: 10px;
    background: #000; color: #fff; padding: 4px 8px; border-radius: 4px;
    font-size: 12px; font-weight: 700; opacity: 0.8;
}
.flash-card-body { padding: 15px; }
.flash-card-title {
    font-size: 15px; margin: 0 0 10px;
    display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical;
    overflow: hidden; line-height: 1.4; min-height: 42px;
}
.flash-price-row { margin-bottom: 10px; }
.flash-price { color: #ff4b2b; font-size: 22px; font-weight: 700; }
.flash-original-price { color: #999; font-size: 13px; text-decoration: line-through; margin-left: 8px; }
.flash-progress { margin-bottom: 12px; }
.progress-bar {
    height: 6px; background: #f0f0f0; border-radius: 3px; overflow: hidden; margin-bottom: 4px;
}
.progress-fill {
    height: 100%; background: linear-gradient(90deg, #ff416c, #ff4b2b);
    border-radius: 3px; transition: width 0.5s;
}
.progress-text {
    display: flex; justify-content: space-between;
    font-size: 12px; color: #999;
}
.flash-countdown { margin-bottom: 12px; text-align: center; }
.countdown-label { font-size: 12px; color: #999; display: block; margin-bottom: 4px; }
.countdown-timer {
    font-size: 18px; font-weight: 700; color: #333; font-family: 'Courier New', monospace;
}
.countdown-timer span { background: #f5f5f5; padding: 2px 6px; border-radius: 3px; }
.flash-btn { margin-top: 5px; }

/* 即将开始 */
.upcoming-section { padding: 40px 0; background: #fafafa; }
.upcoming-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
    gap: 16px;
}
.upcoming-card {
    background: #fff; border-radius: 10px; padding: 12px;
    display: flex; gap: 12px; align-items: center;
    box-shadow: 0 1px 6px rgba(0,0,0,0.06);
}
.upcoming-image {
    width: 70px; height: 70px; flex-shrink: 0;
    border-radius: 8px; overflow: hidden; background: #f9f9f9;
}
.upcoming-image img { width: 100%; height: 100%; object-fit: cover; }
.upcoming-info { flex: 1; min-width: 0; }
.upcoming-info h4 {
    font-size: 13px; margin: 0 0 6px;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.upcoming-price { margin-bottom: 4px; }
.upcoming-price .flash-price { font-size: 16px; }
.upcoming-price .flash-original-price { font-size: 12px; }
.upcoming-start { font-size: 12px; color: #999; }
.upcoming-start i { margin-right: 4px; }

.empty-state { text-align: center; padding: 60px 20px; color: #999; }
.empty-state i { font-size: 3rem; margin-bottom: 15px; display: block; }
.empty-state p { margin: 5px 0; }
.empty-state .text-muted { font-size: 14px; }
</style>

<script nonce="<%= Session("csp_nonce") %>">
// 倒计时功能
function updateCountdowns() {
    var now = new Date().getTime();
    document.querySelectorAll('.countdown-timer').forEach(function(el) {
        var card = el.closest('.flash-card');
        var endStr = card ? card.getAttribute('data-endtime') : null;
        if (!endStr) return;
        var end = new Date(endStr).getTime();
        var diff = end - now;
        if (diff <= 0) {
            el.innerHTML = '<span style="color:#999">已结束</span>';
            var btn = card.querySelector('.flash-btn');
            if (btn) { btn.textContent = '已结束'; btn.disabled = true; btn.classList.add('disabled'); }
            return;
        }
        var d = Math.floor(diff / 86400000);
        var h = Math.floor((diff % 86400000) / 3600000);
        var m = Math.floor((diff % 3600000) / 60000);
        var s = Math.floor((diff % 60000) / 1000);
        el.innerHTML = '<span class="t-d">' + pad(d) + '</span>:<span class="t-h">' + pad(h) + '</span>:<span class="t-m">' + pad(m) + '</span>:<span class="t-s">' + pad(s) + '</span>';
    });
}
function pad(n) { return n < 10 ? '0' + n : n; }
setInterval(updateCountdowns, 1000);
updateCountdowns();
</script>

<!--#include file="includes/footer.asp"-->
<%
If Not rsFlash Is Nothing Then
    If rsFlash.State = 1 Then rsFlash.Close
    Set rsFlash = Nothing
End If
Call CloseConnection()
%>
