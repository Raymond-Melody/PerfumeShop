<%
' ============================================
' 社交分享工具 - Share Utils
' 生成各大社交平台分享链接
' ============================================

' ============================================
' URL编码（VBScript实现）
' ============================================
Function URLEncode(str)
    Dim i, ch, result, charCode
    result = ""
    If IsNull(str) Or IsEmpty(str) Then URLEncode = "" : Exit Function
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "_" Or ch = "." Or ch = "~" Then
            result = result & ch
        ElseIf ch = " " Then
            result = result & "+"
        Else
            ' AscW 对中文字符(>0x7FFF)返回负数, 需转为正数
            charCode = AscW(ch)
            If charCode < 0 Then charCode = charCode + 65536
            result = result & "%" & UCase(Hex(charCode))
        End If
    Next
    URLEncode = result
End Function

' ============================================
' 生成分享HTML代码
' ============================================
Function SU_RenderShareButtons(pageUrl, pageTitle, pageDesc, pageImage)
    Dim encodedUrl, encodedTitle, encodedDesc, encodedImage
    
    If IsNull(pageUrl) Or pageUrl = "" Then pageUrl = Request.ServerVariables("URL")
    If IsNull(pageTitle) Or pageTitle = "" Then pageTitle = SITE_NAME
    If IsNull(pageDesc) Or pageDesc = "" Then pageDesc = ""
    If IsNull(pageImage) Or pageImage = "" Then pageImage = ""
    
    encodedUrl = URLEncode(pageUrl)
    encodedTitle = URLEncode(pageTitle)
    encodedDesc = URLEncode(pageDesc)
    encodedImage = URLEncode(pageImage)
    
    Dim html
    html = ""
    
    ' 微信分享（使用二维码）
    html = html & "<a href=""#"" class=""share-btn share-wechat"" onclick=""showWechatQR('" & encodedUrl & "');return false;"" title=""分享到微信"">"
    html = html & "<i class=""fab fa-weixin""></i></a>" & vbCrLf
    
    ' 微博分享
    html = html & "<a href=""https://service.weibo.com/share/share.php?url=" & encodedUrl & "&title=" & encodedTitle & "&pic=" & encodedImage & """ target=""_blank"" class=""share-btn share-weibo"" title=""分享到微博"">"
    html = html & "<i class=""fab fa-weibo""></i></a>" & vbCrLf
    
    ' QQ空间
    html = html & "<a href=""https://sns.qzone.qq.com/cgi-bin/qzshare/cgi_qzshare_onekey?url=" & encodedUrl & "&title=" & encodedTitle & "&desc=" & encodedDesc & "&summary=" & encodedDesc & "&site=" & encodedUrl & """ target=""_blank"" class=""share-btn share-qzone"" title=""分享到QQ空间"">"
    html = html & "<i class=""fab fa-qq""></i></a>" & vbCrLf
    
    ' Facebook分享
    html = html & "<a href=""https://www.facebook.com/sharer/sharer.php?u=" & encodedUrl & "&quote=" & encodedTitle & """ target=""_blank"" class=""share-btn share-facebook"" title=""Share to Facebook"">"
    html = html & "<i class=""fab fa-facebook-f""></i></a>" & vbCrLf
    
    ' Twitter/X分享
    html = html & "<a href=""https://twitter.com/intent/tweet?text=" & encodedTitle & "&url=" & encodedUrl & """ target=""_blank"" class=""share-btn share-twitter"" title=""Share to Twitter"">"
    html = html & "<i class=""fab fa-twitter""></i></a>" & vbCrLf
    
    ' 复制链接
    html = html & "<a href=""#"" class=""share-btn share-copy"" onclick=""copyShareLink('" & Server.HTMLEncode(pageUrl) & "');return false;"" title=""复制链接"">"
    html = html & "<i class=""fas fa-link""></i></a>" & vbCrLf
    
    SU_RenderShareButtons = html
End Function

' ============================================
' 渲染分享区域（包含样式和脚本）
' ============================================
Sub SU_RenderShareSection(pageUrl, pageTitle, pageDesc, pageImage)
    ' 确保所有参数为字符串类型，避免 NVARCHAR(MAX) 等特殊字段类型不匹配
    pageUrl = pageUrl & ""
    If IsNull(pageTitle) Then pageTitle = "" Else pageTitle = pageTitle & ""
    If IsNull(pageDesc) Then pageDesc = "" Else pageDesc = pageDesc & ""
    If IsNull(pageImage) Then pageImage = "" Else pageImage = pageImage & ""
%>
<div class="share-section">
    <span class="share-label"><i class="fas fa-share-alt"></i> 分享：</span>
    <div class="share-buttons">
        <%= SU_RenderShareButtons(pageUrl, pageTitle, pageDesc, pageImage) %>
    </div>
</div>

<style>
.share-section {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 15px 0;
    margin: 15px 0;
    border-top: 1px solid #eee;
    border-bottom: 1px solid #eee;
}
.share-label {
    font-size: 14px;
    color: #666;
    white-space: nowrap;
}
.share-label i { color: #ff6f61; }
.share-buttons {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}
.share-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 36px;
    height: 36px;
    border-radius: 50%;
    color: #fff;
    font-size: 16px;
    text-decoration: none;
    transition: all 0.3s ease;
    opacity: 0.8;
}
.share-btn:hover { opacity: 1; transform: translateY(-2px); }
.share-wechat { background: #07C160; }
.share-weibo { background: #E6162D; }
.share-qzone { background: #FECE00; color: #333 !important; }
.share-facebook { background: #1877F2; }
.share-twitter { background: #1DA1F2; }
.share-copy { background: #666; }
</style>

<script>
function copyShareLink(url) {
    if (navigator.clipboard) {
        navigator.clipboard.writeText(url).then(function() {
            showShareToast('链接已复制到剪贴板');
        }).catch(function() {
            fallbackCopy(url);
        });
    } else {
        fallbackCopy(url);
    }
}
function fallbackCopy(url) {
    var textarea = document.createElement('textarea');
    textarea.value = url;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    showShareToast('链接已复制到剪贴板');
}
function showShareToast(msg) {
    var toast = document.createElement('div');
    toast.textContent = msg;
    toast.style.cssText = 'position:fixed;bottom:80px;left:50%;transform:translateX(-50%);background:rgba(0,0,0,0.8);color:#fff;padding:12px 24px;border-radius:8px;font-size:14px;z-index:9999;animation:shareFadeIn 0.3s;';
    document.body.appendChild(toast);
    setTimeout(function() { toast.style.opacity = '0'; toast.style.transition = 'opacity 0.5s'; setTimeout(function() { document.body.removeChild(toast); }, 500); }, 2000);
}
function showWechatQR(url) {
    var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=' + encodeURIComponent(url);
    var overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.6);display:flex;align-items:center;justify-content:center;z-index:10000;';
    overlay.innerHTML = '<div style="background:#fff;border-radius:12px;padding:25px;text-align:center;"><img src="' + qrUrl + '" alt="微信二维码" style="width:200px;height:200px;"><p style="margin-top:12px;color:#333;font-size:14px;">打开微信扫一扫分享</p><button onclick="this.parentElement.parentElement.remove()" style="margin-top:10px;padding:8px 20px;border:none;border-radius:6px;background:#07C160;color:#fff;">关闭</button></div>';
    document.body.appendChild(overlay);
}
</script>
<%
End Sub
%>