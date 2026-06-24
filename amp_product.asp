<%@ Language=VBScript CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
' V13.1 AMP产品页 - 先查询数据再输出HTML
Call OpenConnection()

Dim productId, rsProduct, productType, typeDisplayName
productId = Request.QueryString("id")

If productId = "" Or Not IsNumeric(productId) Then
    productId = 1
End If

' 获取产品数据
Set rsProduct = ExecuteQuery("SELECT * FROM Products WHERE ProductID = " & CLng(productId) & " AND IsActive <> 0")
If rsProduct Is Nothing Or rsProduct.EOF Then
    Response.Status = "404 Not Found"
    Response.Write "<!doctype html><html><head><meta charset=""utf-8""><title>产品不存在</title></head><body>"
    Response.Write "<h1>产品不存在</h1><p>该产品可能已下架</p>"
    Response.Write "<a href=""/products.asp"">返回产品列表</a>"
    Response.Write "</body></html>"
    Response.End
End If

' 安全获取字段值（防止Null类型不匹配）
Dim pName, pDesc, pImage, pCategory, pBasePrice, pProductType

pName = rsProduct("ProductName") & ""
pDesc = rsProduct("Description") & ""
pImage = rsProduct("ImageURL") & ""
pCategory = rsProduct("Category") & ""

' 处理价格（BasePrice字段）
On Error Resume Next
pBasePrice = CDbl(rsProduct("BasePrice"))
If Err.Number <> 0 Then pBasePrice = 0
On Error GoTo 0

' 处理产品类型
pProductType = rsProduct("ProductType") & ""
If pProductType = "" Then pProductType = "Custom"

' 安全HTMLEncode包装
Function SafeEncode(val)
    If IsNull(val) Or val = "" Then
        SafeEncode = ""
    Else
        SafeEncode = Server.HTMLEncode(CStr(val))
    End If
End Function

' 截断描述（用于meta description）
Dim shortDesc
shortDesc = Replace(Replace(pDesc, vbCrLf, " "), vbLf, " ")
If Len(shortDesc) > 160 Then shortDesc = Left(shortDesc, 157) & "..."
%>
<!doctype html>
<html amp lang="zh-CN">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,minimum-scale=1,initial-scale=1">
    <title><%= SafeEncode(pName) %> - <%= SITE_NAME %></title>
    
    <!-- AMP必须 -->
    <link rel="canonical" href="http://<%= Request.ServerVariables("SERVER_NAME") %>/product.asp?id=<%= productId %>">
    <meta name="description" content="<%= SafeEncode(shortDesc) %>">
    
    <!-- AMP运行时 -->
    <script async src="https://cdn.ampproject.org/v0.js"></script>
    
    <!-- AMP组件 -->
    <script async custom-element="amp-carousel" src="https://cdn.ampproject.org/v0/amp-carousel-0.2.js"></script>
    <script async custom-element="amp-accordion" src="https://cdn.ampproject.org/v0/amp-accordion-0.1.js"></script>
    
    <!-- AMP Boilerplate -->
    <style amp-boilerplate>body{-webkit-animation:-amp-start 8s steps(1,end) 0s 1 normal both;-moz-animation:-amp-start 8s steps(1,end) 0s 1 normal both;animation:-amp-start 8s steps(1,end) 0s 1 normal both}@-webkit-keyframes -amp-start{from{visibility:hidden}to{visibility:visible}}@-moz-keyframes -amp-start{from{visibility:hidden}to{visibility:visible}}@-ms-keyframes -amp-start{from{visibility:hidden}to{visibility:visible}}@-o-keyframes -amp-start{from{visibility:hidden}to{visibility:visible}}@keyframes -amp-start{from{visibility:hidden}to{visibility:visible}}</style><noscript><style amp-boilerplate>body{-webkit-animation:none;-moz-animation:none;-ms-animation:none;animation:none}</style></noscript>
    
    <!-- AMP自定义样式（<75KB限制） -->
    <style amp-custom>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; color: #333; background: #f8f9fa; }
        .amp-header { background: #fff; padding: 12px 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); position: sticky; top: 0; z-index: 100; display: flex; align-items: center; justify-content: space-between; }
        .amp-header h1 { font-size: 18px; color: #8B4513; margin: 0; }
        .amp-header a { color: #8B4513; text-decoration: none; font-size: 14px; }
        .amp-breadcrumb { padding: 8px 16px; font-size: 12px; color: #999; background: #fff; border-bottom: 1px solid #f0f0f0; }
        .amp-breadcrumb a { color: #8B4513; text-decoration: none; }
        .amp-product { background: #fff; padding: 16px; margin-bottom: 12px; }
        .amp-product-title { font-size: 22px; font-weight: 600; margin-bottom: 8px; color: #2c3e50; }
        .amp-product-category { display: inline-block; padding: 2px 10px; background: #f5e6d3; color: #8B4513; border-radius: 12px; font-size: 12px; margin-bottom: 12px; }
        .amp-product-price { font-size: 28px; color: #8B4513; font-weight: 700; margin-bottom: 4px; }
        .amp-price-note { font-size: 13px; color: #999; margin-bottom: 16px; }
        .amp-product-desc { font-size: 14px; color: #666; line-height: 1.8; margin-bottom: 16px; }
        .amp-specs { background: #fff; padding: 16px; margin-bottom: 12px; }
        .amp-specs h3 { font-size: 16px; margin-bottom: 12px; color: #2c3e50; border-bottom: 2px solid #8B4513; padding-bottom: 8px; }
        .amp-spec-item { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #f0f0f0; font-size: 14px; }
        .amp-spec-label { color: #666; }
        .amp-spec-value { color: #333; font-weight: 500; }
        .amp-desc-section { background: #fff; padding: 16px; margin-bottom: 12px; }
        .amp-desc-section h3 { font-size: 16px; margin-bottom: 12px; color: #2c3e50; border-bottom: 2px solid #8B4513; padding-bottom: 8px; }
        .amp-desc-section p { font-size: 14px; color: #666; line-height: 1.8; }
        .amp-cta { position: fixed; bottom: 0; left: 0; right: 0; background: #fff; padding: 12px 16px; box-shadow: 0 -2px 8px rgba(0,0,0,0.1); z-index: 100; display: flex; gap: 12px; }
        .amp-btn { flex: 1; padding: 14px; background: #8B4513; color: #fff; text-align: center; text-decoration: none; border-radius: 8px; font-size: 16px; font-weight: 600; }
        .amp-btn-outline { flex: 1; padding: 14px; background: #fff; color: #8B4513; border: 2px solid #8B4513; text-align: center; text-decoration: none; border-radius: 8px; font-size: 16px; font-weight: 600; }
        .amp-footer { background: #fff; padding: 16px; margin-bottom: 80px; text-align: center; font-size: 12px; color: #999; }
        .amp-footer a { color: #8B4513; text-decoration: none; }
    </style>
    
    <!-- 结构化数据 (JSON-LD) -->
    <script type="application/ld+json">
    {
        "@context": "https://schema.org",
        "@type": "Product",
        "name": "<%= SafeEncode(pName) %>",
        "image": "<%= SafeEncode(pImage) %>",
        "description": "<%= SafeEncode(Left(pDesc, 200)) %>",
        "category": "<%= SafeEncode(pCategory) %>",
        "offers": {
            "@type": "Offer",
            "price": "<%= FormatNumber(pBasePrice, 2) %>",
            "priceCurrency": "CNY",
            "availability": "https://schema.org/InStock",
            "url": "http://<%= Request.ServerVariables("SERVER_NAME") %>/product.asp?id=<%= productId %>"
        }
    }
    </script>
</head>
<body>
    <!-- 头部 -->
    <div class="amp-header">
        <a href="/index.asp">&larr; 返回</a>
        <h1><%= SITE_NAME %></h1>
        <a href="/products.asp">全部产品</a>
    </div>
    
    <!-- 面包屑 -->
    <div class="amp-breadcrumb">
        <a href="/">首页</a> &gt; <a href="/products.asp">全部香水</a> &gt; <a href="/products.asp?category=<%= Server.URLEncode(pCategory) %>"><%= SafeEncode(pCategory) %></a> &gt; <%= SafeEncode(pName) %>
    </div>
    
    <!-- 产品图片 -->
    <% If pImage <> "" Then %>
    <amp-carousel width="400" height="400" layout="responsive" type="slides" autoplay delay="4000">
        <amp-img src="<%= SafeEncode(pImage) %>" width="400" height="400" layout="responsive" alt="<%= SafeEncode(pName) %>"></amp-img>
    </amp-carousel>
    <% Else %>
    <div style="background:#f0f0f0;height:300px;display:flex;align-items:center;justify-content:center;color:#999;font-size:48px;">&#128247;</div>
    <% End If %>
    
    <!-- 产品信息 -->
    <div class="amp-product">
        <span class="amp-product-category"><%= SafeEncode(pCategory) %></span>
        <h2 class="amp-product-title"><%= SafeEncode(pName) %></h2>
        <div class="amp-product-price">&yen;<%= FormatNumber(pBasePrice, 2) %></div>
        <div class="amp-price-note">
            <% If pProductType = "standard" Then %>
                (固定价格)
            <% Else %>
                起 (根据定制选项价格会有所变化)
            <% End If %>
        </div>
        <% If pDesc <> "" Then %>
        <p class="amp-product-desc"><%= SafeEncode(pDesc) %></p>
        <% End If %>
    </div>
    
    <!-- 产品规格 -->
    <amp-accordion expand-single-section>
        <section expanded>
            <h3 style="padding:12px 16px;background:#fff;margin:0;font-size:16px;color:#2c3e50;border-bottom:2px solid #8B4513;cursor:pointer;">产品规格</h3>
            <div class="amp-specs">
                <div class="amp-spec-item">
                    <span class="amp-spec-label">产品类型</span>
                    <span class="amp-spec-value"><%
                        If pProductType = "standard" Then
                            Response.Write "固定品牌"
                        ElseIf pProductType = "KOL" Then
                            Response.Write "KOL推荐"
                        ElseIf pProductType = "Custom" Then
                            Response.Write "自定义"
                        Else
                            Response.Write SafeEncode(pProductType)
                        End If
                    %></span>
                </div>
                <div class="amp-spec-item">
                    <span class="amp-spec-label">香调分类</span>
                    <span class="amp-spec-value"><%= SafeEncode(pCategory) %></span>
                </div>
                <div class="amp-spec-item">
                    <span class="amp-spec-label">基准价格</span>
                    <span class="amp-spec-value">&yen;<%= FormatNumber(pBasePrice, 2) %></span>
                </div>
            </div>
        </section>
        
        <section>
            <h3 style="padding:12px 16px;background:#fff;margin:0;font-size:16px;color:#2c3e50;border-bottom:2px solid #8B4513;cursor:pointer;">使用建议</h3>
            <div class="amp-desc-section">
                <p>
                    &#8226; 喷洒于手腕、颈部、耳后等脉搏点<br>
                    &#8226; 距离皮肤15-20厘米喷洒<br>
                    &#8226; 避免阳光直射，存放于阴凉处<br>
                    &#8226; 开封后建议12个月内使用完毕
                </p>
            </div>
        </section>
    </amp-accordion>
    
    <!-- 购买按钮 -->
    <div class="amp-cta">
        <a href="/product.asp?id=<%= productId %>" class="amp-btn-outline">查看详情</a>
        <a href="/product.asp?id=<%= productId %>" class="amp-btn">立即定制</a>
    </div>
    
    <!-- 页脚 -->
    <div class="amp-footer">
        <p>&copy; 2026 <%= SITE_NAME %> 版权所有</p>
        <p><a href="/products.asp">全部产品</a> | <a href="/about.asp">关于我们</a> | <a href="/contact.asp">联系我们</a></p>
    </div>
</body>
</html>
<%
rsProduct.Close
Set rsProduct = Nothing
Call CloseConnection()
%>
