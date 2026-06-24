<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<!--#include file="includes/share_utils.asp"-->
<%
Call OpenConnection()

' V14: 会员登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Redirect "/user/login.asp?return=" & Server.URLEncode(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING"))
    Response.End
End If

Dim productId, rsProduct, productType, kolId, reviewStatus, typeDisplayName
productId = Request.QueryString("id")

If productId = "" Or Not IsNumeric(productId) Then
    Response.Redirect "/products.asp"
End If

' 先获取商品类型，以便在打开主查询前获取类型显示名称（避免MARS冲突）
Dim rsType
Set rsType = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = " & CInt(productId) & " AND IsActive <> 0")
If rsType Is Nothing Or rsType.EOF Then
    Response.Redirect "/products.asp"
End If
productType = LCase(rsType("ProductType") & "")
rsType.Close
Set rsType = Nothing

If productType = "" Then productType = "custom"

' 获取类型显示名称（此时没有其他Recordset打开）
typeDisplayName = GetProductTypeDisplayName(productType)

Set rsProduct = ExecuteQuery("SELECT * FROM Products WHERE ProductID = " & CInt(productId) & " AND IsActive <> 0")
If rsProduct Is Nothing Or rsProduct.EOF Then
    Response.Redirect "/products.asp"
End If

kolId = rsProduct("KOLID")
reviewStatus = rsProduct("ReviewStatus")

' 获取刻字配置
Dim productEngravable, productEngravingPrice, engravableRaw
engravableRaw = rsProduct("Engravable")
If IsNull(engravableRaw) Then
    productEngravable = False
ElseIf engravableRaw = True Or engravableRaw = 1 Then
    productEngravable = True
Else
    productEngravable = False
End If
productEngravingPrice = rsProduct("EngravingPrice")
If IsNull(productEngravingPrice) Then productEngravingPrice = 0 Else productEngravingPrice = CDbl(productEngravingPrice)

' 获取 KOL 预设比例
Dim dictKOLRatios
Set dictKOLRatios = Server.CreateObject("Scripting.Dictionary")
If productType = "kol" Then
    Dim rsKOL, kolPct
    Set rsKOL = ExecuteQuery("SELECT * FROM ProductNoteRatios WHERE ProductID = " & CInt(productId))
    If Not rsKOL Is Nothing Then
        Do While Not rsKOL.EOF
            ' DECIMAL类型需CDbl转换，否则输出到HTML时可能异常
            kolPct = rsKOL("Percentage")
            If IsNull(kolPct) Then kolPct = 0 Else kolPct = CDbl(kolPct)
            dictKOLRatios.Item(CStr(rsKOL("NoteID"))) = kolPct
            rsKOL.MoveNext
        Loop
        rsKOL.Close
    End If
End If

' 从SiteSettings获取香调最小比例配置
Dim minTopPercent, minMiddlePercent, minBasePercent
minTopPercent = 10
minMiddlePercent = 10
minBasePercent = 10

Dim rsMinPercent
Set rsMinPercent = ExecuteQuery("SELECT SettingKey, SettingValue FROM SiteSettings WHERE SettingKey IN ('MinTopPercent', 'MinMiddlePercent', 'MinBasePercent')")
If Not rsMinPercent Is Nothing Then
    Do While Not rsMinPercent.EOF
        Select Case rsMinPercent("SettingKey")
            Case "MinTopPercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minTopPercent = CInt(rsMinPercent("SettingValue"))
            Case "MinMiddlePercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minMiddlePercent = CInt(rsMinPercent("SettingValue"))
            Case "MinBasePercent"
                If IsNumeric(rsMinPercent("SettingValue")) Then minBasePercent = CInt(rsMinPercent("SettingValue"))
        End Select
        rsMinPercent.MoveNext
    Loop
    rsMinPercent.Close
End If
Set rsMinPercent = Nothing
%>
<%
' V13.1: AMP产品页链接（移动端加速）
Dim amphtmlLink, ampProtocol
If Request.ServerVariables("HTTPS") = "on" Then
    ampProtocol = "https://"
Else
    ampProtocol = "http://"
End If
amphtmlLink = "<link rel=""amphtml"" href=""" & ampProtocol & Request.ServerVariables("SERVER_NAME") & "/amp_product.asp?id=" & productId & """ />"
%>
<!--#include file="includes/header.asp"-->

<!-- 面包屑导航 -->
<div class="breadcrumb">
    <div class="container">
        <a href="/index.asp">首页</a>
        <span class="separator">/</span>
        <a href="/products.asp">全部香水</a>
        <span class="separator">/</span>
        <a href="/products.asp?category=<%= Server.URLEncode(rsProduct("Category") & "") %>"><%= HTMLEncode(rsProduct("Category") & "") %></a>
        <span class="separator">/</span>
        <span><%= HTMLEncode(rsProduct("ProductName")) %></span>
    </div>
</div>

<div class="container">
    <div class="product-detail">
        <!-- 产品图片 -->
        <div class="product-gallery">
            <div class="main-image">
                <img src="<%= rsProduct("ImageURL") %>" alt="<%= HTMLEncode(rsProduct("ProductName")) %>" id="mainImage" onerror="this.src='<%= DEFAULT_PRODUCT_IMAGE %>'">
            </div>
        </div>

        <!-- 产品信息 -->
        <div class="product-info-detail">
            <div class="product-header-tags">
                <span class="product-category-tag"><%= HTMLEncode(rsProduct("Category") & "") %></span>
                <% If typeDisplayName <> "" Then %>
                <span class="product-type-tag badge-<%= LCase(productType) %>"><%= HTMLEncode(typeDisplayName) %></span>
                <% End If %>
            </div>
            <h1><%= HTMLEncode(rsProduct("ProductName")) %></h1>
            <p class="product-description"><%= HTMLEncode(rsProduct("Description")) %></p>
            

            
            <div class="price-section">
                <span class="base-price" id="basePrice" data-price="<%= CDbl(rsProduct("BasePrice")) %>"><%= FormatMoney(rsProduct("BasePrice")) %></span>
                <span class="price-note" id="priceNote">
                    <% If productType = "standard" Then %>
                        (固定价格)
                    <% Else %>
                        起 (根据定制选项价格会有所变化)
                    <% End If %>
                </span>
            </div>

            <!-- 定制表单 -->
            <!-- 隐藏元素用于传递用户登录状态 -->
                <input type="hidden" id="userLoginStatus" value="<%= IIF(Session("UserID") <> "", "true", "false") %>" />
                
                <form id="customizeForm" class="customize-form" method="post">
                <input type="hidden" name="productId" value="<%= productId %>">
                <input type="hidden" name="productType" value="<%= productType %>">
                <%= GetCSRFTokenField() %>
                
                <% If productType = "custom" Then %>
                <!-- 选择前调 - 仅定制香水显示 -->
                <div class="option-section">
                    <h3><i class="fas fa-wind"></i> 前调 <span class="option-tip">香水的第一印象（最小 <%= minTopPercent %>%）</span></h3>
                    <div class="option-grid" id="topNotes">
                        <%
                        Dim rsNotes
                        Set rsNotes = ExecuteQuery("SELECT n.* FROM FragranceNotes n INNER JOIN ProductNotes pn ON n.NoteID = pn.NoteID WHERE pn.ProductID = " & CInt(productId) & " AND n.NoteType = '前调' AND n.IsActive <> 0")
                        If Not rsNotes Is Nothing Then
                            Do While Not rsNotes.EOF
                        %>
                        <div class="option-card-wrapper">
                            <label class="option-card">
                                <input type="checkbox" name="topNote" value="<%= rsNotes("NoteID") %>" 
                                    data-price="<%= CDbl(rsNotes("PriceAddition")) %>" 
                                    data-name="<%= HTMLEncode(rsNotes("NoteName")) %>" 
                                    onclick="toggleNote(this)">
                                <div class="option-content">
                                    <span class="option-name"><%= HTMLEncode(rsNotes("NoteName")) %></span>
                                    <span class="option-desc"><%= HTMLEncode(rsNotes("Description") & "") %></span>
                                    <% If CDbl(rsNotes("PriceAddition")) > 0 Then %>
                                    <span class="option-price">+<%= FormatMoney(rsNotes("PriceAddition")) %>/100%</span>
                                    <% End If %>
                                </div>
                            </label>
                            <div class="percent-input-wrapper" style="display:none; margin-top: 10px;">
                                <label>比例: <input type="number" name="percent_top_<%= rsNotes("NoteID") %>" class="note-percent" min="1" max="100" 
                                    value="<%= rsNotes("RecommendedPercentage") %>" 
                                    onchange="calculateTotal()"> %</label>
                            </div>
                        </div>
                        <%
                                rsNotes.MoveNext
                            Loop
                            rsNotes.Close
                        End If
                        %>
                    </div>
                </div>

                <!-- 选择中调 -->
                <div class="option-section">
                    <h3><i class="fas fa-heart"></i> 中调 <span class="option-tip">香水的核心灵魂（最小 <%= minMiddlePercent %>%）</span></h3>
                    <div class="option-grid" id="middleNotes">
                        <%
                        Set rsNotes = ExecuteQuery("SELECT n.* FROM FragranceNotes n INNER JOIN ProductNotes pn ON n.NoteID = pn.NoteID WHERE pn.ProductID = " & CInt(productId) & " AND n.NoteType = '中调' AND n.IsActive <> 0")
                        If Not rsNotes Is Nothing Then
                            Do While Not rsNotes.EOF
                        %>
                        <div class="option-card-wrapper">
                            <label class="option-card">
                                <input type="checkbox" name="middleNote" value="<%= rsNotes("NoteID") %>" 
                                    data-price="<%= CDbl(rsNotes("PriceAddition")) %>" 
                                    data-name="<%= HTMLEncode(rsNotes("NoteName")) %>" 
                                    onclick="toggleNote(this)">
                                <div class="option-content">
                                    <span class="option-name"><%= HTMLEncode(rsNotes("NoteName")) %></span>
                                    <span class="option-desc"><%= HTMLEncode(rsNotes("Description") & "") %></span>
                                    <% If CDbl(rsNotes("PriceAddition")) > 0 Then %>
                                    <span class="option-price">+<%= FormatMoney(rsNotes("PriceAddition")) %>/100%</span>
                                    <% End If %>
                                </div>
                            </label>
                            <div class="percent-input-wrapper" style="display:none; margin-top: 10px;">
                                <label>比例: <input type="number" name="percent_mid_<%= rsNotes("NoteID") %>" class="note-percent" min="1" max="100" 
                                    value="<%= rsNotes("RecommendedPercentage") %>" 
                                    onchange="calculateTotal()"> %</label>
                            </div>
                        </div>
                        <%
                                rsNotes.MoveNext
                            Loop
                            rsNotes.Close
                        End If
                        %>
                    </div>
                </div>

                <!-- 选择后调 -->
                <div class="option-section">
                    <h3><i class="fas fa-moon"></i> 后调 <span class="option-tip">持久的余韵（最小 <%= minBasePercent %>%）</span></h3>
                    <div class="option-grid" id="baseNotes">
                        <%
                        Set rsNotes = ExecuteQuery("SELECT n.* FROM FragranceNotes n INNER JOIN ProductNotes pn ON n.NoteID = pn.NoteID WHERE pn.ProductID = " & CInt(productId) & " AND n.NoteType = '后调' AND n.IsActive <> 0")
                        If Not rsNotes Is Nothing Then
                            Do While Not rsNotes.EOF
                        %>
                        <div class="option-card-wrapper">
                            <label class="option-card">
                                <input type="checkbox" name="baseNote" value="<%= rsNotes("NoteID") %>" 
                                    data-price="<%= CDbl(rsNotes("PriceAddition")) %>" 
                                    data-name="<%= HTMLEncode(rsNotes("NoteName")) %>" 
                                    onclick="toggleNote(this)">
                                <div class="option-content">
                                    <span class="option-name"><%= HTMLEncode(rsNotes("NoteName")) %></span>
                                    <span class="option-desc"><%= HTMLEncode(rsNotes("Description") & "") %></span>
                                    <% If CDbl(rsNotes("PriceAddition")) > 0 Then %>
                                    <span class="option-price">+<%= FormatMoney(rsNotes("PriceAddition")) %>/100%</span>
                                    <% End If %>
                                </div>
                            </label>
                            <div class="percent-input-wrapper" style="display:none; margin-top: 10px;">
                                <label>比例: <input type="number" name="percent_base_<%= rsNotes("NoteID") %>" class="note-percent" min="1" max="100" 
                                    value="<%= rsNotes("RecommendedPercentage") %>" 
                                    onchange="calculateTotal()"> %</label>
                            </div>
                        </div>
                        <%
                                rsNotes.MoveNext
                            Loop
                            rsNotes.Close
                        End If
                        Set rsNotes = Nothing
                        %>
                    </div>
                </div>
                
                <!-- 香调配比实时验证面板 - 仅Custom类型显示 -->
                <% If productType = "custom" Then %>
                <div class="option-section ratio-validation-panel" id="ratioValidationPanel">
                    <h3><i class="fas fa-chart-pie"></i> 配比验证</h3>
                    <div class="ratio-status-list">
                        <div class="ratio-status-item" id="topStatus">
                            <span class="status-label">前调：</span>
                            <span class="status-value" id="topCurrentPercent">0%</span>
                            <span class="status-requirement">≥ 最小 <%= minTopPercent %>%</span>
                            <span class="status-icon" id="topStatusIcon"></span>
                        </div>
                        <div class="ratio-status-item" id="middleStatus">
                            <span class="status-label">中调：</span>
                            <span class="status-value" id="middleCurrentPercent">0%</span>
                            <span class="status-requirement">≥ 最小 <%= minMiddlePercent %>%</span>
                            <span class="status-icon" id="middleStatusIcon"></span>
                        </div>
                        <div class="ratio-status-item" id="baseStatus">
                            <span class="status-label">后调：</span>
                            <span class="status-value" id="baseCurrentPercent">0%</span>
                            <span class="status-requirement">≥ 最小 <%= minBasePercent %>%</span>
                            <span class="status-icon" id="baseStatusIcon"></span>
                        </div>
                        <div class="ratio-status-item total-status" id="totalStatus">
                            <span class="status-label">总计：</span>
                            <span class="status-value" id="totalCurrentPercent">0%</span>
                            <span class="status-requirement">= 100%</span>
                            <span class="status-icon" id="totalStatusIcon"></span>
                        </div>
                    </div>
                    <div class="ratio-validation-message" id="ratioValidationMessage" style="display:none;"></div>
                </div>
                <% End If %>
                <% End If %>
                                
                <!-- 选择容量 -->
                <div class="option-section">
                    <h3><i class="fas fa-tint"></i> 容量规格</h3>
                    <div class="option-grid volume-grid" id="volumes">
                        <%
                        Dim rsVolumes, hasProductVolumes
                        hasProductVolumes = False
                        
                        ' 首先尝试查询产品关联的容量配置（包含价格，供Fixed类型使用）
                        Set rsVolumes = ExecuteQuery("SELECT pvp.VolumeID, pvp.Price, v.VolumeML, v.VolumeName, v.PriceMultiplier, v.IsActive FROM ProductVolumePrices pvp INNER JOIN Volumes v ON pvp.VolumeID = v.VolumeID WHERE pvp.ProductID = " & CLng(productId) & " AND v.IsActive <> 0 ORDER BY v.VolumeML")
                        
                        ' 检查是否有产品关联的容量配置
                        If Not rsVolumes Is Nothing Then
                            If Not rsVolumes.EOF Then
                                hasProductVolumes = True
                            Else
                                rsVolumes.Close
                                Set rsVolumes = Nothing
                            End If
                        End If
                        
                        ' 如果没有产品关联的容量配置（旧产品未配置），回退显示全部启用容量
                        If Not hasProductVolumes Then
                            Set rsVolumes = ExecuteQuery("SELECT * FROM Volumes WHERE IsActive <> 0 ORDER BY VolumeML")
                        End If
                        
                        If Not rsVolumes Is Nothing Then
                            Dim isFirst, volMultiplier
                            isFirst = True
                            Do While Not rsVolumes.EOF
                                ' 计算容量系数
                                Select Case CStr(rsVolumes("VolumeML"))
                                    Case "5"
                                        volMultiplier = 0.3
                                    Case "15"
                                        volMultiplier = 0.5
                                    Case "30"
                                        volMultiplier = 1.0
                                    Case "50"
                                        volMultiplier = 1.5
                                    Case "100"
                                        volMultiplier = 2.5
                                    Case Else
                                        volMultiplier = 1.0
                                End Select
                        %>
                        <label class="option-card volume-card">
                            <input type="radio" name="volume" value="<%= rsVolumes("VolumeID") %>" 
                                <% If productType = "standard" Then %>
                                <% If hasProductVolumes Then %>
                                data-fixed-price="<%= CDbl(rsVolumes("Price")) %>"
                                <% Else %>
                                data-fixed-price="<%= CDbl(rsProduct("BasePrice")) * volMultiplier %>"
                                <% End If %>
                                <% Else %>
                                data-multiplier="<%= CDbl(rsVolumes("PriceMultiplier")) %>" 
                                <% End If %>
                                data-name="<%= HTMLEncode(rsVolumes("VolumeName")) %>" <% If isFirst Then Response.Write "checked" End If %>>
                            <div class="option-content">
                                <span class="volume-size"><%= rsVolumes("VolumeML") %>ml</span>
                                <span class="option-name"><%= HTMLEncode(rsVolumes("VolumeName")) %></span>
                                <% If productType = "standard" Then %>
                                <span class="volume-multiplier">×<%= FormatNumber(volMultiplier, 1) %></span>
                                <% Else %>
                                <span class="volume-multiplier">×<%= CDbl(rsVolumes("PriceMultiplier")) %></span>
                                <% End If %>
                            </div>
                        </label>
                        <%
                                isFirst = False
                                rsVolumes.MoveNext
                            Loop
                            rsVolumes.Close
                        End If
                        Set rsVolumes = Nothing
                        %>
                    </div>
                </div>

                <% If productType <> "standard" Then %>
                <!-- 选择瓶子 -->
                <div class="option-section">
                    <h3><i class="fas fa-wine-bottle"></i> 瓶身款式</h3>
                    <div class="option-grid bottle-grid" id="bottles">
                        <%
                        Dim rsBottles, bottleSql, hasProductBottles
                        hasProductBottles = False
                        
                        ' 先尝试查询产品关联的瓶型
                        bottleSql = "SELECT bs.BottleID, bs.BottleName, bs.Description, bs.ImageURL, " & _
                                    "bs.PriceAddition, pbs.CustomPrice " & _
                                    "FROM ProductBottleStyles pbs " & _
                                    "INNER JOIN BottleStyles bs ON pbs.BottleID = bs.BottleID " & _
                                    "WHERE pbs.ProductID = " & CInt(productId) & " AND bs.IsActive <> 0"
                        Set rsBottles = ExecuteQuery(bottleSql)
                        
                        ' 检查是否有产品关联的瓶型
                        If Not rsBottles Is Nothing Then
                            If Not rsBottles.EOF Then
                                hasProductBottles = True
                            Else
                                rsBottles.Close
                                Set rsBottles = Nothing
                            End If
                        End If
                        
                        ' 如果没有产品关联的瓶型，回退显示全部启用瓶型
                        If Not hasProductBottles Then
                            Set rsBottles = ExecuteQuery("SELECT BottleID, BottleName, Description, ImageURL, PriceAddition, NULL AS CustomPrice FROM BottleStyles WHERE IsActive <> 0")
                        End If
                        
                        If Not rsBottles Is Nothing Then
                            isFirst = True
                            Do While Not rsBottles.EOF
                                Dim bottlePrice, displayPrice
                                ' 使用 CustomPrice 如果存在，否则回退到 PriceAddition
                                If IsNull(rsBottles("CustomPrice")) Then
                                    bottlePrice = CDbl(rsBottles("PriceAddition"))
                                Else
                                    bottlePrice = CDbl(rsBottles("CustomPrice"))
                                End If
                                displayPrice = bottlePrice
                        %>
                        <label class="option-card bottle-card">
                            <input type="radio" name="bottle" value="<%= rsBottles("BottleID") %>" data-price="<%= displayPrice %>" data-name="<%= HTMLEncode(rsBottles("BottleName")) %>" <% If isFirst Then Response.Write "checked" End If %>>
                            <div class="option-content">
                                <span class="option-name"><%= HTMLEncode(rsBottles("BottleName")) %></span>
                                <span class="option-desc"><%= HTMLEncode(rsBottles("Description")) %></span>
                                <% If displayPrice > 0 Then %>
                                <span class="option-price">+<%= FormatMoney(displayPrice) %></span>
                                <% End If %>
                            </div>
                        </label>
                        <%
                                isFirst = False
                                rsBottles.MoveNext
                            Loop
                            rsBottles.Close
                        End If
                        Set rsBottles = Nothing
                        %>
                    </div>
                </div>
                <% End If %>

                <% If productEngravable Then %>
                <!-- 个性化标签 -->
                <div class="option-section">
                    <h3><i class="fas fa-pen-fancy"></i> 专属标签 <span class="option-tip">可选</span></h3>
                    <div class="checkbox-group" style="margin-bottom: 10px;">
                        <label style="display: flex; align-items: center; cursor: pointer;">
                            <input type="checkbox" id="engravingCheckbox" name="engravingEnabled" value="1" style="margin-right: 8px;" onchange="toggleEngravingInput()">
                            <span>添加瓶身刻字</span>
                        </label>
                    </div>
                    <div id="engravingInputWrapper" style="display: none;">
                        <div class="label-input">
                            <input type="text" name="customLabel" id="customLabel" maxlength="20" placeholder="输入您想刻印的文字（最多20字）">
                            <span class="char-count"><span id="charCount">0</span>/20</span>
                        </div>
                        <% If productEngravingPrice > 0 Then %>
                        <small style="color:#888;">刻字附加费用: <%= FormatMoney(productEngravingPrice) %></small>
                        <% End If %>
                    </div>
                    <% If productEngravingPrice > 0 Then %>
                    <input type="hidden" id="engravingPrice" value="<%= productEngravingPrice %>">
                    <% End If %>
                </div>
                <% End If %>

                <!-- 数量选择 -->
                <div class="quantity-section">
                    <label>数量:</label>
                    <div class="quantity-input">
                        <button type="button" class="qty-btn minus" onclick="changeQty(-1)">-</button>
                        <input type="number" name="quantity" id="quantity" value="1" min="1" max="99" readonly>
                        <button type="button" class="qty-btn plus" onclick="changeQty(1)">+</button>
                    </div>
                </div>

                <!-- 价格汇总 -->
                <div class="price-summary">
                    <% If productType = "standard" Then %>
                    <!-- 品牌定香价格摘要 -->
                    <div class="summary-row">
                        <span>选择规格:</span>
                        <span id="summaryBase"><%= FormatMoney(rsProduct("BasePrice")) %></span>
                    </div>
                    <% If CBool(productEngravable & "") And productEngravingPrice > 0 Then %>
                    <div class="summary-row" id="engravingSummaryRow" style="display:none;">
                        <span>刻字费用:</span>
                        <span id="summaryEngraving"><%= FormatMoney(productEngravingPrice) %></span>
                    </div>
                    <% End If %>
                    <% ElseIf productType = "kol" Then %>
                    <!-- KOL商品价格摘要 -->
                    <div class="summary-row">
                        <span>基础价格:</span>
                        <span id="summaryBase"><%= FormatMoney(rsProduct("BasePrice")) %></span>
                    </div>
                    <div class="summary-row">
                        <span>瓶身附加:</span>
                        <span id="summaryBottle">¥0.00</span>
                    </div>
                    <div class="summary-row">
                        <span>容量系数:</span>
                        <span id="summaryMultiplier">×1.0</span>
                    </div>
                    <% If CBool(productEngravable) And productEngravingPrice > 0 Then %>
                    <div class="summary-row" id="engravingSummaryRow" style="display:none;">
                        <span>刻字费用:</span>
                        <span id="summaryEngraving"><%= FormatMoney(productEngravingPrice) %></span>
                    </div>
                    <% End If %>
                    <% Else %>
                    <!-- 定制香水价格摘要 -->
                    <div class="summary-row">
                        <span>基础价格:</span>
                        <span id="summaryBase"><%= FormatMoney(rsProduct("BasePrice")) %></span>
                    </div>
                    <div class="summary-row">
                        <span>香调附加:</span>
                        <span id="summaryNotes">¥0.00</span>
                    </div>
                    <div class="summary-row">
                        <span>瓶身附加:</span>
                        <span id="summaryBottle">¥0.00</span>
                    </div>
                    <div class="summary-row">
                        <span>容量系数:</span>
                        <span id="summaryMultiplier">×1.0</span>
                    </div>
                    <div class="summary-row">
                        <span>配比总和:</span>
                        <span id="totalPercentDisplay">0%</span>
                    </div>
                    <% If CBool(productEngravable) And productEngravingPrice > 0 Then %>
                    <div class="summary-row" id="engravingSummaryRow" style="display:none;">
                        <span>刻字费用:</span>
                        <span id="summaryEngraving"><%= FormatMoney(productEngravingPrice) %></span>
                    </div>
                    <% End If %>
                    <% End If %>
                    <div class="summary-total">
                        <span>合计:</span>
                        <span class="total-price" id="totalPrice"><%= FormatMoney(rsProduct("BasePrice")) %></span>
                    </div>
                </div>

                <!-- 操作按钮 -->
                <div class="action-buttons">
                    <button type="submit" class="btn btn-primary btn-lg btn-block" id="btnAddToCart">
                        <i class="fas fa-shopping-cart"></i> 加入购物车
                    </button>
                    <button type="button" class="btn btn-secondary btn-lg btn-block" id="btnBuyNow">
                        <i class="fas fa-bolt"></i> 立即购买
                    </button>
                    <button type="button" class="btn btn-outline btn-lg btn-block" id="favoriteBtn" onclick="toggleFavorite()">
                        <i class="fas fa-heart"></i> <span id="favoriteText">收藏</span>
                    </button>
                </div>
            </form>
        </div>
    </div>

    <!-- 产品详情标签页 -->
    <div class="product-tabs">
        <div class="tabs-header">
            <button class="tab-btn active" data-tab="details">产品详情</button>
            <button class="tab-btn" data-tab="notes">香调说明</button>
            <button class="tab-btn" data-tab="shipping">配送说明</button>
        </div>
        <div class="tabs-content">
            <div class="tab-pane active" id="details">
                <h3>产品介绍</h3>
                <p><%= HTMLEncode(rsProduct("Description")) %></p>
                <h4>使用建议</h4>
                <ul>
                    <li>喷洒于手腕、颈部、耳后等脉搏点</li>
                    <li>距离皮肤15-20厘米喷洒</li>
                    <li>避免阳光直射，存放于阴凉处</li>
                    <li>开封后建议12个月内使用完毕</li>
                </ul>
            </div>
            <div class="tab-pane" id="notes">
                <h3>香调层次</h3>
                <div class="notes-pyramid">
                    <div class="pyramid-level top">
                        <h4><i class="fas fa-wind"></i> 前调</h4>
                        <p>香水的第一印象，通常持续15-30分钟。包含较轻的香调，如柑橘、薰衣草等。</p>
                    </div>
                    <div class="pyramid-level middle">
                        <h4><i class="fas fa-heart"></i> 中调</h4>
                        <p>香水的核心，持续2-4小时。花香、果香等构成香水的主体。</p>
                    </div>
                    <div class="pyramid-level base">
                        <h4><i class="fas fa-moon"></i> 后调</h4>
                        <p>最持久的部分，可持续数小时。木质香、麝香等提供深度和持久力。</p>
                    </div>
                </div>
            </div>
            <div class="tab-pane" id="shipping">
                <h3>配送说明</h3>
                <ul>
                    <li>订单满299元包邮，未满299元收取15元运费</li>
                    <li>定制产品需要3-5个工作日制作</li>
                    <li>发货后一般2-5天送达</li>
                    <li>提供顺丰快递配送，支持全国配送</li>
                </ul>
                <h3>退换政策</h3>
                <p>由于香水为个人定制产品，非质量问题不支持退换。如收到产品有质量问题，请在签收后7天内联系客服。</p>
            </div>
        </div>
    </div>
</div>

<% 
' 提前提取分享参数（必须在 rsProduct.Close 之前）
Dim shareUrl, shareTitle, shareDesc, shareImage
shareUrl = SITE_URL & "/product.asp?id=" & productId
shareTitle = rsProduct("ProductName")
If IsNull(shareTitle) Then shareTitle = ""
shareDesc = rsProduct("Description")
If IsNull(shareDesc) Then shareDesc = "" Else shareDesc = CStr(shareDesc)
shareImage = rsProduct("ImageURL")
If IsNull(shareImage) Then shareImage = ""
%>

<%
rsProduct.Close
Set rsProduct = Nothing
%>

<script>
// 全局变量
var basePrice = 0;
var productType = '<%= productType %>';

// 香调最小比例配置（从后端获取）
var minTopPercent = <%= minTopPercent %>;
var minMiddlePercent = <%= minMiddlePercent %>;
var minBasePercent = <%= minBasePercent %>;

$(document).ready(function() {
    basePrice = parseFloat($('#basePrice').data('price')) || 0;
    
    // 初始计算
    calculateTotal();
    
    // 字符计数
    $('#customLabel').on('input', function() {
        var len = $(this).val().length;
        $('#charCount').text(len);
        // 刻字内容变化时重新计算价格
        calculateTotal();
    });
    
    // 刻字复选框监听
    $('#engravingCheckbox').on('change', function() {
        calculateTotal();
    });
    
    // 标签页切换
    $('.tab-btn').click(function() {
        var tab = $(this).data('tab');
        $('.tab-btn').removeClass('active');
        $(this).addClass('active');
        $('.tab-pane').removeClass('active');
        $('#' + tab).addClass('active');
    });
    
    // 监听选项变化
    $('input[type="radio"]').change(function() {
        calculateTotal();
    });

    // 加入购物车按钮点击
    $('#btnAddToCart').click(function(e) {
        e.preventDefault();
        submitCustomCart();
    });

    // 立即购买按钮点击
    $('#btnBuyNow').click(function(e) {
        e.preventDefault();
        submitCustomBuyNow();
    });
});

// 切换香调选中状态
window.toggleNote = function(checkbox) {
    if (productType === 'kol') return; // KOL产品不可修改

    var $wrapper = $(checkbox).closest('.option-card-wrapper');
    var $percentWrapper = $wrapper.find('.percent-input-wrapper');
    var groupName = $(checkbox).attr('name');
    
    if (checkbox.checked) {
        // 限制每类最多选3种
        if ($('input[name="' + groupName + '"]:checked').length > 3) {
            checkbox.checked = false;
            alert('每类香调最多选择3种基香');
            return;
        }
        $percentWrapper.show();
    } else {
        $percentWrapper.hide();
    }
    calculateTotal();
};

// 切换刻字输入框显示/隐藏
function toggleEngravingInput() {
    var checkbox = document.getElementById('engravingCheckbox');
    var wrapper = document.getElementById('engravingInputWrapper');
    var summaryRow = document.getElementById('engravingSummaryRow');
    var engravingInput = document.getElementById('customLabel');
    if (checkbox && wrapper) {
        wrapper.style.display = checkbox.checked ? 'block' : 'none';
    }
    // 取消选中时清除刻字文本，防止残留文本被提交
    if (!checkbox.checked && engravingInput) {
        engravingInput.value = '';
    }
    // 显示/隐藏刻字费用明细行
    if (summaryRow) {
        var engravingText = engravingInput ? engravingInput.value : '';
        summaryRow.style.display = (checkbox.checked && engravingText.length > 0) ? 'flex' : 'none';
    }
    $('#charCount').text('0');
    calculateTotal();
}

// 计算总价
function calculateTotal() {
    var notesTotal = 0;
    var totalPercent = 0;
    var topPercent = 0, middlePercent = 0, basePercent = 0;  // 每种调性比例
    var bottlePrice = 0, multiplier = 1, fixedPrice = 0;
    var engravingPrice = 0;
    
    if (productType === 'standard') {
        var $selectedVol = $('input[name="volume"]:checked');
        var selectedFixedPrice = parseFloat($selectedVol.data('fixed-price'));
        if (!isNaN(selectedFixedPrice) && selectedFixedPrice > 0) {
            fixedPrice = selectedFixedPrice;
        } else {
            fixedPrice = basePrice;
        }
        
        // 计算刻字费用
        var engravingChecked = $('#engravingCheckbox').is(':checked');
        var engravingText = $('#customLabel').val() || '';
        if (engravingChecked && engravingText.length > 0) {
            engravingPrice = parseFloat($('#engravingPrice').val()) || 0;
        }
        
        // 更新刻字费用明细行显示状态
        var engravingSummaryRow = document.getElementById('engravingSummaryRow');
        if (engravingSummaryRow) {
            engravingSummaryRow.style.display = (engravingChecked && engravingText.length > 0) ? 'flex' : 'none';
        }
        
        var quantity = parseInt($('#quantity').val()) || 1;
        var total = (fixedPrice + engravingPrice) * quantity;
        
        $('#summaryBase').text('¥' + fixedPrice.toFixed(2));
        $('#totalPrice').text('¥' + total.toFixed(2));
        return { total: 100, top: 0, middle: 0, base: 0 };
    }
    
    // KOL商品 - 配比已预设，不需要验证
    if (productType === 'kol') {
        if ($('input[name="bottle"]:checked').length) {
            bottlePrice = parseFloat($('input[name="bottle"]:checked').data('price')) || 0;
        }
        if ($('input[name="volume"]:checked').length) {
            multiplier = parseFloat($('input[name="volume"]:checked').data('multiplier')) || 1;
        }
        
        // 计算刻字费用
        var engravingChecked = $('#engravingCheckbox').is(':checked');
        var engravingText = $('#customLabel').val() || '';
        if (engravingChecked && engravingText.length > 0) {
            engravingPrice = parseFloat($('#engravingPrice').val()) || 0;
        }
        
        // 更新刻字费用明细行显示状态
        var engravingSummaryRow = document.getElementById('engravingSummaryRow');
        if (engravingSummaryRow) {
            engravingSummaryRow.style.display = (engravingChecked && engravingText.length > 0) ? 'flex' : 'none';
        }
        
        var subtotal = (basePrice + bottlePrice) * multiplier + engravingPrice;
        var quantity = parseInt($('#quantity').val()) || 1;
        var total = subtotal * quantity;
        
        $('#summaryBottle').text('¥' + bottlePrice.toFixed(2));
        $('#summaryMultiplier').text('×' + multiplier.toFixed(1));
        $('#summaryBase').text('¥' + basePrice.toFixed(2));
        $('#totalPrice').text('¥' + total.toFixed(2));
        return { total: 100, top: 0, middle: 0, base: 0 };
    }

    // 遍历前调（仅定制香水）
    $('input[name="topNote"]:checked').each(function() {
        var $checkbox = $(this);
        var notePrice = parseFloat($checkbox.data('price')) || 0;
        var $wrapper = $checkbox.closest('.option-card-wrapper');
        var $percentInput = $wrapper.find('.note-percent');
        var percent = parseFloat($percentInput.val()) || 0;
        topPercent += percent;
        totalPercent += percent;
        notesTotal += (notePrice * percent / 100);
    });
    
    // 遍历中调
    $('input[name="middleNote"]:checked').each(function() {
        var $checkbox = $(this);
        var notePrice = parseFloat($checkbox.data('price')) || 0;
        var $wrapper = $checkbox.closest('.option-card-wrapper');
        var $percentInput = $wrapper.find('.note-percent');
        var percent = parseFloat($percentInput.val()) || 0;
        middlePercent += percent;
        totalPercent += percent;
        notesTotal += (notePrice * percent / 100);
    });
    
    // 遍历后调
    $('input[name="baseNote"]:checked').each(function() {
        var $checkbox = $(this);
        var notePrice = parseFloat($checkbox.data('price')) || 0;
        var $wrapper = $checkbox.closest('.option-card-wrapper');
        var $percentInput = $wrapper.find('.note-percent');
        var percent = parseFloat($percentInput.val()) || 0;
        basePercent += percent;
        totalPercent += percent;
        notesTotal += (notePrice * percent / 100);
    });
    
    if ($('input[name="bottle"]:checked').length) {
        bottlePrice = parseFloat($('input[name="bottle"]:checked').data('price')) || 0;
    }
    if ($('input[name="volume"]:checked').length) {
        multiplier = parseFloat($('input[name="volume"]:checked').data('multiplier')) || 1;
    }
    
    // 计算刻字费用
    var engravingChecked = $('#engravingCheckbox').is(':checked');
    var engravingText = $('#customLabel').val() || '';
    if (engravingChecked && engravingText.length > 0) {
        engravingPrice = parseFloat($('#engravingPrice').val()) || 0;
    }
    
    // 更新刻字费用明细行显示状态
    var engravingSummaryRow = document.getElementById('engravingSummaryRow');
    if (engravingSummaryRow) {
        engravingSummaryRow.style.display = (engravingChecked && engravingText.length > 0) ? 'flex' : 'none';
    }
    
    var subtotal = (basePrice + notesTotal + bottlePrice) * multiplier + engravingPrice;
    var quantity = parseInt($('#quantity').val()) || 1;
    var total = subtotal * quantity;
    
    $('#summaryNotes').text('¥' + notesTotal.toFixed(2));
    $('#summaryBottle').text('¥' + bottlePrice.toFixed(2));
    $('#summaryMultiplier').text('×' + multiplier.toFixed(1));
    $('#summaryBase').text('¥' + basePrice.toFixed(2));
    $('#totalPercentDisplay').text(totalPercent + '%');
    
    if (Math.abs(totalPercent - 100) > 0.01) {
        $('#totalPercentDisplay').css('color', 'red');
    } else {
        $('#totalPercentDisplay').css('color', 'green');
    }
    
    $('#totalPrice').text('¥' + total.toFixed(2));
    
    // 更新实时验证面板显示
    updateRatioValidationPanel(totalPercent, topPercent, middlePercent, basePercent);
    
    return { total: totalPercent, top: topPercent, middle: middlePercent, base: basePercent };
}

// 更新香调配比实时验证面板
function updateRatioValidationPanel(totalPercent, topPercent, middlePercent, basePercent) {
    // 仅对Custom类型显示验证
    if (productType !== 'custom') {
        $('#ratioValidationPanel').hide();
        return;
    }
    
    $('#ratioValidationPanel').show();
    
    // 更新各调性当前百分比显示
    $('#topCurrentPercent').text(topPercent + '%');
    $('#middleCurrentPercent').text(middlePercent + '%');
    $('#baseCurrentPercent').text(basePercent + '%');
    $('#totalCurrentPercent').text(totalPercent + '%');
    
    // 验证前调
    var topValid = topPercent >= minTopPercent;
    updateStatusItem('topStatus', 'topStatusIcon', topValid, topPercent + '% ≥ ' + minTopPercent + '%');
    
    // 验证中调
    var middleValid = middlePercent >= minMiddlePercent;
    updateStatusItem('middleStatus', 'middleStatusIcon', middleValid, middlePercent + '% ≥ ' + minMiddlePercent + '%');
    
    // 验证后调
    var baseValid = basePercent >= minBasePercent;
    updateStatusItem('baseStatus', 'baseStatusIcon', baseValid, basePercent + '% ≥ ' + minBasePercent + '%');
    
    // 验证总计（使用容差0.01避免浮点数精度问题）
    var totalValid = Math.abs(totalPercent - 100) < 0.01;
    updateStatusItem('totalStatus', 'totalStatusIcon', totalValid, totalPercent.toFixed(1) + '% / 100%');
    
    // 显示验证消息
    var message = '';
    var allValid = topValid && middleValid && baseValid && totalValid;
    
    if (!allValid) {
        var issues = [];
        if (!topValid) issues.push('前调比例不足（需≥' + minTopPercent + '%）');
        if (!middleValid) issues.push('中调比例不足（需≥' + minMiddlePercent + '%）');
        if (!baseValid) issues.push('后调比例不足（需≥' + minBasePercent + '%）');
        if (!totalValid) {
            var diff = Math.abs(totalPercent - 100);
            if (totalPercent < 100) {
                issues.push('总配比不足100%（还差' + diff.toFixed(1) + '%）');
            } else {
                issues.push('总配比超过100%（超出' + diff.toFixed(1) + '%）');
            }
        }
        message = '⚠️ ' + issues.join('；');
        $('#ratioValidationMessage').removeClass('success').addClass('error').html(message).show();
    } else {
        message = '✓ 配比设置正确，可以提交订单';
        $('#ratioValidationMessage').removeClass('error').addClass('success').html(message).show();
    }
}

// 更新单个状态项的显示
function updateStatusItem(itemId, iconId, isValid, tooltipText) {
    var $item = $('#' + itemId);
    var $icon = $('#' + iconId);
    
    if (isValid) {
        $item.removeClass('invalid').addClass('valid');
        $icon.html('✓').attr('title', tooltipText);
    } else {
        $item.removeClass('valid').addClass('invalid');
        $icon.html('✗').attr('title', tooltipText);
    }
}

function changeQty(delta) {
    var qty = parseInt($('#quantity').val()) || 1;
    qty += delta;
    if (qty < 1) qty = 1;
    if (qty > 99) qty = 99;
    $('#quantity').val(qty);
    calculateTotal();
}

function submitCustomCart() {
    var result = calculateTotal();
    // 只有定制香水需要验证配比，KOL商品配比已在后台验证
    if (productType === 'custom') {
        // 验证总和为100%（使用容差0.01避免浮点数精度问题）
        if (Math.abs(result.total - 100) > 0.01) {
            alert('前、中、后调所选基香的百分比总和必须等于100%（当前：' + result.total.toFixed(1) + '%）');
            return;
        }
        // 验证每种调性最小比例
        if (result.top < minTopPercent - 0.01) {
            alert('前调比例不能低于' + minTopPercent + '%，当前为' + result.top.toFixed(1) + '%');
            return;
        }
        if (result.middle < minMiddlePercent - 0.01) {
            alert('中调比例不能低于' + minMiddlePercent + '%，当前为' + result.middle.toFixed(1) + '%');
            return;
        }
        if (result.base < minBasePercent - 0.01) {
            alert('后调比例不能低于' + minBasePercent + '%，当前为' + result.base.toFixed(1) + '%');
            return;
        }
    }

    var productId = $('#customizeForm input[name="productId"]').val();
    if (!productId) {
        alert('错误：未找到产品ID');
        return;
    }
    
    var $form = $('#customizeForm');
    $form.attr('action', '/api/cart_add.asp');
    $form.attr('method', 'post');
    $form[0].submit();
}

function submitCustomBuyNow() {
    var result = calculateTotal();
    // 只有定制香水需要验证配比，KOL商品配比已在后台验证
    if (productType === 'custom') {
        // 验证总和为100%（使用容差0.01避免浮点数精度问题）
        if (Math.abs(result.total - 100) > 0.01) {
            alert('前、中、后调所选基香的百分比总和必须等于100%（当前：' + result.total.toFixed(1) + '%）');
            return;
        }
        // 验证每种调性最小比例
        if (result.top < minTopPercent - 0.01) {
            alert('前调比例不能低于' + minTopPercent + '%，当前为' + result.top.toFixed(1) + '%');
            return;
        }
        if (result.middle < minMiddlePercent - 0.01) {
            alert('中调比例不能低于' + minMiddlePercent + '%，当前为' + result.middle.toFixed(1) + '%');
            return;
        }
        if (result.base < minBasePercent - 0.01) {
            alert('后调比例不能低于' + minBasePercent + '%，当前为' + result.base.toFixed(1) + '%');
            return;
        }
    }

    var productId = $('#customizeForm input[name="productId"]').val();
    if (!productId) {
        alert('错误：未找到产品ID');
        return;
    }
    
    var $form = $('#customizeForm');
    if ($('#buyNowAction').length === 0) {
        $form.append('<input type="hidden" id="buyNowAction" name="buyNow" value="1">');
    }
    $form.attr('action', '/api/cart_add.asp');
    $form.attr('method', 'post');
    $form[0].submit();
}

function toggleFavorite() {
    var productId = $('input[name="productId"]').val();
    
    if (!productId || productId === '') {
        alert('错误：未找到产品ID');
        return;
    }
    
    // 检查用户是否已登录
    var isLoggedIn = $('#userLoginStatus').val() === 'true';
    if (!isLoggedIn) {
        if (confirm('您需要先登录才能收藏商品，是否前往登录？')) {
            window.location.href = '/user/login.asp?return=' + encodeURIComponent(window.location.href);
        }
        return;
    }
    
    // 获取当前收藏状态
    var isFavorite = $('#favoriteBtn').hasClass('active');
    var action = isFavorite ? 'remove' : 'add';
    
    // 输出调试信息
    console.log('准备发送收藏请求:');
    console.log('产品ID:', productId);
    console.log('操作:', action);
    console.log('API地址:', '/api/favorites.asp');
    
    // 发送AJAX请求
    $.ajax({
        url: '/api/favorites.asp',
        type: 'POST',
        data: {
            action: action,
            productId: productId,
            csrf_token: csrfToken
        },
        dataType: 'json',
        success: function(response) {
            if (response.success) {
                if (action === 'add') {
                    $('#favoriteBtn').addClass('active');
                    $('#favoriteBtn').html('<i class="fas fa-heart"></i> <span>已收藏</span>');
                    alert('收藏成功！');
                } else {
                    $('#favoriteBtn').removeClass('active');
                    $('#favoriteBtn').html('<i class="fas fa-heart"></i> <span>收藏</span>');
                    alert('已取消收藏！');
                }
            } else {
                alert(response.message || '操作失败');
            }
        },
        error: function(xhr, status, error) {
            console.error('收藏请求失败:');
            console.error('状态:', status);
            console.error('错误:', error);
            console.error('响应状态码:', xhr.status);
            console.error('响应文本:', xhr.responseText);
            
            var errorMsg = '网络错误: ' + status;
            if (xhr.status === 404) {
                errorMsg = '错误: API接口未找到 (404)';
            } else if (xhr.status === 500) {
                errorMsg = '错误: 服务器内部错误 (500)';
            } else {
                // 尝试解析响应，看是否是JSON格式
                try {
                    var responseJson = JSON.parse(xhr.responseText);
                    if (responseJson.message) {
                        errorMsg = '错误: ' + responseJson.message;
                    } else {
                        errorMsg = '错误详情: ' + xhr.responseText.substring(0, 200);
                    }
                } catch (e) {
                    // 响应不是JSON格式，直接显示原始内容
                    errorMsg = '非JSON响应: ' + xhr.responseText.substring(0, 200);
                }
            }
            alert(errorMsg + '\n\n请打开浏览器控制台查看详细错误信息');
        }
    });
}

// 页面加载时检查收藏状态
$(document).ready(function() {
    var productId = $('input[name="productId"]').val();
    var isLoggedIn = $('#userLoginStatus').val() === 'true';
    if (productId && isLoggedIn) {
        // 检查当前商品是否已被收藏
        $.ajax({
            url: '/api/favorites.asp',
            type: 'GET',
            data: {
                action: 'check',
                productId: productId
            },
            dataType: 'json',
            success: function(response) {
                if (response.success && response.isFavorite) {
                    $('#favoriteBtn').addClass('active');
                    $('#favoriteBtn').html('<i class="fas fa-heart"></i> <span>已收藏</span>');
                }
            }
        });
    }
});
</script>

<!-- 社交分享 -->
<% Call SU_RenderShareSection(shareUrl, shareTitle, shareDesc, shareImage) %>

<!-- 推荐区域：猜你喜欢 -->
<%
' 获取同类产品推荐
Dim rsRelated
Set rsRelated = RE_GetRelatedProducts(productId, productType, 6)
%>
<div class="recommendation-section">
    <div class="container">
        <h2 class="section-title"><i class="fas fa-magic"></i> 猜你喜欢</h2>
        <p class="section-desc">相似香氛推荐</p>
        <% Call RE_RenderRecommendations(rsRelated, "related-rec", True) %>
        <%
        If Not rsRelated Is Nothing Then
            rsRelated.Close
            Set rsRelated = Nothing
        End If
        %>
    </div>
</div>

<style>
.recommendation-section {
    padding: 40px 0;
    background: #f8f9fa;
    margin-top: 30px;
}
.recommendation-section .section-title {
    font-size: 22px;
    color: #333;
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.recommendation-section .section-title i { color: #ff6f61; }
.recommendation-section .section-desc {
    color: #888;
    font-size: 14px;
    margin-bottom: 25px;
}
.recommendation-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
    gap: 20px;
}
.rec-card {
    background: #fff;
    border-radius: 12px;
    overflow: hidden;
    text-decoration: none;
    border: 1px solid #eee;
    transition: all 0.3s ease;
    display: block;
}
.rec-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 10px 30px rgba(0,0,0,0.1);
}
.rec-img-wrapper {
    position: relative;
    height: 180px;
    background: #f5f5f5;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
}
.rec-img-wrapper img {
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
    transition: transform 0.3s;
}
.rec-card:hover .rec-img-wrapper img { transform: scale(1.05); }
.rec-badge {
    position: absolute;
    top: 10px;
    left: 10px;
    background: #ff6f61;
    color: #fff;
    padding: 3px 8px;
    border-radius: 4px;
    font-size: 11px;
    font-weight: 600;
}
.rec-info {
    padding: 12px;
}
.rec-info h4 {
    font-size: 14px;
    color: #333;
    margin: 0 0 6px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.rec-price {
    font-size: 16px;
    font-weight: 700;
    color: #e74c3c;
}
</style>

<!--#include file="includes/footer.asp"-->
<%
Call CloseConnection()
%>
