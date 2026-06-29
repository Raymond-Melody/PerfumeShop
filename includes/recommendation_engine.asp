<%
' ============================================
' 推荐引擎 - Recommendation Engine
' 基于用户行为数据的产品推荐
' ============================================

' ============================================
' 获取用户个性化推荐（协同过滤）
' 基于：购买历史、收藏夹、同类型商品
' ============================================
Function RE_GetUserRecommendations(userId, topN)
    Dim rs, sql, dictResult, count, pid
    Set dictResult = Server.CreateObject("Scripting.Dictionary")
    count = 0
    
    If topN <= 0 Then topN = 6
    
    ' 策略1: 基于用户购买历史推荐同类型产品
    ' Use GROUP BY with MAX to allow ORDER BY NEWID()
    sql = "SELECT TOP " & topN & " p2.ProductID, MAX(p2.ProductName) AS ProductName, MAX(p2.BasePrice) AS BasePrice, " & _
          "MAX(p2.ImageURL) AS ImageURL, MAX(p2.ProductType) AS ProductType " & _
          "FROM OrderDetails od1 " & _
          "INNER JOIN Products p1 ON od1.ProductID = p1.ProductID " & _
          "INNER JOIN Products p2 ON p1.ProductType = p2.ProductType AND p2.IsActive = 1 " & _
          "INNER JOIN Orders o ON od1.OrderID = o.OrderID " & _
          "WHERE o.UserID = " & userId & " AND p2.ProductID NOT IN " & _
          "(SELECT ProductID FROM OrderDetails od2 INNER JOIN Orders o2 ON od2.OrderID = o2.OrderID WHERE o2.UserID = " & userId & ") " & _
          "AND p2.ProductID NOT IN (SELECT ProductID FROM UserFavorites WHERE UserID = " & userId & ") " & _
          "GROUP BY p2.ProductID ORDER BY NEWID()"
    
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number <> 0 Then Err.Clear : Set rs = Nothing
    On Error GoTo 0
    If Not rs Is Nothing Then
        Do While Not rs.EOF And count < topN
            pid = rs("ProductID")
            If Not dictResult.Exists(CStr(pid)) Then
                dictResult.Add CStr(pid), "{""id"":" & pid & ",""name"":""" & _
                    Replace(rs("ProductName") & "", """", "\""") & """" & _
                    ",""price"":" & rs("BasePrice") & ",""img"":""" & _
                    Replace(rs("ImageURL") & "", """", "\""") & """,""type"":""" & _
                    Replace(rs("ProductType") & "", """", "\""") & """}"
                count = count + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' 策略2: 补充收藏夹同类产品
    If count < topN Then
        sql = "SELECT DISTINCT TOP " & (topN - count) & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType " & _
              "FROM UserFavorites uf INNER JOIN Products p ON uf.ProductID = p.ProductID " & _
              "WHERE uf.UserID = " & userId & " AND p.IsActive = 1"
        On Error Resume Next
        Set rs = conn.Execute(sql)
        If Err.Number <> 0 Then Err.Clear : Set rs = Nothing
        On Error GoTo 0
        If Not rs Is Nothing Then
            Do While Not rs.EOF And count < topN
                pid = rs("ProductID")
                If Not dictResult.Exists(CStr(pid)) Then
                    dictResult.Add CStr(pid), "{""id"":" & pid & ",""name"":""" & _
                        Replace(rs("ProductName") & "", """", "\""") & """" & _
                        ",""price"":" & rs("BasePrice") & ",""img"":""" & _
                        Replace(rs("ImageURL") & "", """", "\""") & """,""type"":""" & _
                        Replace(rs("ProductType") & "", """", "\""") & """}"
                    count = count + 1
                End If
                rs.MoveNext
            Loop
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    ' 策略3: 热门产品
    If count < topN Then
        sql = "SELECT TOP " & (topN - count) & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType, " & _
              "ISNULL((SELECT COUNT(*) FROM OrderDetails od INNER JOIN Orders o ON od.OrderID = o.OrderID " & _
              "WHERE od.ProductID = p.ProductID AND o.Status='Paid'), 0) AS SaleCount " & _
              "FROM Products p WHERE p.IsActive = 1 AND p.ProductID NOT IN " & _
              "(SELECT od.ProductID FROM OrderDetails od INNER JOIN Orders o ON od.OrderID = o.OrderID WHERE o.UserID = " & userId & ") " & _
              "ORDER BY SaleCount DESC"
        On Error Resume Next
        Set rs = conn.Execute(sql)
        If Err.Number <> 0 Then Err.Clear : Set rs = Nothing
        On Error GoTo 0
        If Not rs Is Nothing Then
            Do While Not rs.EOF And count < topN
                pid = rs("ProductID")
                If Not dictResult.Exists(CStr(pid)) Then
                    dictResult.Add CStr(pid), "{""id"":" & pid & ",""name"":""" & _
                        Replace(rs("ProductName") & "", """", "\""") & """" & _
                        ",""price"":" & rs("BasePrice") & ",""img"":""" & _
                        Replace(rs("ImageURL") & "", """", "\""") & """,""type"":""" & _
                        Replace(rs("ProductType") & "", """", "\""") & """}"
                    count = count + 1
                End If
                rs.MoveNext
            Loop
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    Set RE_GetUserRecommendations = dictResult
End Function

' ============================================
' 获取关联产品推荐（基于相同产品类型/分类）
' ============================================
Function RE_GetRelatedProducts(productId, productType, topN)
    Dim rs, sql, result, count
    count = 0
    
    If topN <= 0 Then topN = 4
    
    On Error Resume Next
    
    ' 同类型产品推荐（排除当前产品）
    If productType <> "" Then
        sql = "SELECT TOP " & topN & " ProductID, ProductName, BasePrice, ImageURL, ProductType " & _
              "FROM Products WHERE IsActive = 1 AND ProductID <> " & productId & " " & _
              "AND ProductType = '" & SafeSQL(productType) & "' ORDER BY NEWID()"
    Else
        ' 无类型时随机推荐
        sql = "SELECT TOP " & topN & " ProductID, ProductName, BasePrice, ImageURL, ProductType " & _
              "FROM Products WHERE IsActive = 1 AND ProductID <> " & productId & " ORDER BY NEWID()"
    End If
    
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If count < topN Then
                If count = 0 Then
                    Set RE_GetRelatedProducts = rs
                Else
                    ' VBScript无法直接处理多个Recordset，返回第一个匹配
                End If
                count = count + 1
            End If
            rs.MoveNext
        Loop
        If count > 0 Then
            rs.MoveFirst
            Set RE_GetRelatedProducts = rs
        Else
            Set RE_GetRelatedProducts = Nothing
        End If
    Else
        Set RE_GetRelatedProducts = Nothing
    End If
    ' 注意：调用方需要自行关闭Recordset
End Function

' ============================================
' 获取热销/高分产品
' ============================================
Function RE_GetTopRated(topN)
    Dim rs, sql
    
    If topN <= 0 Then topN = 8
    
    On Error Resume Next
    
    sql = "SELECT TOP " & topN & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType, " & _
          "ISNULL(AVG(CAST(pr.Rating AS FLOAT)), 0) AS AvgRating, " & _
          "ISNULL(COUNT(pr.ReviewID), 0) AS ReviewCount " & _
          "FROM Products p " & _
          "LEFT JOIN ProductReviews pr ON p.ProductID = pr.ProductID AND pr.Status = 'Approved' " & _
          "WHERE p.IsActive = 1 " & _
          "GROUP BY p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType " & _
          "ORDER BY AvgRating DESC, ReviewCount DESC"
    
    Set rs = conn.Execute(sql)
    Set RE_GetTopRated = rs
    ' 调用方需关闭Recordset
End Function

' ============================================
' 获取最畅销产品（按销量排序）
' ============================================
Function RE_GetPopularProducts(topN)
    Dim rs, sql
    
    If topN <= 0 Then topN = 8
    
    On Error Resume Next
    
    sql = "SELECT TOP " & topN & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType, " & _
          "ISNULL(SUM(od.Quantity), 0) AS TotalSold " & _
          "FROM Products p " & _
          "LEFT JOIN OrderDetails od ON p.ProductID = od.ProductID " & _
          "LEFT JOIN Orders o ON od.OrderID = o.OrderID AND o.Status = 'Paid' " & _
          "WHERE p.IsActive = 1 " & _
          "GROUP BY p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType " & _
          "ORDER BY TotalSold DESC"
    
    Set rs = conn.Execute(sql)
    Set RE_GetPopularProducts = rs
End Function

' ============================================
' 获取新品推荐（最近创建的产品）
' ============================================
Function RE_GetNewProducts(topN)
    Dim rs, sql
    
    If topN <= 0 Then topN = 4
    
    On Error Resume Next
    
    sql = "SELECT TOP " & topN & " ProductID, ProductName, BasePrice, ImageURL, ProductType, CreatedAt " & _
          "FROM Products WHERE IsActive = 1 ORDER BY CreatedAt DESC"
    
    Set rs = conn.Execute(sql)
    Set RE_GetNewProducts = rs
End Function

' ============================================
' 渲染推荐产品HTML（在ASP页面中嵌入）
' ============================================
Sub RE_RenderRecommendations(rs, className, showPrice)
    Dim count
    count = 0
    
    If rs Is Nothing Then
        Response.Write "<p style=""color:#888;text-align:center;padding:20px;"">暂无推荐</p>"
        Exit Sub
    End If
    
    If rs.EOF Then
        Response.Write "<p style=""color:#888;text-align:center;padding:20px;"">暂无推荐</p>"
        Exit Sub
    End If
    
    Response.Write "<div class=""recommendation-grid " & className & """>"
    
    Do While Not rs.EOF
        Dim pid, pname, pprice, pimg, ptype
        pid = rs("ProductID")
        pname = rs("ProductName")
        pprice = rs("BasePrice")
        pimg = rs("ImageURL")
        ptype = rs("ProductType")
        
        If IsNull(pimg) Or pimg = "" Then pimg = "/images/default-product.svg"
        
        ' 产品类型标签
        Dim typeLabel
        Select Case ptype
            Case "standard": typeLabel = "品牌"
            Case "Custom": typeLabel = "定制"
            Case "KOL": typeLabel = "KOL推荐"
            Case Else: typeLabel = ""
        End Select
%>
    <a href="/product.asp?id=<%= pid %>" class="rec-card">
        <div class="rec-img-wrapper">
            <img src="<%= Server.HTMLEncode(pimg) %>" alt="<%= Server.HTMLEncode(pname) %>" loading="lazy">
            <% If typeLabel <> "" Then %><span class="rec-badge"><%= typeLabel %></span><% End If %>
        </div>
        <div class="rec-info">
            <h4><%= Server.HTMLEncode(pname) %></h4>
            <% If showPrice Then %>
            <span class="rec-price">¥<%= FormatNumber(pprice, 2) %></span>
            <% End If %>
        </div>
    </a>
<%
        count = count + 1
        rs.MoveNext
    Loop
    
    Response.Write "</div>"
End Sub

' JavaScript转义
Function JSEncode(str)
    If IsNull(str) Then JSEncode = "" : Exit Function
    JSEncode = Replace(str, "\", "\\")
    JSEncode = Replace(JSEncode, """", "\""")
    JSEncode = Replace(JSEncode, "'", "\'")
    JSEncode = Replace(JSEncode, vbCrLf, "")
    JSEncode = Replace(JSEncode, vbLf, "")
End Function

' ============================================
' V18: RE_FetchProductsByIds - 根据ID数组获取产品Recordset
' 将AI服务返回的产品ID列表转为数据库查询结果
' ============================================
Function RE_FetchProductsByIds(productIds)
    If Not IsArray(productIds) Then
        Set RE_FetchProductsByIds = Nothing
        Exit Function
    End If
    If UBound(productIds) < 0 Then
        Set RE_FetchProductsByIds = Nothing
        Exit Function
    End If
    
    Dim idList, i, id
    idList = ""
    For i = 0 To UBound(productIds)
        id = productIds(i)
        If IsNumeric(id) And CLng(id) > 0 Then
            If idList <> "" Then idList = idList & ","
            idList = idList & CLng(id)
        End If
    Next
    
    If idList = "" Then
        Set RE_FetchProductsByIds = Nothing
        Exit Function
    End If
    
    Dim sql
    sql = "SELECT ProductID, ProductName, BasePrice, ImageURL, ProductType, Category, Description " & _
          "FROM Products WHERE IsActive <> 0 AND ProductID IN (" & idList & ") " & _
          "ORDER BY CHARINDEX(',' + CAST(ProductID AS VARCHAR) + ',', '," & idList & ",')"
    
    On Error Resume Next
    Set RE_FetchProductsByIds = conn.Execute(sql)
    On Error GoTo 0
    
    ' 调用方需关闭Recordset
End Function

' ============================================
' V18: RE_GetPersonalizedProducts - AI个性化推荐
' 优先使用AI微服务，不可用时回退到传统SQL推荐
' 参数: userId - 用户ID, limit - 推荐数量(默认6)
' 返回: Recordset (调用方需关闭)
' ============================================
Function RE_GetPersonalizedProducts(userId, limit)
    If limit <= 0 Then limit = 6
    
    On Error Resume Next
    
    ' 策略1: AI推荐
    If FEATURE_AI_RECOMMENDATIONS Then
        Dim aiIds
        aiIds = AI_GetPersonalized(userId, limit, Empty)
        If IsArray(aiIds) Then
            If UBound(aiIds) >= 0 Then
                Dim rsAI
                Set rsAI = RE_FetchProductsByIds(aiIds)
                If Not rsAI Is Nothing Then
                    If Not rsAI.EOF Then
                        Set RE_GetPersonalizedProducts = rsAI
                        Exit Function
                    End If
                    rsAI.Close
                    Set rsAI = Nothing
                End If
            End If
        End If
    End If
    
    ' 策略2: 回退到基于购买历史的推荐
    Dim dictResult, rsFallback
    Set dictResult = RE_GetUserRecommendations(userId, limit)
    If dictResult.Count > 0 Then
        ' 从Dictionary提取产品ID并查询数据库
        Dim fbIds(), j, pid
        j = 0
        ReDim fbIds(dictResult.Count - 1)
        Dim key
        For Each key In dictResult.Keys()
            fbIds(j) = CLng(key)
            j = j + 1
        Next
        Set rsFallback = RE_FetchProductsByIds(fbIds)
        If Not rsFallback Is Nothing Then
            If Not rsFallback.EOF Then
                Set RE_GetPersonalizedProducts = rsFallback
                Exit Function
            End If
            rsFallback.Close
            Set rsFallback = Nothing
        End If
    End If
    
    ' 策略3: 最终回退到热门产品
    Set RE_GetPersonalizedProducts = RE_GetPopularProducts(limit)
End Function

' ============================================
' V18: RE_GetSimilarFragrances - AI相似产品推荐
' 优先使用AI微服务，不可用时回退到传统同类型推荐
' 参数: productId - 参考产品ID, limit - 推荐数量(默认6)
' 返回: Recordset (调用方需关闭)
' ============================================
Function RE_GetSimilarFragrances(productId, limit)
    If limit <= 0 Then limit = 6
    
    On Error Resume Next
    
    ' 策略1: AI相似推荐
    If FEATURE_AI_RECOMMENDATIONS Then
        Dim aiIds
        aiIds = AI_GetSimilarProducts(productId, limit)
        If IsArray(aiIds) Then
            If UBound(aiIds) >= 0 Then
                Dim rsAI
                Set rsAI = RE_FetchProductsByIds(aiIds)
                If Not rsAI Is Nothing Then
                    If Not rsAI.EOF Then
                        Set RE_GetSimilarFragrances = rsAI
                        Exit Function
                    End If
                    rsAI.Close
                    Set rsAI = Nothing
                End If
            End If
        End If
    End If
    
    ' 策略2: 回退到传统同类型推荐
    Dim productType, rsType
    productType = ""
    Set rsType = conn.Execute("SELECT ProductType FROM Products WHERE ProductID = " & CLng(productId))
    If Not rsType Is Nothing Then
        If Not rsType.EOF Then
            productType = rsType("ProductType") & ""
        End If
        rsType.Close
    End If
    Set rsType = Nothing
    Set RE_GetSimilarFragrances = RE_GetRelatedProducts(productId, CStr(productType), limit)
End Function

' ============================================
' V18: RE_GetTrendingNow - AI趋势推荐
' 优先使用AI微服务趋势数据，不可用时回退到畅销排行
' 参数: limit - 推荐数量(默认8)
' 返回: Recordset (调用方需关闭)
' ============================================
Function RE_GetTrendingNow(limit)
    If limit <= 0 Then limit = 8
    
    On Error Resume Next
    
    ' 策略1: AI趋势推荐
    If FEATURE_AI_RECOMMENDATIONS Then
        Dim result, dataDict, trendIds(), i
        Set result = AI_CallServiceGET("recommend/trending?limit=" & limit)
        If Not IsEmpty(result) And IsObject(result) Then
            If result.Exists("data") Then
                Set dataDict = result("data")
                If IsObject(dataDict) And dataDict.Count > 0 Then
                    ReDim trendIds(dataDict.Count - 1)
                    i = 0
                    Dim idx
                    For idx = 0 To dataDict.Count - 1
                        If IsObject(dataDict(idx)) Then
                            If dataDict(idx).Exists("product_id") Then
                                trendIds(i) = CLng(dataDict(idx)("product_id"))
                                i = i + 1
                            End If
                        End If
                    Next
                    If i > 0 Then
                        ReDim Preserve trendIds(i - 1)
                        Dim rsAI
                        Set rsAI = RE_FetchProductsByIds(trendIds)
                        If Not rsAI Is Nothing Then
                            If Not rsAI.EOF Then
                                Set RE_GetTrendingNow = rsAI
                                Exit Function
                            End If
                            rsAI.Close
                            Set rsAI = Nothing
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ' 策略2: 回退到畅销排行
    Set RE_GetTrendingNow = RE_GetPopularProducts(limit)
End Function

' ============================================
' V18: RE_RenderRecommendationsFromRS - 安全渲染推荐（自动处理空结果）
' 对 RE_RenderRecommendations 的增强包装，增加空结果友好提示
' ============================================
Sub RE_RenderRecommendationsSafe(rs, className, showPrice, emptyMessage)
    If rs Is Nothing Then
        Response.Write "<p style=""color:#999;text-align:center;padding:20px;font-size:14px;"">" & emptyMessage & "</p>"
        Exit Sub
    End If
    If rs.EOF Then
        Response.Write "<p style=""color:#999;text-align:center;padding:20px;font-size:14px;"">" & emptyMessage & "</p>"
        Exit Sub
    End If
    
    Call RE_RenderRecommendations(rs, className, showPrice)
End Sub
%>