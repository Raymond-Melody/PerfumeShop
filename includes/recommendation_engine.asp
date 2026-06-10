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
    Dim rs, sql, dictResult, count, pid, pname
    Set dictResult = Server.CreateObject("Scripting.Dictionary")
    count = 0
    
    If topN <= 0 Then topN = 6
    
    On Error Resume Next
    
    ' 策略1: 基于用户购买历史推荐同类型产品
    sql = "SELECT DISTINCT TOP " & topN & " p2.ProductID, p2.ProductName, p2.BasePrice, p2.ImageURL, p2.ProductType " & _
          "FROM OrderDetails od1 " & _
          "INNER JOIN Products p1 ON od1.ProductID = p1.ProductID " & _
          "INNER JOIN Products p2 ON p1.ProductType = p2.ProductType AND p2.IsActive = 1 " & _
          "INNER JOIN Orders o ON od1.OrderID = o.OrderID " & _
          "WHERE o.UserID = " & userId & " AND p2.ProductID NOT IN " & _
          "(SELECT DISTINCT od2.ProductID FROM OrderDetails od2 INNER JOIN Orders o2 ON od2.OrderID = o2.OrderID WHERE o2.UserID = " & userId & ") " & _
          "AND p2.ProductID NOT IN (SELECT ProductID FROM UserFavorites WHERE UserID = " & userId & ") " & _
          "ORDER BY NEWID()"
    
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        Do While Not rs.EOF And count < topN
            pid = rs("ProductID")
            If Not dictResult.Exists(CStr(pid)) Then
                dictResult.Add CStr(pid), _
                    "{""id"":" & pid & ",""name"":""" & JSEncode(rs("ProductName")) & """,""price"":" & rs("BasePrice") & _
                    ",""img"":""" & JSEncode(rs("ImageURL")) & """,""type"":""" & JSEncode(rs("ProductType")) & """}"
                count = count + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    ' 策略2: 如果不足，补充收藏夹同类产品
    If count < topN Then
        sql = "SELECT DISTINCT TOP " & (topN - count) & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType " & _
              "FROM UserFavorites uf " & _
              "INNER JOIN Products p ON uf.ProductID = p.ProductID " & _
              "WHERE uf.UserID = " & userId & " AND p.IsActive = 1"
        Set rs = conn.Execute(sql)
        If Not rs Is Nothing Then
            Do While Not rs.EOF And count < topN
                pid = rs("ProductID")
                If Not dictResult.Exists(CStr(pid)) Then
                    dictResult.Add CStr(pid), _
                        "{""id"":" & pid & ",""name"":""" & JSEncode(rs("ProductName")) & """,""price"":" & rs("BasePrice") & _
                        ",""img"":""" & JSEncode(rs("ImageURL")) & """,""type"":""" & JSEncode(rs("ProductType")) & """}"
                    count = count + 1
                End If
                rs.MoveNext
            Loop
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    ' 策略3: 如果仍然不足，推荐热门产品
    If count < topN Then
        sql = "SELECT TOP " & (topN - count) & " p.ProductID, p.ProductName, p.BasePrice, p.ImageURL, p.ProductType, " & _
              "ISNULL((SELECT COUNT(*) FROM OrderDetails od INNER JOIN Orders o ON od.OrderID = o.OrderID WHERE od.ProductID = p.ProductID AND o.Status='Paid'), 0) AS SaleCount " & _
              "FROM Products p WHERE p.IsActive = 1 AND p.ProductID NOT IN " & _
              "(SELECT DISTINCT od.ProductID FROM OrderDetails od INNER JOIN Orders o ON od.OrderID = o.OrderID WHERE o.UserID = " & userId & ") " & _
              "ORDER BY SaleCount DESC"
        Set rs = conn.Execute(sql)
        If Not rs Is Nothing Then
            Do While Not rs.EOF And count < topN
                pid = rs("ProductID")
                If Not dictResult.Exists(CStr(pid)) Then
                    dictResult.Add CStr(pid), _
                        "{""id"":" & pid & ",""name"":""" & JSEncode(rs("ProductName")) & """,""price"":" & rs("BasePrice") & _
                        ",""img"":""" & JSEncode(rs("ImageURL")) & """,""type"":""" & JSEncode(rs("ProductType")) & """}"
                    count = count + 1
                End If
                rs.MoveNext
            Loop
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    RE_GetUserRecommendations = dictResult
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
                    RE_GetRelatedProducts = rs
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
            Case "Fixed": typeLabel = "品牌"
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
%>