<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/dal_products.asp"-->
<!--#include file="../includes/api_response.asp"-->
<!--#include file="../includes/api_guard.asp"-->
<%
' ============================================
' V18.0 智能搜索建议 API (Smart Search Suggestions)
' V18新增: 拼音搜索、模糊容错(编辑距离≤2)、同义词扩展、搜索意图识别
' 用法: GET /api/search_suggestions.asp?q=关键词&type=suggest&max=8
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

Call OpenConnection()

' V18: API 守卫（速率限制）
If Not API_Guard("api", False) Then Response.End

Dim q, searchType, maxResults
q = Trim(Request.QueryString("q"))
searchType = Trim(Request.QueryString("type"))
If searchType = "" Then searchType = "suggest"
maxResults = CInt(Request.QueryString("max"))
If maxResults < 1 Or maxResults > 20 Then maxResults = 8

Select Case searchType
    Case "suggest"
        If q <> "" Then
            ' V18: 智能搜索增强
            If FEATURE_AI_SEARCH Then
                Call SEARCH_SuggestSmart(q, maxResults)
            Else
                Dim rsSuggest
                Set rsSuggest = DAL_Products_GetSuggestions(q, maxResults)
                If Not rsSuggest Is Nothing Then
                    Dim suggestions, itemCount
                    suggestions = "["
                    itemCount = 0
                    Do While Not rsSuggest.EOF
                        If itemCount > 0 Then suggestions = suggestions & ","
                        suggestions = suggestions & "{"
                        suggestions = suggestions & """id"":" & rsSuggest("ProductID")
                        suggestions = suggestions & ",""name"":" & API_JsonEncode(rsSuggest("ProductName"))
                        suggestions = suggestions & ",""price"":" & CDbl(rsSuggest("BasePrice"))
                        suggestions = suggestions & ",""image"":" & API_JsonEncode(rsSuggest("ImageURL"))
                        suggestions = suggestions & ",""type"":" & API_JsonEncode(rsSuggest("ProductType"))
                        suggestions = suggestions & "}"
                        itemCount = itemCount + 1
                        rsSuggest.MoveNext
                    Loop
                    suggestions = suggestions & "]"
                    rsSuggest.Close
                    Set rsSuggest = Nothing
                    If Session("UserID") <> "" Then
                        Call DAL_Products_RecordSearch(Session("UserID"), q)
                    End If
                    Call API_Success(suggestions, "获取成功")
                Else
                    Call API_Success("[]", "无匹配结果")
                End If
            End If
        Else
            Call API_Success("[]", "请输入搜索关键词")
        End If

    Case "history"
        If Session("UserID") <> "" Then
            Dim rsHistory
            Set rsHistory = DAL_Products_GetSearchHistory(Session("UserID"), maxResults)
            If Not rsHistory Is Nothing Then
                Dim historyItems, hCount
                historyItems = "["
                hCount = 0
                Do While Not rsHistory.EOF
                    If hCount > 0 Then historyItems = historyItems & ","
                    historyItems = historyItems & "{"
                    historyItems = historyItems & """keyword"":" & API_JsonEncode(rsHistory("Keyword"))
                    historyItems = historyItems & ",""count"":" & CLng(rsHistory("SearchCount"))
                    historyItems = historyItems & "}"
                    hCount = hCount + 1
                    rsHistory.MoveNext
                Loop
                historyItems = historyItems & "]"
                rsHistory.Close
                Set rsHistory = Nothing
                Call API_Success(historyItems, "获取成功")
            Else
                Call API_Success("[]", "暂无搜索历史")
            End If
        Else
            Call API_Error(API_ERR_AUTH_REQUIRED, "请先登录")
        End If

    Case "clear_history"
        If Session("UserID") <> "" Then
            Dim clearParams(0)
            clearParams(0) = Array("@UserID", DAL_adInteger, 0, CLng(Session("UserID")))
            DAL_Execute "DELETE FROM SearchHistory WHERE UserID=@UserID", clearParams
            Call API_Success(Null, "搜索历史已清空")
        Else
            Call API_Error(API_ERR_AUTH_REQUIRED, "请先登录")
        End If

    Case Else
        Call API_Error(API_ERR_PARAM_INVALID, "未知查询类型: " & searchType)
End Select

' ============================================
' V18 智能搜索函数库
' ============================================

' 主入口：智能搜索建议
Sub SEARCH_SuggestSmart(keyword, maxResults)
    Dim expanded, rsCombined, suggestions, itemCount, seenIds
    Set seenIds = Server.CreateObject("Scripting.Dictionary")
    suggestions = "["
    itemCount = 0
    
    ' 1. 搜索意图识别
    Dim intent
    intent = SEARCH_DetectIntent(keyword)
    
    ' 2. 同义词扩展
    expanded = SEARCH_ExpandSynonyms(keyword)
    
    ' 3. 拼音匹配
    Dim pinyinChinese
    pinyinChinese = SEARCH_PinyinToChinese(keyword)
    
    ' 4. 构建多关键词查询
    Dim searchTerms, i, term, rs, safeKw
    ReDim searchTerms(0)
    searchTerms(0) = keyword
    
    ' 添加扩展词
    If IsArray(expanded) Then
        For i = 0 To UBound(expanded)
            If expanded(i) <> "" And expanded(i) <> keyword Then
                Dim foundDup
                foundDup = False
                Dim di
                For di = 0 To UBound(searchTerms)
                    If LCase(searchTerms(di)) = LCase(expanded(i)) Then foundDup = True
                Next
                If Not foundDup Then
                    Dim newSize
                    newSize = UBound(searchTerms) + 1
                    ReDim Preserve searchTerms(newSize)
                    searchTerms(newSize) = expanded(i)
                End If
            End If
        Next
    End If
    
    ' 添加拼音匹配结果
    If pinyinChinese <> "" Then
        Dim pcTerms, pci
        pcTerms = Split(pinyinChinese, ",")
        For pci = 0 To UBound(pcTerms)
            Dim pcTerm
            pcTerm = Trim(pcTerms(pci))
            If pcTerm <> "" And pcTerm <> keyword Then
                Dim foundDup2
                foundDup2 = False
                Dim di2
                For di2 = 0 To UBound(searchTerms)
                    If LCase(searchTerms(di2)) = LCase(pcTerm) Then foundDup2 = True
                Next
                If Not foundDup2 Then
                    Dim newSize2
                    newSize2 = UBound(searchTerms) + 1
                    ReDim Preserve searchTerms(newSize2)
                    searchTerms(newSize2) = pcTerm
                End If
            End If
        Next
    End If
    
    ' 5. 查询数据库（先用原始关键词）
    For i = 0 To UBound(searchTerms)
        If itemCount >= maxResults Then Exit For
        term = Trim(searchTerms(i))
        If term = "" Then
            ' nothing
        Else
            safeKw = SafeLike(Left(term, 100))
            If safeKw <> "" Then
                Set rs = SEARCH_QuerySuggestions(safeKw, maxResults)
                If Not rs Is Nothing Then
                    Do While Not rs.EOF And itemCount < maxResults
                        Dim pid
                        pid = CLng(rs("ProductID"))
                        If Not seenIds.Exists(pid) Then
                            seenIds.Add pid, True
                            If itemCount > 0 Then suggestions = suggestions & ","
                            suggestions = suggestions & SEARCH_BuildSuggestItem(rs, "keyword")
                            itemCount = itemCount + 1
                        End If
                        rs.MoveNext
                    Loop
                    rs.Close
                    Set rs = Nothing
                End If
            End If
        End If
    Next
    
    ' 6. 如果结果为空，尝试模糊匹配
    If itemCount = 0 Then
        Dim fuzzyTerm
        fuzzyTerm = SEARCH_FuzzySearch(keyword)
        If fuzzyTerm <> "" And fuzzyTerm <> keyword Then
            safeKw = SafeLike(Left(fuzzyTerm, 100))
            Set rs = SEARCH_QuerySuggestions(safeKw, maxResults)
            If Not rs Is Nothing Then
                Do While Not rs.EOF And itemCount < maxResults
                    pid = CLng(rs("ProductID"))
                    If Not seenIds.Exists(pid) Then
                        seenIds.Add pid, True
                        If itemCount > 0 Then suggestions = suggestions & ","
                        suggestions = suggestions & SEARCH_BuildSuggestItem(rs, "fuzzy")
                        itemCount = itemCount + 1
                    End If
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
            End If
        End If
    End If
    
    ' 7. 意图匹配：如果意图识别为特定类型，追加相关产品
    If itemCount < maxResults And intent <> "" Then
        Dim intentType
        intentType = SEARCH_GetIntentProductType(intent)
        If intentType <> "" Then
            Set rs = SEARCH_QueryByType(intentType, maxResults - itemCount)
            If Not rs Is Nothing Then
                Do While Not rs.EOF And itemCount < maxResults
                    pid = CLng(rs("ProductID"))
                    If Not seenIds.Exists(pid) Then
                        seenIds.Add pid, True
                        If itemCount > 0 Then suggestions = suggestions & ","
                        suggestions = suggestions & SEARCH_BuildSuggestItem(rs, "intent")
                        itemCount = itemCount + 1
                    End If
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
            End If
        End If
    End If
    
    suggestions = suggestions & "]"
    
    ' 记录搜索
    If Session("UserID") <> "" Then
        Call DAL_Products_RecordSearch(Session("UserID"), keyword)
    End If
    
    ' 添加意图元数据
    Dim resultJson
    resultJson = "{""items"":" & suggestions & ",""intent"":" & API_JsonEncode(intent) & "}"
    Call API_Success(resultJson, "获取成功")
End Sub

' 数据库查询：模糊搜索建议
Function SEARCH_QuerySuggestions(safeKeyword, maxResults)
    Dim sql
    sql = "SELECT TOP " & CLng(maxResults) & " ProductID, ProductName, BasePrice, ImageURL, ProductType, Category " & _
          "FROM Products WHERE IsActive=1 AND (ProductName LIKE '%" & safeKeyword & "%' " & _
          "OR Description LIKE '%" & safeKeyword & "%' OR Category LIKE '%" & safeKeyword & "%') " & _
          "ORDER BY CASE WHEN ProductName LIKE '" & safeKeyword & "%' THEN 0 ELSE 1 END, ProductName ASC"
    Set SEARCH_QuerySuggestions = conn.Execute(sql)
End Function

' 按类型查询
Function SEARCH_QueryByType(productType, maxResults)
    Dim sql
    sql = "SELECT TOP " & CLng(maxResults) & " ProductID, ProductName, BasePrice, ImageURL, ProductType, Category " & _
          "FROM Products WHERE IsActive=1 AND ProductType='" & SafeLike(productType) & "' " & _
          "ORDER BY ISNULL(CreatedAt, '2099-12-31') DESC"
    Set SEARCH_QueryByType = conn.Execute(sql)
End Function

' 构建建议项 JSON
Function SEARCH_BuildSuggestItem(rs, matchType)
    Dim item
    item = "{"
    item = item & """id"":" & rs("ProductID")
    item = item & ",""name"":" & API_JsonEncode(rs("ProductName"))
    item = item & ",""price"":" & CDbl(rs("BasePrice"))
    item = item & ",""image"":" & API_JsonEncode(rs("ImageURL"))
    item = item & ",""type"":" & API_JsonEncode(rs("ProductType"))
    If matchType <> "" Then
        item = item & ",""match_type"":" & API_JsonEncode(matchType)
    End If
    item = item & "}"
    SEARCH_BuildSuggestItem = item
End Function

' ============================================
' 搜索意图识别
' 输入: "送女友" → 输出: "gift_female"
' 输入: "夏天清新" → 输出: "season_summer"
' ============================================
Function SEARCH_DetectIntent(keyword)
    Dim kw
    kw = LCase(keyword)
    
    ' 送礼意图
    If InStr(kw, "送女友") > 0 Or InStr(kw, "送女生") > 0 Or InStr(kw, "女朋友") > 0 Or InStr(kw, "老婆") > 0 Or InStr(kw, "女生") > 0 Then
        SEARCH_DetectIntent = "gift_female"
        Exit Function
    End If
    If InStr(kw, "送男友") > 0 Or InStr(kw, "送男生") > 0 Or InStr(kw, "男朋友") > 0 Or InStr(kw, "老公") > 0 Or InStr(kw, "男生") > 0 Then
        SEARCH_DetectIntent = "gift_male"
        Exit Function
    End If
    If InStr(kw, "送礼") > 0 Or InStr(kw, "礼物") > 0 Or InStr(kw, "送人") > 0 Or InStr(kw, "送朋友") > 0 Then
        SEARCH_DetectIntent = "gift_general"
        Exit Function
    End If
    
    ' 季节意图
    If InStr(kw, "夏天") > 0 Or InStr(kw, "夏季") > 0 Or InStr(kw, "夏日") > 0 Then
        SEARCH_DetectIntent = "season_summer"
        Exit Function
    End If
    If InStr(kw, "冬天") > 0 Or InStr(kw, "冬季") > 0 Or InStr(kw, "冬日") > 0 Then
        SEARCH_DetectIntent = "season_winter"
        Exit Function
    End If
    If InStr(kw, "春天") > 0 Or InStr(kw, "春季") > 0 Then
        SEARCH_DetectIntent = "season_spring"
        Exit Function
    End If
    If InStr(kw, "秋天") > 0 Or InStr(kw, "秋季") > 0 Then
        SEARCH_DetectIntent = "season_autumn"
        Exit Function
    End If
    
    ' 场景意图
    If InStr(kw, "上班") > 0 Or InStr(kw, "工作") > 0 Or InStr(kw, "职场") > 0 Or InStr(kw, "商务") > 0 Then
        SEARCH_DetectIntent = "scene_work"
        Exit Function
    End If
    If InStr(kw, "约会") > 0 Or InStr(kw, "派对") > 0 Or InStr(kw, "晚宴") > 0 Then
        SEARCH_DetectIntent = "scene_social"
        Exit Function
    End If
    If InStr(kw, "运动") > 0 Or InStr(kw, "健身") > 0 Or InStr(kw, "户外") > 0 Then
        SEARCH_DetectIntent = "scene_sport"
        Exit Function
    End If
    
    ' 风格意图
    If InStr(kw, "清新") > 0 Or InStr(kw, "清爽") > 0 Or InStr(kw, "干净") > 0 Then
        SEARCH_DetectIntent = "style_fresh"
        Exit Function
    End If
    If InStr(kw, "浓郁") > 0 Or InStr(kw, "持久") > 0 Or InStr(kw, "浓香") > 0 Then
        SEARCH_DetectIntent = "style_strong"
        Exit Function
    End If
    
    SEARCH_DetectIntent = ""
End Function

' 意图映射到产品类型
Function SEARCH_GetIntentProductType(intent)
    Select Case intent
        Case "gift_female", "season_spring": SEARCH_GetIntentProductType = "Custom"
        Case "gift_male":       SEARCH_GetIntentProductType = "KOL"
        Case "scene_social":    SEARCH_GetIntentProductType = "Custom"
        Case "scene_work":      SEARCH_GetIntentProductType = "standard"
        Case "style_fresh":     SEARCH_GetIntentProductType = "Custom"
        Case Else:              SEARCH_GetIntentProductType = ""
    End Select
End Function

' ============================================
' 同义词扩展
' 输入: "清新" → 输出: Array("海洋", "柑橘", "绿茶")
' ============================================
Function SEARCH_ExpandSynonyms(keyword)
    Dim synonyms, kw
    kw = LCase(keyword)
    
    ' 建立同义词映射
    Select Case kw
        Case "清新", "清爽", "干净":
            synonyms = Array("海洋", "柑橘", "绿茶", "薰衣草")
        Case "浓郁", "持久", "浓香":
            synonyms = Array("东方", "木质", "琥珀", "檀香")
        Case "温柔", "柔和", "淡雅":
            synonyms = Array("花香", "玫瑰", "茉莉")
        Case "阳光", "活力", "运动":
            synonyms = Array("柑橘", "柠檬", "海洋")
        Case "性感", "神秘", "魅惑":
            synonyms = Array("东方", "麝香", "琥珀")
        Case "成熟", "稳重", "经典":
            synonyms = Array("木质", "皮革", "雪松")
        Case "甜美", "可爱", "少女":
            synonyms = Array("果香", "花香", "香草")
        Case "中性", "百搭", "日常":
            synonyms = Array("柑橘", "绿叶", "绿茶")
        Case Else:
            synonyms = Array()
    End Select
    
    SEARCH_ExpandSynonyms = synonyms
End Function

' ============================================
' 拼音匹配
' 输入: "hua" → 输出: "花香,花"
' 输入: "ganju" → 输出: "柑橘"
' ============================================
Function SEARCH_PinyinToChinese(keyword)
    Dim kw
    kw = LCase(Trim(keyword))
    
    ' 香调相关拼音映射
    Select Case kw
        Case "hua", "huaxiang":           SEARCH_PinyinToChinese = "花香,玫瑰,茉莉"
        Case "ganju", "juzi":             SEARCH_PinyinToChinese = "柑橘,柠檬"
        Case "muzhi", "mu":               SEARCH_PinyinToChinese = "木质,檀香,雪松"
        Case "dongfang":                  SEARCH_PinyinToChinese = "东方"
        Case "qingxin", "qing":           SEARCH_PinyinToChinese = "清新,海洋"
        Case "haian", "haiyang", "hai":   SEARCH_PinyinToChinese = "海洋"
        Case "meigui", "gui":             SEARCH_PinyinToChinese = "玫瑰"
        Case "moli":                      SEARCH_PinyinToChinese = "茉莉"
        Case "tanxiang", "tan":           SEARCH_PinyinToChinese = "檀香"
        Case "xuesong", "song":           SEARCH_PinyinToChinese = "雪松"
        Case "hupo":                      SEARCH_PinyinToChinese = "琥珀"
        Case "shexiang", "she":           SEARCH_PinyinToChinese = "麝香"
        Case "xiangcao":                  SEARCH_PinyinToChinese = "香草"
        Case "guoxiang", "guo":           SEARCH_PinyinToChinese = "果香,桃子"
        Case "xiangshui":                 SEARCH_PinyinToChinese = "香水"
        Case "dingzhi":                   SEARCH_PinyinToChinese = "定制"
        Case "xiangfen":                  SEARCH_PinyinToChinese = "香氛"
        Case "nishi":                     SEARCH_PinyinToChinese = "女士"
        Case "nanshi":                    SEARCH_PinyinToChinese = "男士"
        Case "zhongxing", "zhong":        SEARCH_PinyinToChinese = "中性"
        Case "lvye", "lvcha", "lv":       SEARCH_PinyinToChinese = "绿叶,绿茶"
        Case "xunyi", "xun":              SEARCH_PinyinToChinese = "薰衣草"
        Case Else:                        SEARCH_PinyinToChinese = ""
    End Select
End Function

' ============================================
' 模糊搜索：在所有候选词中找编辑距离最小的
' ============================================
Function SEARCH_FuzzySearch(keyword)
    Dim candidates, i, bestMatch, bestDist, dist
    candidates = Array("花香", "柑橘", "木质", "东方", "清新", "海洋", "玫瑰", "茉莉", "薰衣草", "檀香", "雪松", "琥珀", "麝香", "香草", "果香", "绿茶", "薄荷", "橙花", "皮革")
    bestMatch = ""
    bestDist = 99
    
    For i = 0 To UBound(candidates)
        dist = SEARCH_LevenshteinDistance(keyword, candidates(i))
        If dist <= 2 And dist < bestDist Then
            bestDist = dist
            bestMatch = candidates(i)
        End If
    Next
    
    SEARCH_FuzzySearch = bestMatch
End Function

' ============================================
' 编辑距离 (Levenshtein Distance)
' ============================================
Function SEARCH_LevenshteinDistance(s1, s2)
    Dim len1, len2, matrix(), i, j, cost
    len1 = Len(s1)
    len2 = Len(s2)
    
    If len1 = 0 Then SEARCH_LevenshteinDistance = len2 : Exit Function
    If len2 = 0 Then SEARCH_LevenshteinDistance = len1 : Exit Function
    
    ReDim matrix(len1, len2)
    
    For i = 0 To len1 : matrix(i, 0) = i : Next
    For j = 0 To len2 : matrix(0, j) = j : Next
    
    For i = 1 To len1
        For j = 1 To len2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            matrix(i, j) = SEARCH_Min3(matrix(i-1, j) + 1, matrix(i, j-1) + 1, matrix(i-1, j-1) + cost)
        Next
    Next
    
    SEARCH_LevenshteinDistance = matrix(len1, len2)
End Function

Function SEARCH_Min3(a, b, c)
    SEARCH_Min3 = a
    If b < SEARCH_Min3 Then SEARCH_Min3 = b
    If c < SEARCH_Min3 Then SEARCH_Min3 = c
End Function

Call CloseConnection()
%>
