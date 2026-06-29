<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/api_response.asp"-->
<!--#include file="../includes/ai_client.asp"-->
<%
' ============================================
' V18 智能香氛匹配 API
' POST /api/fragrance_match.asp
' Body: JSON { "answers": { "style":"floral", "occasion":"daily", ... } }
' 返回: 匹配的香调家族、推荐香调组合、强度建议
' ============================================

' 只接受 POST
If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
    Call API_Error(API_ERR_PARAM_INVALID, "仅支持POST请求")
    Response.End
End If

' 读取请求体
Dim requestBody
requestBody = ""
On Error Resume Next
Dim byteCount
byteCount = Request.TotalBytes
If byteCount > 0 Then
    requestBody = BytesToStr(Request.BinaryRead(byteCount))
End If
On Error GoTo 0

If requestBody = "" Then
    Call API_Error(API_ERR_PARAM_MISSING, "请求体为空，请提供answers参数")
    Response.End
End If

' 提取 answers 子对象
Dim answersDict
Set answersDict = FM_ExtractAnswers(requestBody)

If answersDict Is Nothing Then
    ' 尝试从 form 数据获取
    Dim style, occasion, season, gender, intensity, budget
    style = Request.Form("style")
    occasion = Request.Form("occasion")
    season = Request.Form("season")
    gender = Request.Form("gender")
    intensity = Request.Form("intensity")
    budget = Request.Form("budget")
    
    If style = "" Then
        Call API_Error(API_ERR_PARAM_MISSING, "缺少必填参数: style (香型偏好)")
        Response.End
    End If
    
    Set answersDict = Server.CreateObject("Scripting.Dictionary")
    answersDict.Add "style", style
    If occasion <> "" Then answersDict.Add "occasion", occasion
    If season <> "" Then answersDict.Add "season", season
    If gender <> "" Then answersDict.Add "gender", gender
    If intensity <> "" Then answersDict.Add "intensity", intensity
    If budget <> "" Then answersDict.Add "budget", budget
End If

' 尝试 AI 匹配
Dim result
Set result = Nothing

If FEATURE_AI_FRAGRANCE_MATCH And FEATURE_AI_RECOMMENDATIONS Then
    Set result = AI_MatchFragrance(answersDict)
End If

' AI 回退：本地匹配
If IsEmpty(result) Or (Not IsObject(result)) Then
    Set result = FM_LocalMatch(answersDict)
End If

' 返回结果
Call API_Success(result, "匹配成功")

' ============================================
' FM_ExtractAnswers: 从 JSON 请求体提取 answers 对象
' ============================================
Function FM_ExtractAnswers(body)
    Dim regex, matches, match, subJson
    Set FM_ExtractAnswers = Nothing
    
    Set regex = New RegExp
    regex.Pattern = """answers""\s*:\s*\{"
    regex.IgnoreCase = True
    Set matches = regex.Execute(body)
    
    If matches.Count = 0 Then
        Set regex = Nothing
        Exit Function
    End If
    
    Dim startPos, braceCount, i, ch
    startPos = matches(0).FirstIndex + Len(matches(0).Value)
    braceCount = 1
    
    For i = startPos To Len(body)
        ch = Mid(body, i, 1)
        If ch = "{" Then braceCount = braceCount + 1
        If ch = "}" Then
            braceCount = braceCount - 1
            If braceCount = 0 Then
                subJson = Mid(body, startPos, i - startPos + 1)
                Exit For
            End If
        End If
    Next
    
    If subJson <> "" Then
        Set FM_ExtractAnswers = AI_ParseJsonObject(subJson)
    End If
    
    Set regex = Nothing
End Function

' ============================================
' FM_LocalMatch: 本地香氛匹配（AI 不可用时的回退逻辑）
' 实现与 fragrance_matcher.py 相同的评分算法
' ============================================
Function FM_LocalMatch(answers)
    Dim scores, style, occasion, season, gender, intensity, budget
    Set scores = Server.CreateObject("Scripting.Dictionary")
    
    ' 获取答案
    style = FM_GetAnswer(answers, "style")
    occasion = FM_GetAnswer(answers, "occasion")
    season = FM_GetAnswer(answers, "season")
    gender = FM_GetAnswer(answers, "gender")
    intensity = FM_GetAnswer(answers, "intensity")
    budget = FM_GetAnswer(answers, "budget")
    
    ' 香调家族权重
    Dim noteWeights
    Set noteWeights = Server.CreateObject("Scripting.Dictionary")
    noteWeights.Add "floral", 0.9
    noteWeights.Add "citrus", 0.9
    noteWeights.Add "woody", 0.8
    noteWeights.Add "oriental", 0.8
    noteWeights.Add "fresh", 0.85
    noteWeights.Add "fruity", 0.75
    noteWeights.Add "green", 0.7
    
    ' 风格偏好 (权重 2.0)
    If style <> "" And noteWeights.Exists(style) Then
        FM_AddScore scores, style, 2.0
    End If
    
    ' 场合匹配 (权重 1.5)
    FM_AddOccasionScores scores, occasion, 1.5
    
    ' 季节匹配 (权重 1.0)
    FM_AddSeasonScores scores, season, 1.0
    
    ' 性别偏好 (权重 1.2)
    FM_AddGenderScores scores, gender, 1.2
    
    ' 应用家族权重
    Dim key
    For Each key In scores.Keys
        If noteWeights.Exists(key) Then
            scores(key) = scores(key) * noteWeights(key)
        End If
    Next
    
    ' 排序（冒泡排序）
    Dim families, famKeys, i, j, tmpScore, tmpKey
    ReDim famKeys(scores.Count - 1)
    i = 0
    For Each key In scores.Keys
        famKeys(i) = key
        i = i + 1
    Next
    
    For i = 0 To UBound(famKeys) - 1
        For j = i + 1 To UBound(famKeys)
            If scores(famKeys(i)) < scores(famKeys(j)) Then
                tmpKey = famKeys(i)
                famKeys(i) = famKeys(j)
                famKeys(j) = tmpKey
            End If
        Next
    Next
    
    ' 构建匹配家族数组
    Dim matchedFamilies(), mfIdx
    mfIdx = 0
    Dim maxFamilies
    maxFamilies = 3
    If scores.Count < 3 Then maxFamilies = scores.Count
    ReDim matchedFamilies(maxFamilies - 1)
    For i = 0 To maxFamilies - 1
        Dim famKey, famDict
        famKey = famKeys(i)
        Set famDict = Server.CreateObject("Scripting.Dictionary")
        famDict.Add "family", famKey
        famDict.Add "score", Round(scores(famKey), 2)
        famDict.Add "keywords", FM_GetFamilyKeywords(famKey)
        Set matchedFamilies(i) = famDict
    Next
    
    ' 推荐香调组合
    Dim recommendedNotes
    Set recommendedNotes = FM_GetNoteRecommendations(famKeys, maxFamilies)
    
    ' 强度建议
    Dim intensityAdvice
    Select Case LCase(intensity)
        Case "light":   intensityAdvice = "建议选择EDT淡香水，清新不张扬"
        Case "medium":  intensityAdvice = "建议选择EDP淡香精，持久适中"
        Case "strong":  intensityAdvice = "建议选择Parfum浓香精，持久浓郁"
        Case Else:      intensityAdvice = "建议选择EDP淡香精，持久适中"
    End Select
    
    ' 构建返回结果
    Dim retDict
    Set retDict = Server.CreateObject("Scripting.Dictionary")
    Set retDict("matched_families") = FM_ArrayToDict(matchedFamilies)
    Set retDict("recommended_notes") = recommendedNotes
    retDict.Add "intensity_advice", intensityAdvice
    retDict.Add "budget_level", budget
    
    Set FM_LocalMatch = retDict
End Function

' ============================================
' Helper Functions
' ============================================
Function FM_GetAnswer(dict, key)
    If IsObject(dict) Then
        If dict.Exists(key) Then
            FM_GetAnswer = dict(key)
            Exit Function
        End If
    End If
    FM_GetAnswer = ""
End Function

Sub FM_AddScore(scores, family, value)
    If scores.Exists(family) Then
        scores(family) = scores(family) + value
    Else
        scores.Add family, value
    End If
End Sub

Sub FM_AddOccasionScores(scores, occasion, value)
    Dim families
    Select Case LCase(occasion)
        Case "daily":  families = Array("fresh", "citrus", "green")
        Case "work":   families = Array("woody", "green", "fresh")
        Case "date":   families = Array("floral", "oriental", "fruity")
        Case "party":  families = Array("oriental", "fruity", "floral")
        Case "sport":  families = Array("fresh", "citrus")
        Case "formal": families = Array("woody", "oriental", "floral")
        Case Else:     Exit Sub
    End Select
    Dim f
    For Each f In families
        FM_AddScore scores, f, value
    Next
End Sub

Sub FM_AddSeasonScores(scores, season, value)
    Dim families
    Select Case LCase(season)
        Case "spring": families = Array("floral", "green", "fruity")
        Case "summer": families = Array("citrus", "fresh", "fruity")
        Case "autumn": families = Array("woody", "oriental")
        Case "winter": families = Array("oriental", "woody", "floral")
        Case Else:     Exit Sub
    End Select
    Dim f
    For Each f In families
        FM_AddScore scores, f, value
    Next
End Sub

Sub FM_AddGenderScores(scores, gender, value)
    Dim families
    Select Case LCase(gender)
        Case "female": families = Array("floral", "fruity", "oriental")
        Case "male":   families = Array("woody", "fresh", "citrus")
        Case "unisex": families = Array("citrus", "green", "woody")
        Case Else:     Exit Sub
    End Select
    Dim f
    For Each f In families
        FM_AddScore scores, f, value
    Next
End Sub

Function FM_GetFamilyKeywords(family)
    Dim kw
    Select Case LCase(family)
        Case "floral":   kw = Array("花香", "玫瑰", "茉莉", "百合", "温柔", "浪漫", "女性")
        Case "citrus":   kw = Array("柑橘", "柠檬", "清新", "活力", "阳光", "清爽", "夏天")
        Case "woody":    kw = Array("木质", "檀香", "雪松", "沉稳", "成熟", "中性", "秋天")
        Case "oriental": kw = Array("东方", "琥珀", "香草", "神秘", "性感", "浓郁", "夜晚")
        Case "fresh":    kw = Array("海洋", "水生", "绿叶", "运动", "干净", "春天", "日常")
        Case "fruity":   kw = Array("果香", "桃子", "莓果", "甜美", "活泼", "年轻", "派对")
        Case "green":    kw = Array("青草", "绿茶", "自然", "素雅", "文艺", "中性")
        Case Else:       kw = Array()
    End Select
    FM_GetFamilyKeywords = kw
End Function

Function FM_GetNoteRecommendations(famKeys, count)
    Dim noteMap, retDict
    Set noteMap = Server.CreateObject("Scripting.Dictionary")
    
    ' floral
    Dim floralNotes
    Set floralNotes = Server.CreateObject("Scripting.Dictionary")
    floralNotes.Add "top", Array("佛手柑", "粉红胡椒")
    floralNotes.Add "middle", Array("玫瑰", "茉莉", "鸢尾花")
    floralNotes.Add "base", Array("麝香", "琥珀")
    Set noteMap("floral") = floralNotes
    
    ' citrus
    Dim citrusNotes
    Set citrusNotes = Server.CreateObject("Scripting.Dictionary")
    citrusNotes.Add "top", Array("柠檬", "葡萄柚", "佛手柑")
    citrusNotes.Add "middle", Array("橙花", "薄荷")
    citrusNotes.Add "base", Array("雪松", "白麝香")
    Set noteMap("citrus") = citrusNotes
    
    ' woody
    Dim woodyNotes
    Set woodyNotes = Server.CreateObject("Scripting.Dictionary")
    woodyNotes.Add "top", Array("香柠檬", "胡椒")
    woodyNotes.Add "middle", Array("雪松", "檀香木")
    woodyNotes.Add "base", Array("香根草", "广藿香", "皮革")
    Set noteMap("woody") = woodyNotes
    
    ' oriental
    Dim orientalNotes
    Set orientalNotes = Server.CreateObject("Scripting.Dictionary")
    orientalNotes.Add "top", Array("肉桂", "小豆蔻")
    orientalNotes.Add "middle", Array("琥珀", "香草", "零陵香豆")
    orientalNotes.Add "base", Array("檀香", "麝香", "广藿香")
    Set noteMap("oriental") = orientalNotes
    
    ' fresh
    Dim freshNotes
    Set freshNotes = Server.CreateObject("Scripting.Dictionary")
    freshNotes.Add "top", Array("柑橘", "海洋")
    freshNotes.Add "middle", Array("薰衣草", "迷迭香")
    freshNotes.Add "base", Array("白麝香", "苔藓")
    Set noteMap("fresh") = freshNotes
    
    ' fruity
    Dim fruityNotes
    Set fruityNotes = Server.CreateObject("Scripting.Dictionary")
    fruityNotes.Add "top", Array("桃子", "黑加仑")
    fruityNotes.Add "middle", Array("玫瑰", "紫罗兰")
    fruityNotes.Add "base", Array("香草", "麝香")
    Set noteMap("fruity") = fruityNotes
    
    ' green
    Dim greenNotes
    Set greenNotes = Server.CreateObject("Scripting.Dictionary")
    greenNotes.Add "top", Array("佛手柑", "绿叶")
    greenNotes.Add "middle", Array("绿茶", "茉莉")
    greenNotes.Add "base", Array("白麝香", "雪松")
    Set noteMap("green") = greenNotes
    
    ' 收集前2个匹配家族的香调
    Set retDict = Server.CreateObject("Scripting.Dictionary")
    retDict.Add "top", Array()
    retDict.Add "middle", Array()
    retDict.Add "base", Array()
    
    Dim topNotes, middleNotes, baseNotes
    ReDim topNotes(0)
    ReDim middleNotes(0)
    ReDim baseNotes(0)
    
    Dim maxFam, fi, fam, layerNotes, layer, note, topIdx, midIdx, baseIdx
    maxFam = count
    If maxFam > 2 Then maxFam = 2
    topIdx = 0
    midIdx = 0
    baseIdx = 0
    
    For fi = 0 To maxFam - 1
        fam = LCase(famKeys(fi))
        If noteMap.Exists(fam) Then
            Set layerNotes = noteMap(fam)
            
            ' top notes
            For Each note In layerNotes("top")
                If Not FM_InArray(topNotes, note) Then
                    If topIdx > 0 Then ReDim Preserve topNotes(topIdx)
                    topNotes(topIdx) = note
                    topIdx = topIdx + 1
                End If
            Next
            
            ' middle notes
            For Each note In layerNotes("middle")
                If Not FM_InArray(middleNotes, note) Then
                    If midIdx > 0 Then ReDim Preserve middleNotes(midIdx)
                    middleNotes(midIdx) = note
                    midIdx = midIdx + 1
                End If
            Next
            
            ' base notes
            For Each note In layerNotes("base")
                If Not FM_InArray(baseNotes, note) Then
                    If baseIdx > 0 Then ReDim Preserve baseNotes(baseIdx)
                    baseNotes(baseIdx) = note
                    baseIdx = baseIdx + 1
                End If
            Next
        End If
    Next
    
    ' 限制每层最多3个
    FM_LimitArray topNotes, 3
    FM_LimitArray middleNotes, 3
    FM_LimitArray baseNotes, 3
    
    retDict("top") = topNotes
    retDict("middle") = middleNotes
    retDict("base") = baseNotes
    
    Set FM_GetNoteRecommendations = retDict
End Function

Function FM_InArray(arr, val)
    If Not IsArray(arr) Then
        FM_InArray = False
        Exit Function
    End If
    Dim i
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            FM_InArray = True
            Exit Function
        End If
    Next
    FM_InArray = False
End Function

Sub FM_LimitArray(arr, maxCount)
    If Not IsArray(arr) Then Exit Sub
    If UBound(arr) + 1 <= maxCount Then Exit Sub
    ReDim Preserve arr(maxCount - 1)
End Sub

Function FM_ArrayToDict(arr)
    If Not IsArray(arr) Then
        Set FM_ArrayToDict = Nothing
        Exit Function
    End If
    Dim dict, i
    Set dict = Server.CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(arr)
        dict.Add i, arr(i)
    Next
    Set FM_ArrayToDict = dict
End Function

' ============================================
' BytesToStr: 将二进制数据转为字符串
' ============================================
Function BytesToStr(bytes)
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    BytesToStr = stream.ReadText
    stream.Close
    Set stream = Nothing
End Function
%>
