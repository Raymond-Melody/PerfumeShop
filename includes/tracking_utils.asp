<%
' ============================================
' 用户行为追踪 - Tracking Utils
' 收集用户浏览、搜索、加购等行为数据
' ============================================

' ============================================
' 记录用户行为
' ============================================
Sub TU_LogBehavior(userId, actionType, targetId, targetType, extraData)
    Dim sql, sessionIdStr, ipAddr
    
    On Error Resume Next
    
    ' 获取会话信息
    If IsEmpty(userId) Or userId = "" Then
        userId = 0
    ElseIf Not IsNumeric(userId) Then
        userId = 0
    End If
    
    sessionIdStr = Session.SessionID
    If sessionIdStr = "" Then sessionIdStr = "ANON"
    
    ipAddr = Request.ServerVariables("REMOTE_ADDR")
    If ipAddr = "" Then ipAddr = "0.0.0.0"
    
    ' 使用 SQL Server 的辅助表记录（利用 SiteSettings 作为存储）
    Dim logKey
    logKey = "UV_" & actionType & "_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Hour(Now) & Minute(Now) & Second(Now) & "_" & sessionIdStr
    
    Dim logVal
    logVal = "UID:" & userId & "|Action:" & actionType & "|Target:" & targetId & "|Type:" & targetType & "|IP:" & ipAddr & "|Extra:" & extraData
    
    sql = "INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('" & SafeSQL(logKey) & "', '" & SafeSQL(logVal) & "')"
    conn.Execute sql
End Sub

' ============================================
' 记录商品浏览
' ============================================
Sub TU_LogProductView(userId, productId)
    Call TU_LogBehavior(userId, "PRODUCT_VIEW", productId, "product", "")
End Sub

' ============================================
' 记录搜索
' ============================================
Sub TU_LogSearch(userId, keyword)
    Call TU_LogBehavior(userId, "SEARCH", keyword, "search", "")
End Sub

' ============================================
' 记录加购
' ============================================
Sub TU_LogCartAdd(userId, productId, quantity)
    Call TU_LogBehavior(userId, "CART_ADD", productId, "product", "Qty:" & quantity)
End Sub

' ============================================
' 记录收藏
' ============================================
Sub TU_LogFavorite(userId, productId)
    Call TU_LogBehavior(userId, "FAVORITE", productId, "product", "")
End Sub

' ============================================
' 获取今日PV/UV统计（管理员用）
' ============================================
Function TU_GetDailyStats()
    Dim rs, todayStr, pv, uv
    pv = 0
    uv = 0
    
    todayStr = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)
    
    On Error Resume Next
    
    ' PV: 当日所有行为记录数
    Set rs = conn.Execute("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey LIKE 'UV_%_" & todayStr & "_%'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then pv = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    ' UV: 当日不同会话数
    Set rs = conn.Execute("SELECT COUNT(DISTINCT RIGHT(SettingKey, CHARINDEX('_', REVERSE(SettingKey))-1)) FROM SiteSettings WHERE SettingKey LIKE 'UV_%_" & todayStr & "_%'")
    If Not rs Is Nothing Then
        If Not rs.EOF Then uv = CLng(rs(0))
        rs.Close
    End If
    Set rs = Nothing
    
    Dim result
    result = "{""pv"":" & pv & ",""uv"":" & uv & "}"
    TU_GetDailyStats = result
End Function

' ============================================
' 获取热门搜索关键词
' ============================================
Function TU_GetHotSearches(topN)
    Dim rs, sql, result
    Set result = Server.CreateObject("Scripting.Dictionary")
    
    If topN <= 0 Then topN = 10
    
    On Error Resume Next
    
    sql = "SELECT SettingValue FROM SiteSettings WHERE SettingKey LIKE 'UV_SEARCH_%'"
    Set rs = conn.Execute(sql)
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            Dim val, keyword
            val = rs(0)
            keyword = ""
            
            ' 从 SettingValue 中提取搜索词
            ' 格式: UID:x|Action:SEARCH|Target:keyword|...
            If InStr(val, "Target:") > 0 Then
                Dim startPos, endPos
                startPos = InStr(val, "Target:") + 7
                endPos = InStr(startPos, val, "|")
                If endPos > startPos Then
                    keyword = Mid(val, startPos, endPos - startPos)
                Else
                    keyword = Mid(val, startPos)
                End If
                
                If keyword <> "" Then
                    If result.Exists(keyword) Then
                        result(keyword) = result(keyword) + 1
                    Else
                        result.Add keyword, 1
                    End If
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
    
    Set TU_GetHotSearches = result
End Function

' ============================================
' 前端追踪脚本
' ============================================
Sub TU_RenderTrackingScript()
%>
<script>
(function() {
    // 页面浏览追踪
    var pageUrl = window.location.pathname;
    var pageTitle = document.title;
    
    // 产品浏览追踪 (product?id=XXX)
    var match = pageUrl.match(/\/product\.asp\?id=(\d+)/);
    if (match) {
        var img = new Image();
        img.src = '/api/track.asp?action=view&target=' + match[1] + '&t=' + new Date().getTime();
    }
    
    // 搜索追踪
    var urlParams = new URLSearchParams(window.location.search);
    var keyword = urlParams.get('keyword');
    if (keyword) {
        var img = new Image();
        img.src = '/api/track.asp?action=search&keyword=' + encodeURIComponent(keyword) + '&t=' + new Date().getTime();
    }
})();
</script>
<%
End Sub
%>