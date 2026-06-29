<%
' ============================================
' V18.0 业务指标采集 (Metrics)
' 页面加载耗时、API 响应时间、购物车转化漏斗
' 依赖: config.asp, connection.asp, dal.asp
' 用法: <!--#include file="metrics.asp"-->
'        Call METRICS_RecordPageLoad("index", pageStartTime)
'        Call METRICS_RecordApiTime("cart_add", apiStartTime)
'        Call METRICS_TrackFunnel("view_product", userId, productId)
' ============================================

Const METRICS_ENABLED = True
Const METRICS_SLOW_PAGE_THRESHOLD = 2000    ' 慢页面阈值（ms）
Const METRICS_SLOW_API_THRESHOLD = 500      ' 慢API阈值（ms）

' ============================================
' METRICS_RecordPageLoad: 记录页面加载耗时
' 参数:
'   pageName   - 页面名称（如 "index", "product"）
'   startTime  - Timer() 获取的开始时间
'   userId     - 用户ID（可选）
' ============================================
Sub METRICS_RecordPageLoad(pageName, startTime, userId)
    If Not METRICS_ENABLED Then Exit Sub
    
    Dim elapsed, isSlow
    elapsed = Round((Timer() - startTime) * 1000, 1)
    isSlow = (elapsed > METRICS_SLOW_PAGE_THRESHOLD)
    
    ' 记录到 Application（供 health_check 读取）
    Application.Lock
    Application("METRICS_PageLoad_" & pageName) = elapsed
    Application("METRICS_PageCount") = CLng(Application("METRICS_PageCount")) + 1
    Application.UnLock
    
    ' 慢页面记录到 AppLogs
    If isSlow Then
        On Error Resume Next
        If IsNull(userId) Or userId = "" Then userId = 0
        Call DAL_Execute("INSERT INTO AppLogs (LogType, LogLevel, UserID, Message, IPAddress, CreatedAt) VALUES (@Type, @Level, @UID, @Msg, @IP, GETDATE())", _
            Array( _
                Array("@Type", DAL_adVarChar, 50, "PERF"), _
                Array("@Level", DAL_adVarChar, 20, "WARN"), _
                Array("@UID", DAL_adInteger, 0, CLng(userId)), _
                Array("@Msg", DAL_adVarChar, 500, "Slow page: " & pageName & " took " & elapsed & "ms"), _
                Array("@IP", DAL_adVarChar, 50, Request.ServerVariables("REMOTE_ADDR")) _
            ))
        On Error GoTo 0
    End If
End Sub

' ============================================
' METRICS_RecordApiTime: 记录 API 响应时间
' ============================================
Sub METRICS_RecordApiTime(apiName, startTime)
    If Not METRICS_ENABLED Then Exit Sub
    
    Dim elapsed, isSlow
    elapsed = Round((Timer() - startTime) * 1000, 1)
    isSlow = (elapsed > METRICS_SLOW_API_THRESHOLD)
    
    ' 记录到 Application
    Application.Lock
    Application("METRICS_ApiTime_" & apiName) = elapsed
    Application("METRICS_ApiCount") = CLng(Application("METRICS_ApiCount")) + 1
    Application.UnLock
    
    ' 慢 API 记录
    If isSlow Then
        On Error Resume Next
        Call DAL_Execute("INSERT INTO AppLogs (LogType, LogLevel, UserID, Message, IPAddress, CreatedAt) VALUES (@Type, @Level, @UID, @Msg, @IP, GETDATE())", _
            Array( _
                Array("@Type", DAL_adVarChar, 50, "PERF"), _
                Array("@Level", DAL_adVarChar, 20, "WARN"), _
                Array("@UID", DAL_adInteger, 0, 0), _
                Array("@Msg", DAL_adVarChar, 500, "Slow API: " & apiName & " took " & elapsed & "ms"), _
                Array("@IP", DAL_adVarChar, 50, Request.ServerVariables("REMOTE_ADDR")) _
            ))
        On Error GoTo 0
    End If
End Sub

' ============================================
' METRICS_TrackFunnel: 购物车转化漏斗埋点
' 步骤: view_product → add_to_cart → begin_checkout → purchase
' ============================================
Sub METRICS_TrackFunnel(step, userId, targetId)
    If Not METRICS_ENABLED Then Exit Sub
    If IsNull(userId) Or userId = "" Or userId = 0 Then Exit Sub
    
    Dim validSteps
    validSteps = ",view_product,add_to_cart,begin_checkout,purchase,"
    If InStr(validSteps, "," & step & ",") = 0 Then Exit Sub
    
    On Error Resume Next
    Call DAL_Execute("INSERT INTO UserBehavior (UserID, BehaviorType, TargetID, TargetType, CreatedAt) VALUES (@UID, @Type, @TID, @TType, GETDATE())", _
        Array( _
            Array("@UID", DAL_adInteger, 0, CLng(userId)), _
            Array("@Type", DAL_adVarChar, 50, "funnel_" & step), _
            Array("@TID", DAL_adInteger, 0, CLng(targetId)), _
            Array("@TType", DAL_adVarChar, 30, "product") _
        ))
    On Error GoTo 0
End Sub

' ============================================
' METRICS_GetFunnelStats: 获取转化漏斗统计
' 返回: Dictionary(step → count)
' ============================================
Function METRICS_GetFunnelStats()
    Dim result, steps, i, step, count
    Set result = Server.CreateObject("Scripting.Dictionary")
    steps = Array("view_product", "add_to_cart", "begin_checkout", "purchase")
    
    On Error Resume Next
    For i = 0 To UBound(steps)
        step = steps(i)
        count = CLng(DAL_GetScalar("SELECT COUNT(*) FROM UserBehavior WHERE BehaviorType='funnel_" & step & "' AND CreatedAt > DATEADD(DAY, -7, GETDATE())", Empty, 0))
        If Err.Number <> 0 Then
            count = 0
            Err.Clear
        End If
        result.Add step, count
    Next
    On Error GoTo 0
    
    Set METRICS_GetFunnelStats = result
End Function

' ============================================
' METRICS_GetSummary: 获取指标摘要（用于监控面板）
' ============================================
Function METRICS_GetSummary()
    Dim json
    json = "{"
    
    ' 页面指标
    Dim pageCount, apiCount
    pageCount = 0 : apiCount = 0
    On Error Resume Next
    pageCount = CLng(Application("METRICS_PageCount"))
    apiCount = CLng(Application("METRICS_ApiCount"))
    On Error GoTo 0
    
    json = json & """pageCount"":" & pageCount & ","
    json = json & """apiCount"":" & apiCount & ","
    
    ' 漏斗数据
    Dim funnel, step, count, isFirst
    Set funnel = METRICS_GetFunnelStats()
    json = json & """funnel"":{"
    isFirst = True
    Dim funnelKeys : funnelKeys = funnel.Keys
    Dim fi
    For fi = 0 To funnel.Count - 1
        step = funnelKeys(fi)
        count = funnel.Item(step)
        If Not isFirst Then json = json & ","
        json = json & """" & step & """:" & count
        isFirst = False
    Next
    json = json & "}"
    Set funnel = Nothing
    
    json = json & "}"
    METRICS_GetSummary = json
End Function
%>
