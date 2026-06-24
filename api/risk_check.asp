<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/connection.asp"-->
<%
' ============================================
' 风控校验API - Risk Check API
' 下单前风险评估：检查用户信用、金额异常、地址重复
' 返回JSON格式风险评级
' ============================================
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

' 安全数值函数
Function RC_SafeNum(val)
    If IsNull(val) Or IsEmpty(val) Or val = "" Then RC_SafeNum = 0 Else On Error Resume Next: RC_SafeNum = CDbl(val): If Err.Number <> 0 Then RC_SafeNum = 0: Err.Clear: End If
End Function

' 检查风控
Function CheckRisk(userId, orderTotal, productIds, shippingAddress, shippingPhone)
    Dim result, risks, riskCount, maxRisk
    Set result = Server.CreateObject("Scripting.Dictionary")
    Set risks = Server.CreateObject("Scripting.Dictionary")
    riskCount = 0
    maxRisk = "low"
    
    On Error Resume Next
    
    Call OpenConnection()
    
    ' 1. 用户信用检查
    If userId > 0 Then
        Dim rsUser, totalOrders, totalSpent, returnCount, unpaidOrders, cancelCount
        Set rsUser = conn.Execute("SELECT " & _
            "(SELECT COUNT(*) FROM Orders WHERE UserID=" & userId & " AND Status NOT IN ('Cancelled')) AS TotalOrders, " & _
            "(SELECT ISNULL(SUM(TotalAmount),0) FROM Orders WHERE UserID=" & userId & " AND Status IN ('Paid','Completed')) AS TotalSpent, " & _
            "(SELECT COUNT(*) FROM Orders WHERE UserID=" & userId & " AND Status='Returned') AS ReturnCount, " & _
            "(SELECT COUNT(*) FROM Orders WHERE UserID=" & userId & " AND DATEDIFF(day, CreatedAt, GETDATE()) > 7 AND Status='Pending') AS UnpaidOrders, " & _
            "(SELECT COUNT(*) FROM Orders WHERE UserID=" & userId & " AND Status='Cancelled') AS CancelCount")
        
        If Not rsUser Is Nothing And Not rsUser.EOF Then
            totalOrders = RC_SafeNum(rsUser("TotalOrders"))
            totalSpent = RC_SafeNum(rsUser("TotalSpent"))
            returnCount = RC_SafeNum(rsUser("ReturnCount"))
            unpaidOrders = RC_SafeNum(rsUser("UnpaidOrders"))
            cancelCount = RC_SafeNum(rsUser("CancelCount"))
            rsUser.Close
            
            ' D级用户判断
            Dim returnRate, cancelRate
            If totalOrders > 0 Then
                returnRate = returnCount / totalOrders
                cancelRate = cancelCount / totalOrders
            Else
                returnRate = 0
                cancelRate = 0
            End If
            
            If totalSpent < 100 And totalOrders > 0 And (returnRate > 0.3 Or cancelRate > 0.3) Then
                risks.Add "credit", "{""type"":""credit"",""level"":""high"",""msg"":""用户信用评级低，退货/取消率高""}"
                riskCount = riskCount + 1
                maxRisk = "high"
            ElseIf unpaidOrders > 2 Then
                risks.Add "credit", "{""type"":""credit"",""level"":""medium"",""msg"":""用户有 " & unpaidOrders & " 笔未支付订单""}"
                riskCount = riskCount + 1
                If maxRisk <> "high" Then maxRisk = "medium"
            End If
            
            ' 2. 金额异常检查
            If orderTotal > 0 And totalOrders > 0 Then
                Dim avgOrderAmount
                avgOrderAmount = totalSpent / totalOrders
                If orderTotal > avgOrderAmount * 5 And orderTotal > 2000 Then
                    risks.Add "amount", "{""type"":""amount"",""level"":""medium"",""msg"":""订单金额¥" & FormatNumber(orderTotal,2) & "远超历史平均¥" & FormatNumber(avgOrderAmount,2) & """}"
                    riskCount = riskCount + 1
                    If maxRisk <> "high" Then maxRisk = "medium"
                End If
            End If
        End If
        Set rsUser = Nothing
    End If
    
    ' 3. 地址/电话重复检查
    If shippingAddress <> "" Or shippingPhone <> "" Then
        Dim addrClause
        addrClause = ""
        If shippingAddress <> "" Then
            Dim safeAddr
            safeAddr = Replace(shippingAddress, "'", "''")
            addrClause = "o.ShippingAddress='" & safeAddr & "'"
        End If
        If shippingPhone <> "" Then
            Dim safePhone
            safePhone = Replace(shippingPhone, "'", "''")
            If addrClause <> "" Then addrClause = addrClause & " OR "
            addrClause = addrClause & "o.ShippingPhone='" & safePhone & "'"
        End If
        
        If addrClause <> "" Then
            Dim rsAddr
            Set rsAddr = conn.Execute("SELECT COUNT(*) AS Cnt FROM Orders o WHERE (" & addrClause & ") AND o.Status NOT IN ('Cancelled') AND DATEDIFF(day, o.CreatedAt, GETDATE()) <= 30")
            If Not rsAddr Is Nothing Then
                If Not rsAddr.EOF Then
                    Dim addrCount
                    addrCount = RC_SafeNum(rsAddr("Cnt"))
                    If addrCount >= 5 Then
                        risks.Add "address", "{""type"":""address"",""level"":""high"",""msg"":""该地址/电话近30天下单" & addrCount & "次，异常频繁""}"
                        riskCount = riskCount + 1
                        maxRisk = "high"
                    ElseIf addrCount >= 3 Then
                        risks.Add "address", "{""type"":""address"",""level"":""low"",""msg"":""该地址/电话近30天下单" & addrCount & "次""}"
                        riskCount = riskCount + 1
                    End If
                End If
                rsAddr.Close
            End If
            Set rsAddr = Nothing
        End If
    End If
    
    ' 4. IP频率检查（30分钟内同一IP下单次数）
    Dim ipAddr
    ipAddr = Request.ServerVariables("REMOTE_ADDR")
    If ipAddr <> "" Then
        Dim rsIP
        Set rsIP = conn.Execute("SELECT COUNT(*) AS Cnt FROM AdminLogs WHERE Notes='" & Replace(ipAddr, "'", "''") & "' AND CreatedAt >= DATEADD(minute, -30, GETDATE())")
        If Not rsIP Is Nothing Then
            If Not rsIP.EOF Then
                Dim ipCount
                ipCount = RC_SafeNum(rsIP("Cnt"))
                If ipCount > 10 Then
                    risks.Add "ip", "{""type"":""ip"",""level"":""medium"",""msg"":""IP " & ipAddr & " 30分钟内请求" & ipCount & "次""}"
                    riskCount = riskCount + 1
                    If maxRisk <> "high" Then maxRisk = "medium"
                End If
            End If
            rsIP.Close
        End If
        Set rsIP = Nothing
    End If
    
    Call CloseConnection()
    
    ' 构建返回JSON
    Dim jsonRisks, i
    jsonRisks = "["
    i = 0
    Dim keys, key
    keys = risks.Keys()
    For Each key In keys
        If i > 0 Then jsonRisks = jsonRisks & ","
        jsonRisks = jsonRisks & risks(key)
        i = i + 1
    Next
    jsonRisks = jsonRisks & "]"
    
    ' 构建最终结果
    result.Add "risk_level", maxRisk
    result.Add "risk_count", riskCount
    result.Add "risks", jsonRisks
    result.Add "passed", (maxRisk <> "high")
    result.Add "timestamp", Now()
    
    Set CheckRisk = result
End Function

' ============================================
' 处理请求
' ============================================
Dim userId, orderTotal, productIds, shippingAddress, shippingPhone
userId = CLng(Request("userId"))
orderTotal = RC_SafeNum(Request("orderTotal"))
productIds = Request("productIds")
shippingAddress = Request("shippingAddress")
shippingPhone = Request("shippingPhone")

Dim riskResult
Set riskResult = CheckRisk(userId, orderTotal, productIds, shippingAddress, shippingPhone)

' 输出JSON
Response.Write "{"
Response.Write """risk_level"":""" & riskResult("risk_level") & ""","
Response.Write """risk_count"":" & riskResult("risk_count") & ","
Response.Write """passed"":" & LCase(riskResult("passed")) & ","
Response.Write """timestamp"":""" & riskResult("timestamp") & ""","
Response.Write """risks"":" & riskResult("risks")
Response.Write "}"

Set riskResult = Nothing
%>