<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V18 生产工单同步API
' 修复: 已付款订单未自动生成生产工单
' 用法: POST /api/sync_production_orders.asp
' 返回: JSON {success, synced, errors}
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
Call OpenConnection()

Dim result, syncedCount, errorCount, errorMessages
syncedCount = 0
errorCount = 0
errorMessages = ""

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

' 获取所有需要同步的订单（已支付/处理中，但没有生产工单）
Dim sqlOrders, rsOrders
sqlOrders = "SELECT o.OrderID, o.OrderNo FROM Orders o " & _
            "WHERE o.Status IN ('Paid','Processing') " & _
            "AND o.OrderID NOT IN (SELECT DISTINCT OrderID FROM ProductionOrders WHERE Status <> 'Cancelled') " & _
            "ORDER BY o.OrderID"

On Error Resume Next
Set rsOrders = conn.Execute(sqlOrders)
If Err.Number <> 0 Then
    Response.Write "{""success"":false,""error"":""查询订单失败: " & Replace(Err.Description, """", "\""") & """}"
    Response.End
End If

If rsOrders Is Nothing Or rsOrders.EOF Then
    Response.Write "{""success"":true,""synced"":0,""message"":""所有订单都已同步，无需操作""}"
    Response.End
End If

' 遍历每个订单，为每个明细创建工单
Do While Not rsOrders.EOF
    Dim orderId, orderNo
    orderId = CLng(rsOrders("OrderID"))
    orderNo = rsOrders("OrderNo") & ""
    
    ' 获取订单明细
    Dim sqlDetails, rsDetails
    sqlDetails = "SELECT od.DetailID, od.Quantity, od.ProductID, p.ProductName, p.RecipeID, " & _
                 "r.RecipeName, r.RecipeCode " & _
                 "FROM OrderDetails od " & _
                 "LEFT JOIN Products p ON od.ProductID = p.ProductID " & _
                 "LEFT JOIN Recipes r ON p.RecipeID = r.RecipeID " & _
                 "WHERE od.OrderID = " & orderId
    
    Set rsDetails = conn.Execute(sqlDetails)
    
    If Not rsDetails Is Nothing Then
        Dim totalBottles : totalBottles = 0
        Dim bottleIndex : bottleIndex = 0
        
        ' 先生成工单前缀
        Dim workOrderPrefix
        workOrderPrefix = "WO-" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "-"
        
        Do While Not rsDetails.EOF
            Dim detailId, qty, productId, productName
            Dim recipeId, recipeName, recipeCode, fullRecipeName
            Dim hasRecipe
            
            detailId = CLng(rsDetails("DetailID"))
            qty = CLng(rsDetails("Quantity"))
            productId = CLng(rsDetails("ProductID"))
            productName = rsDetails("ProductName") & ""
            hasRecipe = False
            recipeId = 0
            recipeName = ""
            recipeCode = ""
            fullRecipeName = productName
            
            If Not IsNull(rsDetails("RecipeID")) Then
                hasRecipe = True
                recipeId = CLng(rsDetails("RecipeID"))
            End If
            If Not IsNull(rsDetails("RecipeName")) Then
                recipeName = rsDetails("RecipeName") & ""
                fullRecipeName = recipeName
            End If
            If Not IsNull(rsDetails("RecipeCode")) Then
                recipeCode = rsDetails("RecipeCode") & ""
                If recipeCode <> "" Then
                    fullRecipeName = "[" & recipeCode & "] " & recipeName
                End If
            End If
            
            ' 为每瓶创建工单
            Dim i
            For i = 1 To qty
                bottleIndex = bottleIndex + 1
                totalBottles = totalBottles + 1
                
                Dim workOrderNo
                workOrderNo = workOrderPrefix & Right("000" & bottleIndex, 4)
                
                ' 构建INSERT语句
                If hasRecipe And recipeId > 0 Then
                    conn.Execute "INSERT INTO ProductionOrders (OrderID, DetailID, WorkOrderNo, BottleIndex, TotalBottles, Status, Priority, RecipeID, RecipeName, CreatedAt, UpdatedAt) VALUES (" & _
                        orderId & ", " & detailId & ", '" & SafeSQL(workOrderNo) & "', " & bottleIndex & ", " & qty & ", 'Pending', 0, " & recipeId & ", '" & SafeSQL(fullRecipeName) & "', GETDATE(), GETDATE())"
                Else
                    conn.Execute "INSERT INTO ProductionOrders (OrderID, DetailID, WorkOrderNo, BottleIndex, TotalBottles, Status, Priority, CreatedAt, UpdatedAt) VALUES (" & _
                        orderId & ", " & detailId & ", '" & SafeSQL(workOrderNo) & "', " & bottleIndex & ", " & qty & ", 'Pending', 0, GETDATE(), GETDATE())"
                End If
                
                If Err.Number <> 0 Then
                    errorCount = errorCount + 1
                    If errorMessages <> "" Then errorMessages = errorMessages & "; "
                    errorMessages = errorMessages & "订单" & orderNo & " 工单" & workOrderNo & " 创建失败: " & Err.Description
                    Err.Clear
                Else
                    ' 获取新创建的ProductionID并记录日志
                    Dim newProdId, rsGetId
                    Set rsGetId = conn.Execute("SELECT MAX(ProductionID) AS NewId FROM ProductionOrders WHERE OrderID = " & orderId)
                    If Not rsGetId Is Nothing Then
                        If Not rsGetId.EOF Then
                            newProdId = CLng(rsGetId("NewId"))
                            If newProdId > 0 Then
                                conn.Execute "INSERT INTO ProductionLogs (ProductionID, Status, Notes, CreatedBy, CreatedAt) VALUES (" & _
                                    newProdId & ", 'Pending', '系统批量同步创建生产工单 (订单" & orderNo & " 第" & bottleIndex & "瓶/共" & qty & "瓶)', 'SYSTEM_SYNC', GETDATE())"
                                Err.Clear
                            End If
                        End If
                        rsGetId.Close
                    End If
                    Set rsGetId = Nothing
                    syncedCount = syncedCount + 1
                End If
            Next
            
            rsDetails.MoveNext
        Loop
        
        ' 更新订单状态为Processing
        If syncedCount > 0 Then
            conn.Execute "UPDATE Orders SET Status = 'Processing', UpdatedAt = GETDATE() WHERE OrderID = " & orderId & " AND Status = 'Paid'"
            Err.Clear
        End If
        
        rsDetails.Close
    End If
    Set rsDetails = Nothing
    
    rsOrders.MoveNext
Loop

rsOrders.Close
Set rsOrders = Nothing

On Error GoTo 0

' 返回结果
Dim jsonResponse
jsonResponse = "{""success"":true,""synced"":" & syncedCount & ",""errors"":" & errorCount
If errorMessages <> "" Then
    jsonResponse = jsonResponse & ",""errorMessages"":""" & Replace(errorMessages, """", "\""") & """"
End If
jsonResponse = jsonResponse & ",""message"":""成功同步 " & syncedCount & " 个生产工单"
If errorCount > 0 Then
    jsonResponse = jsonResponse & "，" & errorCount & " 个失败"
End If
jsonResponse = jsonResponse & """}"

Response.Write jsonResponse
%>
