<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' 收藏功能API
' 处理添加、删除和检查收藏状态
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<%
On Error Resume Next

' 安全的查询函数，不输出错误信息到响应流
Function SafeExecuteQuery(sql)
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open sql, conn, 1, 1
    If Err.Number <> 0 Then
        ' 不输出错误到响应，只返回Nothing
        Set SafeExecuteQuery = Nothing
        ' 为了调试，将错误保存到Session
        Session("LastDBError") = "查询错误: " & Err.Description & " SQL: " & sql
    Else
        Set SafeExecuteQuery = rs
    End If
    On Error GoTo 0
End Function

' 安全的非查询函数，不输出错误信息到响应流
Function SafeExecuteNonQuery(sql)
    On Error Resume Next
    conn.Execute sql
    If Err.Number <> 0 Then
        SafeExecuteNonQuery = False
    Else
        SafeExecuteNonQuery = True
    End If
    On Error GoTo 0
End Function

' 检查用户是否登录
If Session("UserID") = "" Or IsEmpty(Session("UserID")) Then
    Response.Write "{""success"": false, ""message"": ""请先登录""}"
    Response.End
End If

' 打开数据库连接
Call OpenConnection()

' CSRF验证 - 对于修改操作（add/remove）需要验证
Dim needsCSRF
needsCSRF = False
If Request.Form("action") = "add" Or Request.Form("action") = "remove" Then
    needsCSRF = True
End If

If needsCSRF And Not ValidateCSRFToken() Then
    Response.Write "{""success"": false, ""message"": ""安全验证失败，请刷新页面重试""}"
    Call CloseConnection()
    Response.End
End If

' 获取参数
Dim action, productId, userId
action = Request.Form("action")
If action = "" Then action = Request.QueryString("action")
productId = Request.Form("productId")
If productId = "" Then productId = Request.QueryString("productId")
userId = Session("UserID")

' 验证参数
If action = "" Then
    Response.Write "{""success"": false, ""message"": ""缺少操作类型参数""}"
    Call CloseConnection()
    Response.End
End If

If Not IsNumeric(productId) Then
    Response.Write "{""success"": false, ""message"": ""无效的产品ID""}"
    Call CloseConnection()
    Response.End
End If

' 处理不同的操作
Select Case action
    Case "add"
        ' 检查是否已收藏
        Dim rsCheck
        Set rsCheck = SafeExecuteQuery("SELECT FavoriteID FROM UserFavorites WHERE UserID = " & userId & " AND ProductID = " & productId)
        
        If Not rsCheck Is Nothing Then
            If Not rsCheck.EOF Then
                ' 已收藏
                Response.Write "{""success"": true, ""message"": ""已收藏"", ""action"": ""added""}"
                rsCheck.Close
                Set rsCheck = Nothing
            Else
                ' 添加收藏
                Dim sql
                sql = "INSERT INTO UserFavorites (UserID, ProductID, CreatedTime) VALUES (" & userId & ", " & productId & ", GETDATE())"
                If SafeExecuteNonQuery(sql) Then
                    Response.Write "{""success"": true, ""message"": ""收藏成功"", ""action"": ""added""}"
                Else
                    Dim insertError
                    insertError = ""
                    If Session("LastDBError") <> "" Then
                        insertError = " (" & Session("LastDBError") & ")"
                    End If
                    Response.Write "{""success"": false, ""message"": ""收藏失败" & insertError & """}"
                End If
                rsCheck.Close
                Set rsCheck = Nothing
            End If
        Else
            Dim queryError
            queryError = ""
            If Session("LastDBError") <> "" Then
                queryError = " (" & Session("LastDBError") & ")"
            End If
            Response.Write "{""success"": false, ""message"": ""数据库查询失败" & queryError & """}"
        End If
        
    Case "remove"
        ' 删除收藏
        Dim sqlDel
        sqlDel = "DELETE FROM UserFavorites WHERE UserID = " & userId & " AND ProductID = " & productId
        If SafeExecuteNonQuery(sqlDel) Then
            Response.Write "{""success"": true, ""message"": ""取消收藏成功"", ""action"": ""removed""}"
        Else
            Dim deleteError
            deleteError = ""
            If Session("LastDBError") <> "" Then
                deleteError = " (" & Session("LastDBError") & ")"
            End If
            Response.Write "{""success"": false, ""message"": ""取消收藏失败" & deleteError & """}"
        End If
        
    Case "check"
        ' 检查是否已收藏
        Dim rsCheck2
        Set rsCheck2 = SafeExecuteQuery("SELECT FavoriteID FROM UserFavorites WHERE UserID = " & userId & " AND ProductID = " & productId)
        
        If Not rsCheck2 Is Nothing Then
            If Not rsCheck2.EOF Then
                Response.Write "{""success"": true, ""isFavorite"": true}"
            Else
                Response.Write "{""success"": true, ""isFavorite"": false}"
            End If
            rsCheck2.Close
            Set rsCheck2 = Nothing
        Else
            Dim checkError
            checkError = ""
            If Session("LastDBError") <> "" Then
                checkError = " (" & Session("LastDBError") & ")"
            End If
            Response.Write "{""success"": false, ""message"": ""数据库查询失败" & checkError & """}"
        End If
        
    Case Else
        Response.Write "{""success"": false, ""message"": ""无效的操作""}"
End Select

Call CloseConnection()
On Error GoTo 0
%>