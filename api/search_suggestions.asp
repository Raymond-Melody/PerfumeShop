<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V17.0 搜索建议 API (Search Suggestions)
' 支持: 关键词自动完成、热门搜索、搜索历史
' 用法: /api/search_suggestions.asp?q=关键词&type=suggest
' 返回: JSON格式的搜索结果建议列表
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/dal_products.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
Call OpenConnection()

Dim q, searchType, maxResults
q = Trim(Request.QueryString("q"))
searchType = Trim(Request.QueryString("type"))
If searchType = "" Then searchType = "suggest"
maxResults = CInt(Request.QueryString("max"))
If maxResults < 1 Or maxResults > 20 Then maxResults = 8

Select Case searchType
    Case "suggest"
        ' 关键词自动完成建议
        If q <> "" Then
            Dim rsSuggest
            Set rsSuggest = DAL_Products_GetSuggestions(q, maxResults)
            
            If Not rsSuggest Is Nothing Then
                Dim suggestions, itemArr, itemCount
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
                
                ' 记录搜索历史
                If q <> "" And Session("UserID") <> "" Then
                    Call DAL_Products_RecordSearch(Session("UserID"), q)
                End If
                
                Call API_Success(suggestions, "获取成功")
            Else
                Call API_Success("[]", "无匹配结果")
            End If
        Else
            Call API_Success("[]", "请输入搜索关键词")
        End If
        
    Case "history"
        ' 获取用户搜索历史
        If Session("UserID") <> "" Then
            Dim rsHistory
            Set rsHistory = DAL_Products_GetSearchHistory(Session("UserID"), maxResults)
            
            If Not rsHistory Is Nothing Then
                Dim historyItems
                historyItems = "["
                Dim hCount : hCount = 0
                
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
            Call API_Success("[]", "请先登录")
        End If
        
    Case "clear_history"
        ' 清空搜索历史
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

Call CloseConnection()
%>
<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V17.0 搜索建议 API
' 用法: GET /api/search_suggestions.asp?q=keyword&limit=5
' 返回JSON: {"code":0,"data":{"suggestions":["...","..."]}}
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

' 包含必要文件
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/dal_products.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
Call OpenConnection()

Dim keyword, limit, result, rs, suggestions, i, suggestion

keyword = Trim(Request.QueryString("q"))
limit = Request.QueryString("limit")
If limit = "" Or Not IsNumeric(limit) Then limit = 5
limit = CInt(limit)
If limit < 1 Then limit = 5
If limit > 20 Then limit = 20

If keyword = "" Then
    API_Response API_ERR_SUCCESS, "success", Array()
    Call CloseConnection()
    Response.End
End If

' 搜索产品名称（去重）
Dim sql, params(0)
sql = "SELECT DISTINCT TOP " & limit & " ProductName FROM Products " & _
      "WHERE IsActive=1 AND ProductName LIKE '%' + @Keyword + '%' " & _
      "ORDER BY ProductName ASC"
params(0) = Array("@Keyword", DAL_adVarChar, 100, keyword)
Set rs = DAL_GetList(sql, params)

If Not rs Is Nothing And IsObject(rs) Then
    suggestions = Array()
    i = 0
    Do While Not rs.EOF And i < limit
        suggestion = rs("ProductName") & ""
        If suggestion <> "" Then
            ReDim Preserve suggestions(i)
            suggestions(i) = suggestion
            i = i + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
Else
    suggestions = Array()
End If

API_Response API_ERR_SUCCESS, "success", suggestions

Call CloseConnection()
%>
