<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/dal.asp"-->
<!--#include file="../includes/dal_products.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
' ============================================
' V17.0 搜索建议 API (Search Suggestions)
' 支持: 关键词自动完成(suggest)、历史(history)、清空历史(clear_history)
' 用法: GET /api/search_suggestions.asp?q=关键词&type=suggest&max=8
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

Call OpenConnection()

Dim q, searchType, maxResults
q = Trim(Request.QueryString("q"))
searchType = Trim(Request.QueryString("type"))
If searchType = "" Then searchType = "suggest"
maxResults = CInt(Request.QueryString("max"))
If maxResults < 1 Or maxResults > 20 Then maxResults = 8

Select Case searchType
    Case "suggest"
        If q <> "" Then
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

Call CloseConnection()
%>
