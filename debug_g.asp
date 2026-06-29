<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<%
Call OpenConnection()
Dim t1 : t1 = Timer()
Dim dictR : Set dictR = RE_GetUserRecommendations(25, 8)
Dim t2 : t2 = Timer()
Response.Write "RE_GetUserRecommendations: " & FormatNumber((t2-t1)*1000, 0) & "ms, count=" & dictR.Count
%>