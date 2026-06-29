<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<%
Call OpenConnection()
Response.Write "1. RE_GetUserRecommendations(25,8)...<br>":Response.Flush
Dim dictR : Set dictR = RE_GetUserRecommendations(25, 8)
Response.Write "   Count: " & dictR.Count & "<br>":Response.Flush

Response.Write "2. RE_GetPopularProducts(8)...<br>":Response.Flush
Dim rsP : Set rsP = RE_GetPopularProducts(8)
If Not rsP Is Nothing Then Response.Write "   OK<br>" Else Response.Write "   Nothing<br>"
If Not rsP Is Nothing Then rsP.Close : Set rsP = Nothing
Response.Flush

Response.Write "3. RE_FetchProductsByIds with hardcoded IDs...<br>":Response.Flush
Dim testIds(2) : testIds(0)=1:testIds(1)=3:testIds(2)=6
Dim rsF : Set rsF = RE_FetchProductsByIds(testIds)
If Not rsF Is Nothing Then Response.Write "   OK<br>" Else Response.Write "   Nothing<br>"
If Not rsF Is Nothing Then rsF.Close : Set rsF = Nothing
Response.Flush

Call CloseConnection()
Response.Write "ALL DONE<br>"
%>