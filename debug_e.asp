<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<%
Response.Write "A"
Call OpenConnection()
Response.Write "B"

Response.Write " Calling RE_GetPersonalizedProducts(25,8)..."
Dim rsR : Set rsR = RE_GetPersonalizedProducts(25, 8)
Response.Write " Done."

If Not rsR Is Nothing Then rsR.Close : Set rsR = Nothing
Call CloseConnection()
Response.Write "C"
%>