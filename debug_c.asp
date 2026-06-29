<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<%
Response.Write "A"
Call OpenConnection()
Response.Write "B"
Dim at : at = GetActiveProductTypes()
Response.Write "C"
Call CloseConnection()
Response.Write "D"
%>