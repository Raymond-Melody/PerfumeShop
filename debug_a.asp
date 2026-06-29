<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Response.Write "A"
Call OpenConnection()
Response.Write "B"
Call CloseConnection()
Response.Write "C"
%>