<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Response.Write "X"
Call OpenConnection()
Response.Write "Y"
Response.Write "<br>OK<br>"
Call CloseConnection()
Response.Write "Z"
%>
