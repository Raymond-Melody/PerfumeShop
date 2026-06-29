<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/points_engine.asp"-->
<%
Call OpenConnection()
Response.Write "Cache OK<br>"

Response.Write "p=" & PE_GetRule("purchase_rate") & "<br>"
Response.Write "s=" & PE_GetRule("signin_points") & "<br>"
Response.Write "r=" & PE_GetRule("redeem_discount_rate") & "<br>"
Response.Write "DONE"
Call CloseConnection()
%>
