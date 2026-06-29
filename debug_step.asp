<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = False
Response.Write "Start<br>"
Session("TestID") = 25
Response.Write "Session set<br>"
%>
<!--#include file="includes/config.asp"-->
<%
Response.Write "Config included, FEATURE_COMMUNITY=" & FEATURE_COMMUNITY & "<br>"
%>
<!--#include file="includes/connection.asp"-->
<%
Response.Write "Connection included<br>"
Call OpenConnection()
Response.Write "Connection opened<br>"
%>
<!--#include file="includes/dal.asp"-->
<%
Response.Write "DAL included<br>"
%>
<!--#include file="includes/recommendation_engine.asp"-->
<%
Response.Write "Recommendation engine included<br>"
%>
<!--#include file="includes/share_utils.asp"-->
<%
Response.Write "Share utils included<br>"
%>
Done
