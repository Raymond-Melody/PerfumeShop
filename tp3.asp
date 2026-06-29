<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<%
Call OpenConnection()
Response.Write "A"
%>
<!--#include file="includes/recommendation_engine.asp"-->
<%
Response.Write "B"
Dim rs : Set rs = RE_GetSimilarFragrances(1, 2)
If Not rs Is Nothing Then
    Response.Write "C"
    rs.Close : Set rs = Nothing
End If
Response.Write "D"
%>
<!--#include file="includes/share_utils.asp"-->
<%
SU_RenderShareSection "http://localhost/test", "T", "D", ""
Response.Write "E"
Call CloseConnection()
Response.Write "F"
%>
