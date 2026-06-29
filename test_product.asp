<%@ Language="VBScript" CodePage="65001" %>
<%
Session("UserID") = 25
Session("Username") = "raymond"
Response.Redirect "product.asp?id=1"
%>
