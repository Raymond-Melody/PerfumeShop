<%@ Language="VBScript" CodePage="65001" %>
<%
' 清除所有Session
Session.Abandon

' 跳转到首页
Response.Redirect "/index.asp"
%>
