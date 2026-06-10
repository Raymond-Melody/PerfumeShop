<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"

' 智能补货页面包装 - 转发到 replenishment.asp
Dim tabParam
tabParam = Trim(Request.QueryString("tab"))
If tabParam = "" Then tabParam = "RawMaterial"

' 重定向到实际的智能补货页面
Response.Redirect "replenishment.asp?tab=" & tabParam
%>
