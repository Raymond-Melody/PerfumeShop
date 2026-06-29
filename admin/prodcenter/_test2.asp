<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Dim testVal : testVal = "index.asp"
Dim result : result = IIf(testVal = "index.asp", "active", "")
Response.Write "Result: " & result
%>
