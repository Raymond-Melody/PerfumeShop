<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<% Call OpenConnection() %>
Step1: connection OK<br>
<%
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Write "ERROR: Not logged in"
    Response.End
End If
Response.Write "Step2: UserID=" & Session("UserID") & "<br>"
%>
<!--#include file="includes/dal.asp"-->
Step3: dal.asp included<br>
<%
Dim productId : productId = Request.QueryString("id")
If productId = "" Then productId = 1
Response.Write "Step4: productId=" & productId & "<br>"
%>
<!--#include file="includes/product_type_utils.asp"-->
Step5: product_type_utils.asp included<br>
<!--#include file="includes/ai_client.asp"-->
Step6: ai_client.asp included<br>
<!--#include file="includes/recommendation_engine.asp"-->
Step7: recommendation_engine.asp included<br>
<!--#include file="includes/share_utils.asp"-->
Step8: share_utils.asp included<br>
<% Response.Write "ALL INCLUDES OK<br>" %>
