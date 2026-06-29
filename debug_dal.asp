<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()
Response.Write "conn OK<br>"
%>
<!--#include file="includes/dal.asp"-->
<%
Response.Write "dal OK, FEATURE_COMMUNITY=" & FEATURE_COMMUNITY & "<br>"
' Test ProductReviews
Dim rsS : Set rsS = DAL_GetRow("SELECT COUNT(*) AS Cnt FROM ProductReviews", Null)
If Not rsS Is Nothing Then
    Response.Write "ProductReviews: " & rsS("Cnt") & " rows<br>"
    rsS.Close : Set rsS = Nothing
Else
    Response.Write "ProductReviews query returned Nothing<br>"
End If
' Test ReviewLikes
Set rsS = DAL_GetRow("SELECT COUNT(*) AS Cnt FROM ReviewLikes", Null)
If Not rsS Is Nothing Then
    Response.Write "ReviewLikes: " & rsS("Cnt") & " rows<br>"
    rsS.Close : Set rsS = Nothing
Else
    Response.Write "ReviewLikes query returned Nothing<br>"
End If
Response.Write "DONE<br>"
Call CloseConnection()
%>
