<%@ Language="VBScript" CodePage="65001" %>
<%
Session("UserID") = 25
Session("Username") = "raymond"
Response.Write "Session set<br>"
Response.Flush
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<!--#include file="includes/share_utils.asp"-->
<%
Call OpenConnection()
Response.Write "conn OK<br>"

Dim productId : productId = 1
' Get product type
Dim rsType : Set rsType = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = 1 AND IsActive <> 0")
If rsType Is Nothing Or rsType.EOF Then
    Response.Write "Product not found"
    Response.End
End If
Dim productType : productType = LCase(rsType("ProductType") & "")
rsType.Close : Set rsType = Nothing
Response.Write "productType=" & productType & "<br>"

' Test recommendation engine
Response.Write "Testing RE_GetSimilarFragrances...<br>"
Response.Flush
Dim t0 : t0 = Timer()
Dim rsRelated
Set rsRelated = RE_GetSimilarFragrances(1, 6)
If rsRelated Is Nothing Then
    Response.Write "rsRelated = Nothing<br>"
ElseIf rsRelated.EOF Then
    Response.Write "rsRelated is empty<br>"
    rsRelated.Close : Set rsRelated = Nothing
Else
    Response.Write "rsRelated: " & rsRelated.RecordCount & " records in " & FormatNumber((Timer()-t0)*1000,0) & "ms<br>"
    rsRelated.Close : Set rsRelated = Nothing
End If

' Test share utils
Response.Write "Testing SU_RenderShareSection...<br>"
Response.Flush
t0 = Timer()
SU_RenderShareSection "http://localhost/product.asp?id=1", "Test", "Desc", ""
Response.Write "Share OK: " & FormatNumber((Timer()-t0)*1000,0) & "ms<br>"

' Test FEATURE_COMMUNITY
Response.Write "FEATURE_COMMUNITY=" & FEATURE_COMMUNITY & "<br>"
Response.Flush
If FEATURE_COMMUNITY Then
    t0 = Timer()
    Dim rsStats
    Set rsStats = DAL_GetRow("SELECT COUNT(*) AS Cnt FROM ProductReviews WHERE ProductID=1 AND IsActive=1", _
        Array(Array("@PID", DAL_adInteger, 0, 1)))
    If Not rsStats Is Nothing Then
        Response.Write "ProductReviews: " & rsStats("Cnt") & " reviews in " & FormatNumber((Timer()-t0)*1000,0) & "ms<br>"
        rsStats.Close : Set rsStats = Nothing
    Else
        Response.Write "ProductReviews query returned Nothing<br>"
    End If
End If

Call CloseConnection()
Response.Write "ALL OK<br>"
%>
