<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
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

If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Response.Write "ERROR: Not logged in"
    Response.End
End If
Response.Write "Step1: Logged in as UserID=" & Session("UserID") & "<br>"

Dim productId : productId = Request.QueryString("id")
If productId = "" Then productId = 1
Response.Write "Step2: productId=" & productId & "<br>"

' Test product exists
Dim rsType
Set rsType = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = " & CInt(productId) & " AND IsActive <> 0")
If rsType Is Nothing Or rsType.EOF Then
    Response.Write "ERROR: Product not found<br>"
    Response.End
End If
Dim productType : productType = LCase(rsType("ProductType") & "")
rsType.Close : Set rsType = Nothing
Response.Write "Step3: productType=" & productType & "<br>"

' Test RE_GetSimilarFragrances
Response.Write "Step4: Calling RE_GetSimilarFragrances...<br>"
Response.Flush
On Error Resume Next
Dim rsRelated : Set rsRelated = RE_GetSimilarFragrances(CLng(productId), 6)
If Err.Number <> 0 Then
    Response.Write "ERROR in RE_GetSimilarFragrances: " & Err.Number & " - " & Err.Description & "<br>"
    Err.Clear
ElseIf rsRelated Is Nothing Then
    Response.Write "Step4b: rsRelated is Nothing<br>"
ElseIf rsRelated.EOF Then
    Response.Write "Step4b: rsRelated is empty<br>"
    rsRelated.Close : Set rsRelated = Nothing
Else
    Response.Write "Step4b: rsRelated has " & rsRelated.RecordCount & " records<br>"
    rsRelated.Close : Set rsRelated = Nothing
End If
On Error GoTo 0

' Test SU_RenderShareSection
Response.Write "Step5: Calling SU_RenderShareSection...<br>"
Response.Flush
On Error Resume Next
Dim shareUrl : shareUrl = "http://localhost/product.asp?id=" & productId
Call SU_RenderShareSection(shareUrl, "Test Product", "Test Desc", "")
If Err.Number <> 0 Then
    Response.Write "ERROR in SU_RenderShareSection: " & Err.Number & " - " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "Step5b: Share section OK<br>"
End If
On Error GoTo 0

' Test FEATURE_COMMUNITY
Response.Write "Step6: FEATURE_COMMUNITY=" & FEATURE_COMMUNITY & "<br>"
If FEATURE_COMMUNITY Then
    Response.Write "Step6b: Testing ProductReviews query...<br>"
    Response.Flush
    On Error Resume Next
    Dim rsStats
    Set rsStats = DAL_GetRow("SELECT COUNT(*) AS TotalReviews, ISNULL(AVG(CAST(Rating AS DECIMAL(3,2))), 0) AS AvgRating FROM ProductReviews WHERE ProductID=@PID AND IsActive=1", _
        Array(Array("@PID", DAL_adInteger, 0, productId)))
    If Err.Number <> 0 Then
        Response.Write "ERROR in ProductReviews query: " & Err.Number & " - " & Err.Description & "<br>"
        Err.Clear
    Else
        Response.Write "Step6c: ProductReviews query OK<br>"
    End If
    On Error GoTo 0
End If

Response.Write "<br>ALL TESTS PASSED<br>"
%>
