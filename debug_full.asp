<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
' Simulate authenticated session for testing
Session("UserID") = 25
Session("Username") = "raymond"
Response.Write "Session set: UserID=25<br>"

' Now include and run product.asp's core logic
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
Response.Write "Step1: connection OK<br>"
Response.Flush

Dim productId : productId = 1
Response.Write "Step2: productId=" & productId & "<br>"

' query product type
Dim rsType
Set rsType = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = 1 AND IsActive <> 0")
If rsType Is Nothing Or rsType.EOF Then
    Response.Write "ERROR: Product not found<br>"
    Response.End
End If
productType = LCase(rsType("ProductType") & "")
rsType.Close : Set rsType = Nothing
Response.Write "Step3: productType=" & productType & "<br>"
Response.Flush

' Test RE_GetSimilarFragrances
Response.Write "Step4: Testing RE_GetSimilarFragrances...<br>"
Response.Flush
Dim rsRelated
On Error Resume Next
Set rsRelated = RE_GetSimilarFragrances(1, 6)
If Err.Number <> 0 Then
    Response.Write "<b style='color:red'>ERROR: " & Err.Number & " - " & Err.Description & "</b><br>"
    Err.Clear
ElseIf rsRelated Is Nothing Then
    Response.Write "Step4b: rsRelated = Nothing<br>"
ElseIf rsRelated.EOF Then
    Response.Write "Step4b: rsRelated EOF<br>"
    rsRelated.Close : Set rsRelated = Nothing
Else
    Response.Write "Step4b: rsRelated OK, records=" & rsRelated.RecordCount & "<br>"
    rsRelated.Close : Set rsRelated = Nothing
End If
On Error GoTo 0

' Test RE_RenderRecommendations
Response.Write "Step5: Testing RE_RenderRecommendations...<br>"
Response.Flush
Set rsRelated = RE_GetSimilarFragrances(1, 4)
If Not rsRelated Is Nothing And Not rsRelated.EOF Then
    On Error Resume Next
    RE_RenderRecommendations rsRelated, "test-class", True
    If Err.Number <> 0 Then
        Response.Write "<b style='color:red'>ERROR in Render: " & Err.Number & " - " & Err.Description & "</b><br>"
        Err.Clear
    Else
        Response.Write "Step5b: Render OK<br>"
    End If
    On Error GoTo 0
    rsRelated.Close
End If
Set rsRelated = Nothing

' Test SU_RenderShareSection
Response.Write "Step6: Testing SU_RenderShareSection...<br>"
Response.Flush
On Error Resume Next
SU_RenderShareSection "http://localhost/product.asp?id=1", "Test", "Desc", ""
If Err.Number <> 0 Then
    Response.Write "<b style='color:red'>ERROR in Share: " & Err.Number & " - " & Err.Description & "</b><br>"
    Err.Clear
Else
    Response.Write "Step6b: Share OK<br>"
End If
On Error GoTo 0

' Test ProductReviews
Response.Write "Step7: FEATURE_COMMUNITY=" & FEATURE_COMMUNITY & "<br>"
If FEATURE_COMMUNITY Then
    Response.Write "Step7b: Testing DAL_GetRow on ProductReviews...<br>"
    Response.Flush
    On Error Resume Next
    Dim rsStats
    Set rsStats = DAL_GetRow("SELECT COUNT(*) AS TotalReviews, ISNULL(AVG(CAST(Rating AS DECIMAL(3,2))), 0) AS AvgRating FROM ProductReviews WHERE ProductID=@PID AND IsActive=1", _
        Array(Array("@PID", DAL_adInteger, 0, productId)))
    If Err.Number <> 0 Then
        Response.Write "<b style='color:red'>ERROR: " & Err.Number & " - " & Err.Description & "</b><br>"
        Err.Clear
    ElseIf rsStats Is Nothing Then
        Response.Write "Step7c: rsStats = Nothing<br>"
    Else
        Response.Write "Step7c: TotalReviews=" & rsStats("TotalReviews") & ", AvgRating=" & rsStats("AvgRating") & "<br>"
        rsStats.Close : Set rsStats = Nothing
    End If
    On Error GoTo 0

    ' Test review list
    Response.Write "Step8: Testing DAL_GetListPaged on ProductReviews...<br>"
    Response.Flush
    On Error Resume Next
    Dim reviewPageInfo, rsReviews
    Set rsReviews = DAL_GetListPaged("SELECT pr.*, u.Username, u.FullName FROM ProductReviews pr LEFT JOIN Users u ON pr.UserID=u.UserID WHERE pr.ProductID=@PID AND pr.IsActive=1 ORDER BY pr.CreatedAt DESC", _
        Array(Array("@PID", DAL_adInteger, 0, productId)), 1, 5, reviewPageInfo)
    If Err.Number <> 0 Then
        Response.Write "<b style='color:red'>ERROR: " & Err.Number & " - " & Err.Description & "</b><br>"
        Err.Clear
    ElseIf rsReviews Is Nothing Then
        Response.Write "Step8b: rsReviews = Nothing<br>"
    ElseIf rsReviews.EOF Then
        Response.Write "Step8b: rsReviews EOF (no reviews)<br>"
        rsReviews.Close : Set rsReviews = Nothing
    Else
        Response.Write "Step8b: Reviews found: " & rsReviews.RecordCount & "<br>"
        ' Test ReviewLikes query for each review
        Dim testLiked
        Do While Not rsReviews.EOF
            On Error Resume Next
            testLiked = CLng(DAL_GetScalar("SELECT COUNT(*) FROM ReviewLikes WHERE ReviewID=@RID AND UserID=@UID", _
                Array(Array("@RID", DAL_adInteger, 0, rsReviews("ReviewID")), Array("@UID", DAL_adInteger, 0, 25)), 0))
            If Err.Number <> 0 Then
                Response.Write "<b style='color:red'>ERROR in ReviewLikes query: " & Err.Number & " - " & Err.Description & "</b><br>"
                Err.Clear
            Else
                Response.Write "  Review " & rsReviews("ReviewID") & ": liked=" & testLiked & "<br>"
            End If
            On Error GoTo 0
            rsReviews.MoveNext
        Loop
        rsReviews.Close : Set rsReviews = Nothing
    End If
    On Error GoTo 0
End If

Response.Write "<br><b>ALL TESTS COMPLETED</b><br>"
Call CloseConnection()
%>
