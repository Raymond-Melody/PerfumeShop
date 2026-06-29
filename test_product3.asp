<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<!--#include file="includes/share_utils.asp"-->
<%
Session("UserID") = 25
Session("Username") = "raymond"
Call OpenConnection()

Dim tStart : tStart = Timer()
Dim productId : productId = 1

Dim t1 : t1 = Timer()
Dim rsType : Set rsType = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = 1 AND IsActive <> 0")
Dim productType : productType = ""
If Not rsType Is Nothing And Not rsType.EOF Then
    productType = LCase(rsType("ProductType") & "")
    rsType.Close
End If
Set rsType = Nothing
Response.Write "1.productType=" & productType & " (" & FormatNumber((Timer()-t1)*1000,0) & "ms)<br>"

Dim t2 : t2 = Timer()
Dim rsRelated : Set rsRelated = RE_GetSimilarFragrances(1, 6)
If Not rsRelated Is Nothing Then
    Response.Write "2.Related: " & rsRelated.RecordCount & " records (" & FormatNumber((Timer()-t2)*1000,0) & "ms)<br>"
Else
    Response.Write "2.Related: Nothing (" & FormatNumber((Timer()-t2)*1000,0) & "ms)<br>"
End If

Dim t3 : t3 = Timer()
If FEATURE_COMMUNITY Then
    Dim rsStats : Set rsStats = DAL_GetRow("SELECT COUNT(*) AS Cnt, ISNULL(AVG(CAST(Rating AS DECIMAL(3,2))),0) AS AvgR FROM ProductReviews WHERE ProductID=1 AND IsActive=1", Array(Array("@PID", DAL_adInteger, 0, 1)))
    If Not rsStats Is Nothing Then
        Response.Write "3.Reviews: " & rsStats("Cnt") & " reviews, avg=" & FormatNumber(rsStats("AvgR"),1) & " (" & FormatNumber((Timer()-t3)*1000,0) & "ms)<br>"
        rsStats.Close : Set rsStats = Nothing
    Else
        Response.Write "3.Reviews: query failed (" & FormatNumber((Timer()-t3)*1000,0) & "ms)<br>"
    End If
End If

If Not rsRelated Is Nothing Then
    If rsRelated.State = 1 Then rsRelated.Close
    Set rsRelated = Nothing
End If

Dim t4 : t4 = Timer()
SU_RenderShareSection "http://localhost/product.asp?id=1", "Test", "Desc", ""
Response.Write "4.Share OK (" & FormatNumber((Timer()-t4)*1000,0) & "ms)<br>"

Call CloseConnection()
Response.Write "TOTAL: " & FormatNumber((Timer()-tStart)*1000,0) & "ms<br>"
%>
