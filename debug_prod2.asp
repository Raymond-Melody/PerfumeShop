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
Session("UserID") = 25
Response.Write "A<br>"

Dim productId : productId = 1

Dim rsType
Set rsType = ExecuteQuery("SELECT ProductType FROM Products WHERE ProductID = " & productId & " AND IsActive <> 0")
If rsType Is Nothing Or rsType.EOF Then
    Response.Write "Product not found<br>"
Else
    Dim productType : productType = LCase(rsType("ProductType") & "")
    rsType.Close
    Response.Write "productType: " & productType & "<br>"
End If
Set rsType = Nothing

Response.Write "B<br>"

Dim rsProduct
Set rsProduct = ExecuteQuery("SELECT * FROM Products WHERE ProductID = " & productId & " AND IsActive <> 0")
If Not rsProduct Is Nothing And Not rsProduct.EOF Then
    Response.Write "Product: " & rsProduct("ProductName") & "<br>"
    rsProduct.Close
End If
Set rsProduct = Nothing

Response.Write "C<br>"

' Test RE_GetSimilarFragrances
Dim rsRelated
Response.Write "Calling RE_GetSimilarFragrances...<br>"
Set rsRelated = RE_GetSimilarFragrances(productId, 6)
If rsRelated Is Nothing Then
    Response.Write "rsRelated is Nothing<br>"
ElseIf rsRelated.EOF Then
    Response.Write "rsRelated is Empty<br>"
    rsRelated.Close
Else
    Response.Write "rsRelated has " & rsRelated.RecordCount & " rows<br>"
    rsRelated.Close
End If
Set rsRelated = Nothing

Response.Write "D<br>"

' Test community feature
If FEATURE_COMMUNITY Then
    Response.Write "Community check...<br>"
    On Error Resume Next
    Dim cnt : cnt = DAL_GetScalar("SELECT COUNT(*) FROM ProductReviews WHERE ProductID=@PID AND UserID=@UID AND IsActive=1", _
        Array(Array("@PID", DAL_adInteger, 0, productId), Array("@UID", DAL_adInteger, 0, 25)), 0)
    If Err.Number <> 0 Then
        Response.Write "ERROR: " & Err.Description & "<br>"
        Err.Clear
    Else
        Response.Write "Reviews: " & cnt & "<br>"
    End If
    On Error GoTo 0
End If

Response.Write "E - DONE<br>"
Call CloseConnection()
%>
