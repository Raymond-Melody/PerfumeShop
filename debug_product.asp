<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<%
Call OpenConnection()

' Simulate logged-in state
Session("UserID") = 25

Response.Write "Step 1: product.asp imports OK<br>"

' Test RE_GetRelatedProducts directly
Response.Write "Step 2: RE_GetRelatedProducts(1, 'standard', 6)..."
On Error Resume Next
Dim rsRel : Set rsRel = RE_GetRelatedProducts(1, "standard", 6)
If Err.Number <> 0 Then
    Response.Write " ERROR: " & Err.Description & "<br>"
    Err.Clear
ElseIf rsRel Is Nothing Then
    Response.Write " Nothing<br>"
Else
    Response.Write " OK, EOF=" & rsRel.EOF & "<br>"
    rsRel.Close : Set rsRel = Nothing
End If
On Error GoTo 0

' Test RE_GetSimilarFragrances
Response.Write "Step 3: RE_GetSimilarFragrances(1, 6)..."
On Error Resume Next
Dim rsSim : Set rsSim = RE_GetSimilarFragrances(1, 6)
If Err.Number <> 0 Then
    Response.Write " ERROR: " & Err.Description & "<br>"
    Err.Clear
ElseIf rsSim Is Nothing Then
    Response.Write " Nothing<br>"
Else
    Response.Write " OK, EOF=" & rsSim.EOF & "<br>"
    rsSim.Close : Set rsSim = Nothing
End If
On Error GoTo 0

' Test product query (like product.asp does)
Response.Write "Step 4: Query product 1..."
On Error Resume Next
Dim rsProd : Set rsProd = ExecuteQuery("SELECT * FROM Products WHERE ProductID = 1 AND IsActive <> 0")
If Err.Number <> 0 Then
    Response.Write " ERROR: " & Err.Description & "<br>"
    Err.Clear
ElseIf rsProd Is Nothing Then
    Response.Write " Nothing<br>"
Else
    Response.Write " OK, Name=" & rsProd("ProductName") & "<br>"
    rsProd.Close : Set rsProd = Nothing
End If
On Error GoTo 0

Response.Write "Step 5: ALL DONE<br>"
Call CloseConnection()
%>
