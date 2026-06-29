<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/product_type_utils.asp"-->
<!--#include file="includes/ai_client.asp"-->
<!--#include file="includes/recommendation_engine.asp"-->
<%
Call OpenConnection()

Dim t1, t2, t3, t4

' Simulate what index.asp does for authenticated user
Response.Write "<h2>Debug Performance Profile</h2>"

' Step 1: GetActiveProductTypes
t1 = Timer()
Dim activeTypes : activeTypes = GetActiveProductTypes()
t1 = FormatNumber((Timer() - t1) * 1000, 0)
Response.Write "GetActiveProductTypes: " & t1 & "ms<br>"

' Step 2: RE_GetPersonalizedProducts for user 25
t2 = Timer()
Dim rsRec
Set rsRec = RE_GetPersonalizedProducts(25, 8)
t2 = FormatNumber((Timer() - t2) * 1000, 0)
Response.Write "RE_GetPersonalizedProducts(25, 8): " & t2 & "ms<br>"
If Not rsRec Is Nothing Then
    Dim cnt : cnt = 0
    Do While Not rsRec.EOF
        cnt = cnt + 1
        rsRec.MoveNext
    Loop
    If rsRec.EOF Then rsRec.MoveFirst
    Response.Write "  Rows: " & cnt & "<br>"
    rsRec.Close : Set rsRec = Nothing
Else
    Response.Write "  Result: Nothing<br>"
End If

' Step 3: RE_GetUserRecommendations directly
t3 = Timer()
Dim dictRec : Set dictRec = RE_GetUserRecommendations(25, 8)
t3 = FormatNumber((Timer() - t3) * 1000, 0)
Response.Write "RE_GetUserRecommendations(25, 8): " & t3 & "ms<br>"
Response.Write "  Dict count: " & dictRec.Count & "<br>"

' Step 4: Product listing query
t4 = Timer()
Dim rsProd : Set rsProd = ExecuteQuery("SELECT TOP 8 * FROM Products WHERE IsActive <> 0 ORDER BY CreatedAt DESC")
t4 = FormatNumber((Timer() - t4) * 1000, 0)
Response.Write "Product listing query: " & t4 & "ms<br>"
If Not rsProd Is Nothing Then
    Dim cnt2 : cnt2 = 0
    Do While Not rsProd.EOF
        cnt2 = cnt2 + 1
        rsProd.MoveNext
    Loop
    Response.Write "  Rows: " & cnt2 & "<br>"
    rsProd.Close : Set rsProd = Nothing
End If

Call CloseConnection()
Response.Write "<h3>Done!</h3>"
%>