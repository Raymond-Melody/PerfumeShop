<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/dal_admin.asp"-->
<!--#include file="includes/dal_users.asp"-->
<!--#include file="includes/dal_checkout.asp"-->
<!--#include file="includes/payment_handler.asp"-->
<!--#include file="includes/cost_engine.asp"-->
<!--#include file="includes/member_utils.asp"-->
<!--#include file="includes/points_engine.asp"-->
<%
Call OpenConnection()
Response.Write "=== Test 1: PE_GetRuleCache ===<br>"
Dim cache : Set cache = PE_GetRuleCache()
Response.Write "Cache Count: " & cache.Count & "<br>"
If cache.Count > 0 Then
    Dim k
    For Each k In cache.Keys()
        Response.Write "  " & k & " = " & cache(k) & " (TypeName: " & TypeName(cache(k)) & ")<br>"
    Next
End If

Response.Write "<br>=== Test 2: PE_GetRule ===<br>"
Response.Write "purchase_rate = " & PE_GetRule("purchase_rate") & "<br>"
Response.Write "signin_points = " & PE_GetRule("signin_points") & "<br>"
Response.Write "redeem_discount_rate = " & PE_GetRule("redeem_discount_rate") & "<br>"
Response.Write "max_redeem_pct = " & PE_GetRule("max_redeem_pct") & "<br>"
Response.Write "points_expire_months = " & PE_GetRule("points_expire_months") & "<br>"

Response.Write "<br>=== Test 3: PE_GetAvailablePoints ===<br>"
Dim pts : pts = PE_GetAvailablePoints(1)
Response.Write "User 1 points: " & pts & "<br>"

Call CloseConnection()
%>
