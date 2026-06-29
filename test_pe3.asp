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
Response.Write "A|" & PE_GetRule("purchase_rate") & "<br>"
Response.Write "B|" & PE_GetRule("signin_points") & "<br>"
Response.Write "C|" & PE_GetRule("redeem_discount_rate") & "<br>"
Response.Write "D|" & PE_GetRule("max_redeem_pct") & "<br>"
Response.Write "ALL OK"
Call CloseConnection()
%>
