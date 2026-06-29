<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/promotion_engine.asp"-->
<!--#include file="../includes/api_response.asp"-->
<%
Call OpenConnection()

' 登录检查
If Session("UserID") = "" Or IsNull(Session("UserID")) Then
    Call API_Error("请先登录")
End If

Dim userId, action, couponCode, cartTotal
userId = Session("UserID")
action = Request.QueryString("action")
If action = "" Then action = Request.Form("action")

If action = "validate" Or action = "" Then
    couponCode = Request.Form("code")
    If couponCode = "" Then couponCode = Request.QueryString("code")
    cartTotal = Request.Form("cart_total")
    If cartTotal = "" Then cartTotal = Request.QueryString("cart_total")
    
    If couponCode = "" Then
        Call API_Error("请输入优惠码")
    End If
    If Not IsNumeric(cartTotal) Or CDbl(cartTotal) < 0 Then
        cartTotal = 0
    End If
    
    Dim result
    Set result = PE_CouponValidate(couponCode, userId, CDbl(cartTotal))
    
    If result("valid") Then
        Dim resp
        Set resp = Server.CreateObject("Scripting.Dictionary")
        resp.Add "valid", True
        resp.Add "message", result("message")
        resp.Add "discount", result("discount")
        resp.Add "type", result("type")
        Call API_Success(resp)
    Else
        Call API_Error(result("message"))
    End If

ElseIf action = "list" Then
    ' 获取用户可用券列表
    Dim rsCoupons, jsonArr, rsItem
    Set jsonArr = Server.CreateObject("Scripting.Dictionary")
    Set rsCoupons = PE_CouponGetUserCoupons(userId, "available")
    
    Dim coupons, idx
    Set coupons = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    If Not rsCoupons Is Nothing Then
        Do While Not rsCoupons.EOF
            Dim item
            Set item = Server.CreateObject("Scripting.Dictionary")
            item.Add "userCouponId", rsCoupons("UserCouponID")
            item.Add "code", rsCoupons("CouponCode")
            item.Add "name", rsCoupons("CouponName")
            item.Add "type", rsCoupons("CouponType")
            item.Add "typeName", PE_CouponTypeName(rsCoupons("CouponType"))
            item.Add "value", CDbl(rsCoupons("DiscountValue"))
            item.Add "minSpend", CDbl(rsCoupons("MinSpend"))
            item.Add "maxDiscount", CDbl(rsCoupons("MaxDiscount"))
            item.Add "desc", PE_CouponFormatDesc(rsCoupons("CouponType"), CDbl(rsCoupons("DiscountValue")), CDbl(rsCoupons("MinSpend")), CDbl(rsCoupons("MaxDiscount")))
            item.Add "validTo", SafeFormatDateTime(rsCoupons("ValidTo"), 2)
            coupons.Add CStr(idx), item
            idx = idx + 1
            rsCoupons.MoveNext
        Loop
        rsCoupons.Close
    End If
    Set rsCoupons = Nothing
    
    Call API_Success(coupons)

ElseIf action = "claim" Then
    ' 领取公开优惠券
    couponCode = Request.Form("code")
    If couponCode = "" Then couponCode = Request.QueryString("code")
    If couponCode = "" Then
        Call API_Error("请输入优惠码")
    End If
    
    If PE_CouponIssue(userId, couponCode, "activity") Then
        Call API_Success("领取成功")
    Else
        Call API_Error("领取失败，可能优惠券不存在、已领完或您已有该券")
    End If

Else
    Call API_Error("无效的操作")
End If

Call CloseConnection()
%>
