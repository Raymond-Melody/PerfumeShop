<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' 用户行为追踪API
' 前端通过图片请求（1x1 GIF）记录用户行为
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "image/gif"
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-cache, no-store"

' 输出1x1透明GIF
Dim gifData
gifData = ChrB(71) & ChrB(73) & ChrB(70) & ChrB(56) & ChrB(57) & ChrB(97) & _
          ChrB(1) & ChrB(0) & ChrB(1) & ChrB(0) & ChrB(128) & ChrB(0) & _
          ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(255) & ChrB(255) & _
          ChrB(255) & ChrB(33) & ChrB(249) & ChrB(4) & ChrB(0) & ChrB(0) & _
          ChrB(0) & ChrB(0) & ChrB(0) & ChrB(44) & ChrB(0) & ChrB(0) & _
          ChrB(0) & ChrB(0) & ChrB(1) & ChrB(0) & ChrB(1) & ChrB(0) & _
          ChrB(0) & ChrB(2) & ChrB(2) & ChrB(68) & ChrB(1) & ChrB(0) & ChrB(59)
Response.BinaryWrite gifData

' 记录行为（输出GIF后再记录，不阻塞响应）
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/tracking_utils.asp"-->
<%
Call OpenConnection()

Dim action, target, keyword
action = Request.QueryString("action")
target = Request.QueryString("target")
keyword = Request.QueryString("keyword")

Dim userId
userId = Session("UserID")
If Not IsNumeric(userId) Then userId = 0

Select Case action
    Case "view"
        If IsNumeric(target) Then
            Call TU_LogProductView(userId, CLng(target))
        End If
    Case "search"
        If keyword <> "" Then
            Call TU_LogSearch(userId, keyword)
        End If
    Case "cart"
        If IsNumeric(target) Then
            Dim qty
            qty = Request.QueryString("qty")
            If Not IsNumeric(qty) Then qty = 1
            Call TU_LogCartAdd(userId, CLng(target), qty)
        End If
    Case "fav"
        If IsNumeric(target) Then
            Call TU_LogFavorite(userId, CLng(target))
        End If
End Select

Call CloseConnection()
%>