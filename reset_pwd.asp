<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
%>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<!--#include file="includes/dal.asp"-->
<!--#include file="includes/password_utils.asp"-->
<%
Call OpenConnection()

' Compute V3 hash using ASP's own function
Dim newPwd, newHash
newPwd = "raymond@2026"
newHash = HashPassword(newPwd)

' Update database
Dim pwdParams(1)
pwdParams(0) = Array("@Password", DAL_adVarChar, 255, newHash)
pwdParams(1) = Array("@UserID", DAL_adInteger, 0, 25)
DAL_Execute "UPDATE Users SET [Password] = @Password WHERE UserID = @UserID", pwdParams

Response.Write "Password reset!<br>"
Response.Write "New hash: " & Server.HTMLEncode(newHash) & "<br>"
Response.Write "Password: raymond@2026<br>"

' Verify
Dim rsV
Set rsV = DAL_GetList("SELECT [Password] FROM Users WHERE UserID=25", Array())
If Not rsV Is Nothing And Not rsV.EOF Then
    Response.Write "Stored: " & Server.HTMLEncode(rsV("Password")) & "<br>"
    Response.Write "Match: " & (rsV("Password") = newHash) & "<br>"
    rsV.Close
End If
Set rsV = Nothing

Call CloseConnection()
%>