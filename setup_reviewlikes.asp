<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/connection.asp"-->
<%
Call OpenConnection()
Response.Write "Creating ReviewLikes table...<br>"

Dim sql
sql = "IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ReviewLikes]') AND type in (N'U')) " & _
      "BEGIN CREATE TABLE [dbo].[ReviewLikes] (LikeID INT IDENTITY(1,1) PRIMARY KEY, ReviewID INT NOT NULL, " & _
      "UserID INT NOT NULL, CreatedAt DATETIME NOT NULL DEFAULT GETDATE(), " & _
      "FOREIGN KEY (ReviewID) REFERENCES [ProductReviews](ReviewID) ON DELETE CASCADE, " & _
      "FOREIGN KEY (UserID) REFERENCES [Users](UserID), CONSTRAINT UQ_ReviewLikes UNIQUE (ReviewID, UserID)); END"

On Error Resume Next
conn.Execute sql
If Err.Number = 0 Then
    Response.Write "ReviewLikes table created successfully!<br>"
Else
    Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br>"
    Err.Clear
End If
On Error GoTo 0

' Verify
Dim rs : Set rs = conn.Execute("SELECT TOP 1 1 FROM ReviewLikes")
If Not rs Is Nothing Then
    Response.Write "Verified: ReviewLikes table exists<br>"
    rs.Close
Else
    Response.Write "Failed to verify<br>"
End If
Set rs = Nothing

Call CloseConnection()
%>
