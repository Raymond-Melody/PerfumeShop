<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Dim conn, connStr
connStr = "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connStr

' Create table
conn.Execute "IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PointsRules]') AND type in (N'U')) CREATE TABLE PointsRules (RuleID INT IDENTITY(1,1) PRIMARY KEY, RuleCode NVARCHAR(50) NOT NULL UNIQUE, RuleName NVARCHAR(100) NOT NULL, RuleValue DECIMAL(10,2) NOT NULL, RuleUnit NVARCHAR(20) NOT NULL DEFAULT '', IsEnabled BIT NOT NULL DEFAULT 1, SortOrder INT NOT NULL DEFAULT 0, Description NVARCHAR(200) DEFAULT '', CreatedAt DATETIME NOT NULL DEFAULT GETDATE(), UpdatedAt DATETIME NOT NULL DEFAULT GETDATE())"
Response.Write "Step1: Table OK<br>"

' Insert rules one at a time
Dim rules(9,2)
rules(0,0) = "purchase_rate" : rules(0,1) = "消费积分比例" : rules(0,2) = 1
rules(1,0) = "signin_points" : rules(1,1) = "签到积分" : rules(1,2) = 5
rules(2,0) = "review_points" : rules(2,1) = "评价积分" : rules(2,2) = 20
rules(3,0) = "review_with_photo" : rules(3,1) = "带图评价积分" : rules(3,2) = 10
rules(4,0) = "share_points" : rules(4,1) = "分享积分" : rules(4,2) = 10
rules(5,0) = "referral_points" : rules(5,1) = "推荐注册积分" : rules(5,2) = 100
rules(6,0) = "referral_purchase" : rules(6,1) = "推荐消费积分" : rules(6,2) = 50
rules(7,0) = "redeem_discount_rate" : rules(7,1) = "积分抵扣比例" : rules(7,2) = 100
rules(8,0) = "max_redeem_pct" : rules(8,1) = "最大抵扣比例" : rules(8,2) = 30
rules(9,0) = "points_expire_months" : rules(9,1) = "积分有效期" : rules(9,2) = 12

Dim i, code, name, val, sql, rs
For i = 0 To 9
    code = rules(i,0) : name = rules(i,1) : val = rules(i,2)
    Set rs = conn.Execute("SELECT COUNT(*) FROM PointsRules WHERE RuleCode='" & code & "'")
    If rs(0) = 0 Then
        sql = "INSERT INTO PointsRules (RuleCode,RuleName,RuleValue,RuleUnit,SortOrder) VALUES ('" & code & "',N'" & name & "'," & val & ",'points'," & (i+1) & ")"
        conn.Execute sql
    End If
    rs.Close : Set rs = Nothing
Next
Response.Write "Step2: Rules inserted<br>"

' Verify
Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM PointsRules")
Response.Write "Step3: Count=" & rs(0) & "<br>"
rs.Close : Set rs = Nothing

conn.Close : Set conn = Nothing
Response.Write "DONE"
%>
