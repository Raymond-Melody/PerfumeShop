<%
Response.Buffer = True
Dim connStr, conn
connStr = "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connStr

Dim sql
sql = "IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PointsRules]') AND type in (N'U')) " & _
      "BEGIN CREATE TABLE PointsRules (RuleID INT IDENTITY(1,1) PRIMARY KEY, RuleCode NVARCHAR(50) NOT NULL UNIQUE, " & _
      "RuleName NVARCHAR(100) NOT NULL, RuleValue DECIMAL(10,2) NOT NULL, RuleUnit NVARCHAR(20) NOT NULL DEFAULT '', " & _
      "IsEnabled BIT NOT NULL DEFAULT 1, SortOrder INT NOT NULL DEFAULT 0, Description NVARCHAR(200) DEFAULT '', " & _
      "CreatedAt DATETIME NOT NULL DEFAULT GETDATE(), UpdatedAt DATETIME NOT NULL DEFAULT GETDATE()); END"
conn.Execute sql
Response.Write "Table created or already exists<br>"

sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'purchase_rate','消费积分比例',1,'points',1,'每消费1元获得积分数' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='purchase_rate')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'signin_points','签到积分',5,'points',2,'每日签到奖励积分' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='signin_points')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'review_points','评价积分',20,'points',3,'发表有效评价奖励积分' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='review_points')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'review_with_photo','带图评价积分',10,'points',4,'带图评价额外奖励积分' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='review_with_photo')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'share_points','分享积分',10,'points',5,'分享产品/订单获得积分' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='share_points')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'referral_points','推荐注册积分',100,'points',6,'推荐好友成功注册获得积分' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='referral_points')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'referral_purchase','推荐消费积分',50,'points',7,'推荐好友首单获得积分' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='referral_purchase')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'redeem_discount_rate','积分抵扣比例',100,'rate',10,'每100积分可抵扣1元' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='redeem_discount_rate')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'max_redeem_pct','最大抵扣比例',30,'pct',11,'积分抵扣最高占订单金额百分比' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='max_redeem_pct')"
conn.Execute sql
sql = "INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) " & _
      "SELECT 'points_expire_months','积分有效期(月)',12,'days',12,'积分获取后有效期(月)' WHERE NOT EXISTS (SELECT 1 FROM PointsRules WHERE RuleCode='points_expire_months')"
conn.Execute sql
Response.Write "Default rules inserted<br>"

Dim rs : Set rs = conn.Execute("SELECT COUNT(*) AS Cnt FROM PointsRules")
Response.Write "PointsRules count: " & rs("Cnt") & "<br>"
rs.Close : Set rs = Nothing

conn.Close : Set conn = Nothing
Response.Write "DONE"
%>
