<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.Buffer = True
Response.Charset = "UTF-8"
Dim conn, connStr 
connStr = "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;" 
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open connStr

' ==== 1. PointsLedger ====
Dim rsCheck1
Set rsCheck1 = conn.Execute("SELECT COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID(N'PointsLedger') AND type='U'")
If rsCheck1(0) = 0 Then
    conn.Execute "CREATE TABLE PointsLedger (LedgerID INT IDENTITY(1,1) PRIMARY KEY, UserID INT NOT NULL, Points INT NOT NULL, PointType NVARCHAR(20) NOT NULL, Source NVARCHAR(30) NOT NULL DEFAULT '', ReferenceID INT NULL, Description NVARCHAR(300) DEFAULT '', ExpiresAt DATETIME NULL, IsExpired BIT NOT NULL DEFAULT 0, CreatedAt DATETIME NOT NULL DEFAULT GETDATE())"
    Response.Write "PointsLedger: CREATED<br>"
    On Error Resume Next
    conn.Execute "CREATE INDEX IX_PointsLedger_UserID_CreatedAt ON PointsLedger(UserID, CreatedAt DESC)"
    conn.Execute "CREATE INDEX IX_PointsLedger_ExpiresAt ON PointsLedger(ExpiresAt) WHERE ExpiresAt IS NOT NULL AND IsExpired = 0"
    On Error GoTo 0
Else
    Response.Write "PointsLedger: EXISTS<br>"
End If
rsCheck1.Close : Set rsCheck1 = Nothing

' ==== 2. PointsRedemption ====
Dim rsCheck2
Set rsCheck2 = conn.Execute("SELECT COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID(N'PointsRedemption') AND type='U'")
If rsCheck2(0) = 0 Then
    conn.Execute "CREATE TABLE PointsRedemption (RedemptionID INT IDENTITY(1,1) PRIMARY KEY, ItemName NVARCHAR(100) NOT NULL, ItemType NVARCHAR(30) NOT NULL, PointsCost INT NOT NULL, Stock INT NOT NULL DEFAULT 0, ImageURL NVARCHAR(300) DEFAULT '', RedemptionValue DECIMAL(10,2) NOT NULL DEFAULT 0, MinUserLevel INT NOT NULL DEFAULT 0, IsEnabled BIT NOT NULL DEFAULT 1, SortOrder INT NOT NULL DEFAULT 0, Description NVARCHAR(500) DEFAULT '', Terms NVARCHAR(500) DEFAULT '', CreatedAt DATETIME NOT NULL DEFAULT GETDATE(), UpdatedAt DATETIME NOT NULL DEFAULT GETDATE())"
    Response.Write "PointsRedemption: CREATED<br>"
    conn.Execute "INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES (N'满100减10优惠券','coupon',200,999,10,N'全场通用，满100元可用',1)"
    conn.Execute "INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES (N'满200减30优惠券','coupon',500,500,30,N'全场通用，满200元可用',2)"
    conn.Execute "INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES (N'满500减100优惠券','coupon',1000,200,100,N'全场通用，满500元可用',3)"
    conn.Execute "INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES (N'试用香水小样套装','sample',300,100,25,N'包含3款精选香调小样',4)"
    conn.Execute "INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES (N'限定香水小样礼盒','sample',800,50,80,N'6款热门香调小样，附赠闻香卡',5)"
    conn.Execute "INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES (N'经典方瓶瓶身','bottle',1500,30,150,N'经典方瓶造型，可定制刻字',6)"
    conn.Execute "INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES (N'奢华水晶瓶身','bottle',3000,10,300,N'限量水晶切割工艺瓶身',7)"
    Response.Write "PointsRedemption: DATA INSERTED<br>"
Else
    Response.Write "PointsRedemption: EXISTS<br>"
End If
rsCheck2.Close : Set rsCheck2 = Nothing

' ==== 3. UserPoints ====
Dim rsCheck3
Set rsCheck3 = conn.Execute("SELECT COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID(N'UserPoints') AND type='U'")
If rsCheck3(0) = 0 Then
    conn.Execute "CREATE TABLE UserPoints (PointID INT IDENTITY(1,1) NOT NULL, UserID INT NOT NULL, AvailablePoints INT NULL, TotalPoints INT NULL, UsedPoints INT NULL, ExpiredPoints INT NULL, LastUpdatedAt DATETIME2(7) NULL)"
    Response.Write "UserPoints: CREATED<br>"
Else
    Response.Write "UserPoints: EXISTS<br>"
End If
rsCheck3.Close : Set rsCheck3 = Nothing

' ==== 4. Orders columns ====
Dim rsCheck4
On Error Resume Next
Set rsCheck4 = conn.Execute("SELECT PointsEarned FROM Orders WHERE 1=0")
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE Orders ADD PointsEarned INT NOT NULL DEFAULT 0"
    Response.Write "Orders.PointsEarned: ADDED<br>"
Else
    Response.Write "Orders.PointsEarned: EXISTS<br>"
End If
rsCheck4.Close : Set rsCheck4 = Nothing

Set rsCheck4 = conn.Execute("SELECT PointsRedeemed FROM Orders WHERE 1=0")
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE Orders ADD PointsRedeemed INT NOT NULL DEFAULT 0"
    Response.Write "Orders.PointsRedeemed: ADDED<br>"
Else
    Response.Write "Orders.PointsRedeemed: EXISTS<br>"
End If
rsCheck4.Close : Set rsCheck4 = Nothing

Set rsCheck4 = conn.Execute("SELECT PointsDiscount FROM Orders WHERE 1=0")
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "ALTER TABLE Orders ADD PointsDiscount DECIMAL(10,2) NOT NULL DEFAULT 0"
    Response.Write "Orders.PointsDiscount: ADDED<br>"
Else
    Response.Write "Orders.PointsDiscount: EXISTS<br>"
End If
rsCheck4.Close : Set rsCheck4 = Nothing
On Error GoTo 0

conn.Close : Set conn = Nothing
Response.Write "<br>=== MIGRATION COMPLETE ==="
%>
