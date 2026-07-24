SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
SET XACT_ABORT ON;

/* ===== 生成 V3 密码哈希（与 password_utils.asp 完全一致：1000 轮 SHA-256 + PEPPER） ===== */
DECLARE @pw VARCHAR(100)='Ray123456';
DECLARE @pepper VARCHAR(200)='P3rfum3Sh0p_S@lt_2026!';
DECLARE @salt VARCHAR(64)='SeedMockEnv2026SaltForAllTestAcct';
DECLARE @h VARCHAR(64)=LOWER(CONVERT(VARCHAR(64),HASHBYTES('SHA2_256',@salt+@pw+@pepper),2));
DECLARE @i INT=1;
WHILE @i<=999 BEGIN SET @h=LOWER(CONVERT(VARCHAR(64),HASHBYTES('SHA2_256',@h+@salt),2)); SET @i+=1; END
DECLARE @v3 VARCHAR(200)='V3$'+@salt+'$'+@h;
PRINT 'V3 hash: '+@v3;

BEGIN TRAN;

/* ===================== 前端用户（Raymond 已存在，UserID=25） ===================== */
DECLARE @u TABLE(Username nvarchar(100), Email nvarchar(200), FullName nvarchar(200), Phone nvarchar(40), Tier nvarchar(40), Spent decimal(18,2), Orders int, VIP bit);
INSERT INTO @u VALUES
 ('alice',N'alice@example.com',N'爱丽丝','13800000001','SILVER',1500,3,0),
 ('bob',N'bob@example.com',N'鲍勃','13800000002','GOLD',6800,8,1),
 ('carol',N'carol@example.com',N'卡罗尔','13800000003','BRONZE',200,1,0),
 ('david',N'david@example.com',N'大卫','13800000004','PLATINUM',25000,20,1),
 ('emma',N'emma@example.com',N'艾玛','13800000005','DIAMOND',60000,45,1),
 ('frank',N'frank@example.com',N'弗兰克','13800000006','BRONZE',0,0,0);
INSERT INTO Users(Username,Email,[Password],FullName,Phone,IsActive,IsVIP,Points,UserRole,CustomerTier,TotalSpent,OrderCount,CreatedAt)
SELECT x.Username,x.Email,@v3,x.FullName,x.Phone,1,x.VIP,CAST(x.Spent AS INT),N'member',x.Tier,x.Spent,x.Orders,GETDATE()
FROM @u x WHERE NOT EXISTS(SELECT 1 FROM Users u WHERE u.Username=x.Username);

/* 确保 Raymond 也有点数/等级（保留账户补全画像） */
UPDATE Users SET CustomerTier=ISNULL(CustomerTier,'GOLD'), UserRole=ISNULL(UserRole,'member'), IsActive=1 WHERE Username='Raymond';

/* UserPoints + UserAddresses（为每个前端用户补全） */
INSERT INTO UserPoints(UserID,TotalPoints,AvailablePoints,UsedPoints,ExpiredPoints,LastUpdatedAt)
SELECT u.UserID, ISNULL(u.Points,0), ISNULL(u.Points,0), 0, 0, GETDATE()
FROM Users u WHERE NOT EXISTS(SELECT 1 FROM UserPoints p WHERE p.UserID=u.UserID);

INSERT INTO UserAddresses(UserID,Consignee,Phone,Province,City,District,Address,IsDefault,CreatedAt)
SELECT u.UserID, ISNULL(u.FullName,u.Username), ISNULL(u.Phone,'13900000000'), N'广东省', N'深圳市', N'南山区', N'科技园路100号'+CAST(u.UserID AS nvarchar(10))+N'栋', 1, GETDATE()
FROM Users u WHERE NOT EXISTS(SELECT 1 FROM UserAddresses a WHERE a.UserID=u.UserID);

/* ===================== 后端管理员（每角色 1 个，Ray88 已存在=SUPER_ADMIN） ===================== */
DECLARE @a TABLE(Username nvarchar(100), Email nvarchar(200), FullName nvarchar(200), Dept nvarchar(100), RoleCode nvarchar(40));
INSERT INTO @a VALUES
 ('op_manager',N'op@perfume.local',N'运营主管',N'运营部','OP_ADMIN'),
 ('prod_manager',N'prod@perfume.local',N'生产主管',N'生产部','PROD_ADMIN'),
 ('fin_manager',N'fin@perfume.local',N'财务主管',N'财务部','FIN_ADMIN'),
 ('tech_manager',N'tech@perfume.local',N'技术主管',N'技术中心','TECH_ADMIN'),
 ('purchase_manager',N'purchase@perfume.local',N'采购主管',N'采购部','PURCHASE_ADMIN'),
 ('content_editor',N'content@perfume.local',N'内容编辑',N'运营部','CONTENT_ADMIN');
INSERT INTO AdminUsers(Username,Email,PasswordHash,FullName,Department,RoleID,IsActive,IsLocked,CreatedAt)
SELECT x.Username,x.Email,@v3,x.FullName,x.Dept,r.RoleID,1,0,GETDATE()
FROM @a x JOIN AdminRoles r ON r.RoleCode=x.RoleCode
WHERE NOT EXISTS(SELECT 1 FROM AdminUsers au WHERE au.Username=x.Username);

COMMIT;
PRINT 'SEED_2_USERS_OK';
SELECT 'Users' t,COUNT(*) n FROM Users UNION ALL SELECT 'UserPoints',COUNT(*) FROM UserPoints
UNION ALL SELECT 'UserAddresses',COUNT(*) FROM UserAddresses UNION ALL SELECT 'AdminUsers',COUNT(*) FROM AdminUsers;
SELECT a.Username, r.RoleCode, a.IsActive FROM AdminUsers a LEFT JOIN AdminRoles r ON r.RoleID=a.RoleID ORDER BY a.AdminID;
