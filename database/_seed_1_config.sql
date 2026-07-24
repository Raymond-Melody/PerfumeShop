SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRAN;

/* ===================== 1. AdminRoles（6 个新角色，SUPER_ADMIN 已存在） ===================== */
IF NOT EXISTS(SELECT 1 FROM AdminRoles WHERE RoleCode='OP_ADMIN')
  INSERT INTO AdminRoles(RoleCode,RoleName,Permissions,Description,CreatedAt,UpdatedAt) VALUES('OP_ADMIN',N'运营管理员',N'operation',N'运营/营销/订单/客户/评价/售后',GETDATE(),GETDATE());
IF NOT EXISTS(SELECT 1 FROM AdminRoles WHERE RoleCode='PROD_ADMIN')
  INSERT INTO AdminRoles(RoleCode,RoleName,Permissions,Description,CreatedAt,UpdatedAt) VALUES('PROD_ADMIN',N'生产管理员',N'production',N'半成品/制造/物流/库存',GETDATE(),GETDATE());
IF NOT EXISTS(SELECT 1 FROM AdminRoles WHERE RoleCode='FIN_ADMIN')
  INSERT INTO AdminRoles(RoleCode,RoleName,Permissions,Description,CreatedAt,UpdatedAt) VALUES('FIN_ADMIN',N'财务管理员',N'finance',N'收支/成本/对账/报表',GETDATE(),GETDATE());
IF NOT EXISTS(SELECT 1 FROM AdminRoles WHERE RoleCode='TECH_ADMIN')
  INSERT INTO AdminRoles(RoleCode,RoleName,Permissions,Description,CreatedAt,UpdatedAt) VALUES('TECH_ADMIN',N'技术中心管理员',N'techcenter',N'配方/基香/香调/产品设置',GETDATE(),GETDATE());
IF NOT EXISTS(SELECT 1 FROM AdminRoles WHERE RoleCode='PURCHASE_ADMIN')
  INSERT INTO AdminRoles(RoleCode,RoleName,Permissions,Description,CreatedAt,UpdatedAt) VALUES('PURCHASE_ADMIN',N'采购管理员',N'purchase',N'采购单/供应商/价格/分析',GETDATE(),GETDATE());
IF NOT EXISTS(SELECT 1 FROM AdminRoles WHERE RoleCode='CONTENT_ADMIN')
  INSERT INTO AdminRoles(RoleCode,RoleName,Permissions,Description,CreatedAt,UpdatedAt) VALUES('CONTENT_ADMIN',N'内容管理员',N'operation',N'内容/展示（运营子集）',GETDATE(),GETDATE());

/* ===================== 2. RolePermissions（操作级全权，按角色可访问模块） ===================== */
DECLARE @rp TABLE(RoleCode nvarchar(40), ModuleCode nvarchar(100));
INSERT INTO @rp VALUES
 ('OP_ADMIN','operation'),
 ('CONTENT_ADMIN','operation'),
 ('FIN_ADMIN','finance'),
 ('TECH_ADMIN','techcenter'),
 ('PURCHASE_ADMIN','purchase'),
 ('PROD_ADMIN','prodcenter'),('PROD_ADMIN','semifinished'),('PROD_ADMIN','logistics'),('PROD_ADMIN','inventory'),('PROD_ADMIN','production');
INSERT INTO RolePermissions(RoleID,ModuleCode,CanView,CanCreate,CanEdit,CanDelete,CanExport,CanApprove)
SELECT r.RoleID, x.ModuleCode, 1,1,1,1,1,1
FROM @rp x JOIN AdminRoles r ON r.RoleCode=x.RoleCode
WHERE NOT EXISTS(SELECT 1 FROM RolePermissions p WHERE p.RoleID=r.RoleID AND p.ModuleCode=x.ModuleCode);

/* ===================== 3. ModulePermissions（模块目录 + 表驱动 RequiredRole） ===================== */
DECLARE @mp TABLE(ModuleCode nvarchar(100), ModuleName nvarchar(200), RequiredRole nvarchar(40), Lvl int);
INSERT INTO @mp VALUES
 ('operation',N'运营管理中心',NULL,3),
 ('techcenter',N'产品技术中心','TECH',3),
 ('semifinished',N'半成品生产中心','PROD',3),
 ('prodcenter',N'产品生产中心','PROD',3),
 ('logistics',N'物流管理中心','PROD',3),
 ('inventory',N'库存管理中心','PROD',3),
 ('production',N'生产模块(旧)','PROD',3),
 ('purchase',N'采购管理中心','PURCHASE',3),
 ('finance',N'财务管理中心','FIN',3),
 ('system',N'系统管理中心','SUPER_ADMIN',5);
INSERT INTO ModulePermissions(ModuleCode,ModuleName,RequiredRole,PermissionLevel,IsActive)
SELECT x.ModuleCode,x.ModuleName,x.RequiredRole,x.Lvl,1
FROM @mp x WHERE NOT EXISTS(SELECT 1 FROM ModulePermissions m WHERE m.ModuleCode=x.ModuleCode);

/* ===================== 4. Categories ===================== */
DECLARE @cat TABLE(Nm nvarchar(200), So int);
INSERT INTO @cat VALUES(N'花香调',1),(N'木质调',2),(N'东方调',3),(N'柑橘调',4),(N'清新调',5),(N'美食调',6);
INSERT INTO Categories(CategoryName,IsActive,SortOrder)
SELECT Nm,1,So FROM @cat x WHERE NOT EXISTS(SELECT 1 FROM Categories c WHERE c.CategoryName=x.Nm);

/* ===================== 5. Volumes ===================== */
DECLARE @vol TABLE(ML int, Nm nvarchar(100), Mult float);
INSERT INTO @vol VALUES(10,N'10ml 便携装',0.4),(30,N'30ml 标准装',1.0),(50,N'50ml 经典装',1.6),(100,N'100ml 豪华装',2.8);
INSERT INTO Volumes(VolumeML,VolumeName,PriceMultiplier,IsActive)
SELECT ML,Nm,Mult,1 FROM @vol x WHERE NOT EXISTS(SELECT 1 FROM Volumes v WHERE v.VolumeML=x.ML);

/* ===================== 6. ProductTypeConfig ===================== */
IF NOT EXISTS(SELECT 1 FROM ProductTypeConfig WHERE TypeCode='standard')
  INSERT INTO ProductTypeConfig(TypeCode,DisplayName,NavName,Description,Icon,RequiresReview,RequiresRatio,DisplayOrder,IsActive,CreatedAt) VALUES('standard',N'品牌定香',N'品牌香',N'品牌自营成品香水',N'fa-spray-can',0,0,1,1,GETDATE());
IF NOT EXISTS(SELECT 1 FROM ProductTypeConfig WHERE TypeCode='custom')
  INSERT INTO ProductTypeConfig(TypeCode,DisplayName,NavName,Description,Icon,RequiresReview,RequiresRatio,DisplayOrder,IsActive,CreatedAt) VALUES('custom',N'个性定制',N'定制',N'用户自选前中后调定制',N'fa-flask',0,1,2,1,GETDATE());
IF NOT EXISTS(SELECT 1 FROM ProductTypeConfig WHERE TypeCode='kol')
  INSERT INTO ProductTypeConfig(TypeCode,DisplayName,NavName,Description,Icon,RequiresReview,RequiresRatio,DisplayOrder,IsActive,CreatedAt) VALUES('kol',N'KOL联名',N'KOL',N'达人联名香水',N'fa-star',1,1,3,1,GETDATE());

/* ===================== 7. CostCenters ===================== */
DECLARE @cc TABLE(Cd nvarchar(100), Nm nvarchar(400), Ty nvarchar(100));
INSERT INTO @cc VALUES('CC-PROD',N'生产成本中心',N'Production'),('CC-MKT',N'市场营销中心',N'Marketing'),('CC-LOG',N'物流仓储中心',N'Logistics'),('CC-ADMIN',N'行政管理中心',N'Admin');
INSERT INTO CostCenters(CenterCode,CenterName,CenterType,BudgetAmount,IsActive,CreatedAt,UpdatedAt)
SELECT Cd,Nm,Ty,100000,1,GETDATE(),GETDATE() FROM @cc x WHERE NOT EXISTS(SELECT 1 FROM CostCenters c WHERE c.CenterCode=x.Cd);

/* ===================== 8. ShippingCompanies ===================== */
DECLARE @sc TABLE(Nm nvarchar(200), P nvarchar(100), Ph nvarchar(40));
INSERT INTO @sc VALUES(N'顺丰速运',N'客服',N'95338'),(N'圆通速递',N'客服',N'95554'),(N'中通快递',N'客服',N'95311'),(N'京东物流',N'客服',N'950616');
INSERT INTO ShippingCompanies(CompanyName,ContactPerson,ContactPhone,IsActive,CreatedAt)
SELECT Nm,P,Ph,1,GETDATE() FROM @sc x WHERE NOT EXISTS(SELECT 1 FROM ShippingCompanies s WHERE s.CompanyName=x.Nm);

/* ===================== 9. PurchaseCategories ===================== */
DECLARE @pc TABLE(Cd nvarchar(40), Nm nvarchar(200), So int);
INSERT INTO @pc VALUES('RAW',N'香精原料',1),('BASE',N'基香原料',2),('BOTTLE',N'瓶身器具',3),('PKG',N'包装材料',4),('SPRAY',N'喷头配件',5);
INSERT INTO PurchaseCategories(CategoryCode,CategoryName,Description,DisplayOrder,IsActive)
SELECT Cd,Nm,Nm,So,1 FROM @pc x WHERE NOT EXISTS(SELECT 1 FROM PurchaseCategories c WHERE c.CategoryCode=x.Cd);

/* ===================== 10. MemberTiers + MemberBenefits ===================== */
DECLARE @mt TABLE(Cd varchar(20), Nm nvarchar(100), MinS decimal(18,2), Disc decimal(5,2), FreeShip bit, So int);
INSERT INTO @mt VALUES('BRONZE',N'青铜会员',0,1.00,0,1),('SILVER',N'白银会员',1000,0.98,0,2),('GOLD',N'黄金会员',5000,0.95,1,3),('PLATINUM',N'铂金会员',20000,0.92,1,4),('DIAMOND',N'钻石会员',50000,0.88,1,5);
INSERT INTO MemberTiers(TierCode,TierName,MinSpent,DiscountRate,FreeShipping,PriorityShipping,BirthdayGift,DedicatedSupport,SortOrder,IsActive,CreatedAt)
SELECT Cd,Nm,MinS,Disc,FreeShip,FreeShip,CASE WHEN So>=3 THEN 1 ELSE 0 END,CASE WHEN So>=4 THEN 1 ELSE 0 END,So,1,GETDATE()
FROM @mt x WHERE NOT EXISTS(SELECT 1 FROM MemberTiers t WHERE t.TierCode=x.Cd);
INSERT INTO MemberBenefits(TierCode,BenefitName,BenefitDesc,SortOrder,IsActive)
SELECT Cd,N'专属折扣',N'享受会员专属价格',1,1 FROM @mt x WHERE NOT EXISTS(SELECT 1 FROM MemberBenefits b WHERE b.TierCode=x.Cd AND b.BenefitName=N'专属折扣');

/* ===================== 11. PointsRules ===================== */
DECLARE @pr TABLE(Cd nvarchar(100), Nm nvarchar(200), V decimal(18,2), So int);
INSERT INTO @pr VALUES('EARN_PER_YUAN',N'消费得积分(每元)',1,1),('REDEEM_RATE',N'积分抵现比例(积分/元)',100,2),('SIGNUP_BONUS',N'注册赠送积分',500,3),('REVIEW_BONUS',N'评价赠送积分',50,4);
INSERT INTO PointsRules(RuleCode,RuleName,RuleValue,RuleUnit,IsEnabled,SortOrder,Description,CreatedAt,UpdatedAt)
SELECT Cd,Nm,V,N'点',1,So,Nm,GETDATE(),GETDATE() FROM @pr x WHERE NOT EXISTS(SELECT 1 FROM PointsRules r WHERE r.RuleCode=x.Cd);

/* ===================== 12. Coupons ===================== */
IF NOT EXISTS(SELECT 1 FROM Coupons WHERE CouponCode='WELCOME50')
  INSERT INTO Coupons(CouponCode,CouponName,CouponType,DiscountType,DiscountValue,MinPurchase,MinSpend,MaxDiscount,StartDate,EndDate,ValidFrom,ValidTo,UsageLimit,UsedCount,TotalQty,UsedQty,FirstOrderOnly,IsPublic,IsActive,Description,CreatedAt,UpdatedAt)
  VALUES('WELCOME50',N'新人立减50',N'fixed',N'fixed',50,199,199,50,GETDATE(),DATEADD(month,3,GETDATE()),GETDATE(),DATEADD(month,3,GETDATE()),1000,0,1000,0,1,1,1,N'新用户首单满199减50',GETDATE(),GETDATE());
IF NOT EXISTS(SELECT 1 FROM Coupons WHERE CouponCode='SAVE10PCT')
  INSERT INTO Coupons(CouponCode,CouponName,CouponType,DiscountType,DiscountValue,MinPurchase,MinSpend,MaxDiscount,StartDate,EndDate,ValidFrom,ValidTo,UsageLimit,UsedCount,TotalQty,UsedQty,FirstOrderOnly,IsPublic,IsActive,Description,CreatedAt,UpdatedAt)
  VALUES('SAVE10PCT',N'全场9折',N'percent',N'percent',10,299,299,100,GETDATE(),DATEADD(month,2,GETDATE()),GETDATE(),DATEADD(month,2,GETDATE()),500,0,500,0,0,1,1,N'满299享9折(最高减100)',GETDATE(),GETDATE());

/* ===================== 13. SiteSettings ===================== */
DECLARE @ss TABLE(K nvarchar(100), V nvarchar(510), Nm nvarchar(200));
INSERT INTO @ss VALUES
 ('SiteName',N'YOUR PERFUME 香缇studio',N'站点名称'),
 ('EnableAlipay','1',N'启用支付宝'),
 ('EnableWechatPay','1',N'启用微信支付'),
 ('EnableCOD','1',N'启用货到付款'),
 ('DefaultShippingFee','12',N'默认运费'),
 ('FreeShippingThreshold','199',N'免运费门槛'),
 ('DefaultPackagingCost','5',N'默认包装成本'),
 ('DefaultLaborCost','3',N'默认人工成本'),
 ('EnableLowStockAlert','1',N'启用库存预警'),
 ('Promotion_FirstOrder','1',N'首单优惠开关'),
 ('FEATURE_POINTS','1',N'积分功能开关'),
 ('FEATURE_COUPON','1',N'优惠券功能开关'),
 ('FEATURE_REVIEW','1',N'评价功能开关'),
 ('PointsEarnRate','1',N'每元积分'),
 ('PointsRedeemRate','100',N'积分抵现比例');
INSERT INTO SiteSettings(SettingKey,SettingValue,SettingName,Description,UpdatedAt)
SELECT K,V,Nm,Nm,GETDATE() FROM @ss x WHERE NOT EXISTS(SELECT 1 FROM SiteSettings s WHERE s.SettingKey=x.K);

COMMIT;
PRINT 'SEED_1_CONFIG_OK';
SELECT 'AdminRoles' t,COUNT(*) n FROM AdminRoles UNION ALL SELECT 'RolePermissions',COUNT(*) FROM RolePermissions
UNION ALL SELECT 'ModulePermissions',COUNT(*) FROM ModulePermissions UNION ALL SELECT 'Categories',COUNT(*) FROM Categories
UNION ALL SELECT 'Volumes',COUNT(*) FROM Volumes UNION ALL SELECT 'ProductTypeConfig',COUNT(*) FROM ProductTypeConfig
UNION ALL SELECT 'CostCenters',COUNT(*) FROM CostCenters UNION ALL SELECT 'ShippingCompanies',COUNT(*) FROM ShippingCompanies
UNION ALL SELECT 'PurchaseCategories',COUNT(*) FROM PurchaseCategories UNION ALL SELECT 'MemberTiers',COUNT(*) FROM MemberTiers
UNION ALL SELECT 'MemberBenefits',COUNT(*) FROM MemberBenefits UNION ALL SELECT 'PointsRules',COUNT(*) FROM PointsRules
UNION ALL SELECT 'Coupons',COUNT(*) FROM Coupons UNION ALL SELECT 'SiteSettings',COUNT(*) FROM SiteSettings;
