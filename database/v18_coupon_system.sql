-- ============================================
-- V18 优惠券与促销引擎 - 数据库迁移
-- 迁移: SiteSettings 键值对 → 专用 Coupons/UserCoupons 表
-- 兼容: 现有 Promotion_Threshold/Promotion_FirstOrder 等 SiteSettings 键
-- ============================================

-- 1. 优惠券定义表
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Coupons]') AND type in (N'U'))
BEGIN
    CREATE TABLE Coupons (
        CouponID INT IDENTITY(1,1) PRIMARY KEY,
        CouponCode NVARCHAR(50) NOT NULL UNIQUE,       -- 唯一优惠码
        CouponName NVARCHAR(100) NOT NULL,              -- 优惠券名称
        CouponType NVARCHAR(20) NOT NULL DEFAULT 'fixed', -- fixed/percentage/free_shipping/gift
        DiscountValue DECIMAL(10,2) NOT NULL,           -- 优惠值: 金额/百分比/0
        MinSpend DECIMAL(10,2) NOT NULL DEFAULT 0,      -- 最低消费门槛
        MaxDiscount DECIMAL(10,2) NOT NULL DEFAULT 0,   -- 最大优惠金额(百分比券用)
        ValidFrom DATETIME NOT NULL,                    -- 有效期起始
        ValidTo DATETIME NOT NULL,                      -- 有效期截止
        TotalQty INT NOT NULL DEFAULT 0,                -- 总发行量(0=不限)
        UsedQty INT NOT NULL DEFAULT 0,                 -- 已使用量
        IsActive BIT NOT NULL DEFAULT 1,
        ApplicableCategory NVARCHAR(50) DEFAULT NULL,   -- 限定品类(NULL=全场)
        ApplicableProductID INT NULL,                   -- 限定产品(NULL=全场)
        FirstOrderOnly BIT NOT NULL DEFAULT 0,          -- 仅限首单
        Description NVARCHAR(500) DEFAULT '',
        Terms NVARCHAR(500) DEFAULT '',                 -- 使用条款
        IsPublic BIT NOT NULL DEFAULT 1,                -- 是否公开可领
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        UpdatedAt DATETIME NOT NULL DEFAULT GETDATE()
    );

    CREATE INDEX IX_Coupons_Code ON Coupons(CouponCode) WHERE IsActive = 1;
    CREATE INDEX IX_Coupons_ValidTo ON Coupons(ValidTo) WHERE IsActive = 1;

    -- 默认优惠券数据
    INSERT INTO Coupons (CouponCode, CouponName, CouponType, DiscountValue, MinSpend, MaxDiscount, ValidFrom, ValidTo, TotalQty, Description) VALUES
    ('WELCOME10',  N'新人欢迎券',   'fixed',          10,   100,  0,   '2025-01-01', '2027-12-31', 999, N'全场满100减10'),
    ('WELCOME20',  N'新人满减券',   'fixed',          20,   200,  0,   '2025-01-01', '2027-12-31', 500, N'全场满200减20'),
    ('VIP5',       N'会员95折券',   'percentage',     5,    0,    50,  '2025-01-01', '2027-12-31', 200, N'全场95折，最高优惠50元'),
    ('FREESHIP',   N'免邮券',       'free_shipping',  0,    0,    0,   '2025-01-01', '2027-12-31', 100, N'全场订单免运费'),
    ('BIRTHDAY50', N'生日礼券',     'fixed',          50,   0,    0,   '2025-01-01', '2027-12-31', 0,   N'生日快乐！无门槛50元优惠券'),
    ('SUMMER15',   N'夏日香氛券',   'percentage',     15,   300,  80,  '2025-06-01', '2027-09-30', 300, N'夏日香氛专场85折，最高80元'),
    ('TIER_GOLD',  N'金卡升级礼',   'percentage',     10,   0,    100, '2025-01-01', '2027-12-31', 0,   N'金卡会员升级礼遇，全场9折最高100元');
END
GO

-- 2. 用户优惠券持有表
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UserCoupons]') AND type in (N'U'))
BEGIN
    CREATE TABLE UserCoupons (
        UserCouponID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL,
        CouponID INT NOT NULL,
        CouponCode NVARCHAR(50) NOT NULL,               -- 去规范化冗余
        Source NVARCHAR(30) NOT NULL DEFAULT 'manual',  -- new_user/activity/points_redeem/tier_upgrade/manual
        Status NVARCHAR(10) NOT NULL DEFAULT 'available', -- available/used/expired
        UsedAt DATETIME NULL,
        UsedOrderID INT NULL,
        ObtainedAt DATETIME NOT NULL DEFAULT GETDATE(),
        ExpiresAt DATETIME NULL,                        -- 个人过期时间(可不同于券本身)
        CONSTRAINT FK_UserCoupons_UserID FOREIGN KEY (UserID) REFERENCES Users(UserID),
        CONSTRAINT FK_UserCoupons_CouponID FOREIGN KEY (CouponID) REFERENCES Coupons(CouponID)
    );

    CREATE INDEX IX_UserCoupons_UserID_Status ON UserCoupons(UserID, Status);
    CREATE INDEX IX_UserCoupons_Code ON UserCoupons(CouponCode);
END
GO

-- 3. 扩展 Orders 表：优惠券字段（如不存在则添加）
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Orders') AND name = 'CouponCode')
BEGIN
    ALTER TABLE Orders ADD CouponCode NVARCHAR(50) NULL;
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Orders') AND name = 'CouponDiscount')
BEGIN
    ALTER TABLE Orders ADD CouponDiscount DECIMAL(10,2) NOT NULL DEFAULT 0;
END
GO
