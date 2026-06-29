-- ============================================
-- V18.0 会员等级体系 (Member Tiers System)
-- 任务: Task 11 - 会员等级 + 权益
-- 等级: 银卡(0-3000) → 金卡(3000-10000) → 钻石(10000-30000) → 黑金(30000+)
-- ============================================

-- 1. 会员等级配置表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MemberTiers')
BEGIN
    CREATE TABLE MemberTiers (
        TierID INT IDENTITY(1,1) PRIMARY KEY,
        TierCode VARCHAR(20) NOT NULL UNIQUE,
        TierName NVARCHAR(50) NOT NULL,
        TierNameEN NVARCHAR(50) NULL,
        MinSpent DECIMAL(10,2) NOT NULL DEFAULT 0,
        MaxSpent DECIMAL(10,2) NULL,
        DiscountRate DECIMAL(4,3) NOT NULL DEFAULT 1.000,
        FreeShipping BIT NOT NULL DEFAULT 0,
        PriorityShipping BIT NOT NULL DEFAULT 0,
        BirthdayGift BIT NOT NULL DEFAULT 0,
        DedicatedSupport BIT NOT NULL DEFAULT 0,
        IconClass VARCHAR(50) NULL,
        Color VARCHAR(20) NULL,
        BadgeBg VARCHAR(20) NULL,
        SortOrder INT NOT NULL DEFAULT 0,
        IsActive BIT NOT NULL DEFAULT 1,
        CreatedAt DATETIME2 NOT NULL DEFAULT GETDATE(),
        UpdatedAt DATETIME2 NULL
    );

    -- 插入默认等级数据
    INSERT INTO MemberTiers (TierCode, TierName, TierNameEN, MinSpent, MaxSpent, DiscountRate, FreeShipping, PriorityShipping, BirthdayGift, DedicatedSupport, IconClass, Color, BadgeBg, SortOrder)
    VALUES
    ('silver',  N'银卡会员', 'Silver',   0,      3000,   0.95, 0, 0, 0, 0, 'fa-medal',          '#9E9E9E', '#f5f5f5', 1),
    ('gold',    N'金卡会员', 'Gold',     3000,   10000,  0.90, 1, 0, 0, 0, 'fa-star',           '#FF9800', '#fff3e0', 2),
    ('diamond', N'钻石会员', 'Diamond',  10000,  30000,  0.85, 1, 1, 1, 0, 'fa-gem',            '#2196F3', '#e3f2fd', 3),
    ('black',   N'黑金会员', 'Black Gold',30000, NULL,   0.80, 1, 1, 1, 1, 'fa-crown',          '#212121', '#fafafa', 4);
END
GO

-- 2. 会员权益明细表
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MemberBenefits')
BEGIN
    CREATE TABLE MemberBenefits (
        BenefitID INT IDENTITY(1,1) PRIMARY KEY,
        TierCode VARCHAR(20) NOT NULL,
        BenefitName NVARCHAR(100) NOT NULL,
        BenefitDesc NVARCHAR(500) NULL,
        BenefitIcon VARCHAR(50) NULL DEFAULT 'fa-check-circle',
        SortOrder INT NOT NULL DEFAULT 0,
        IsActive BIT NOT NULL DEFAULT 1,
        CONSTRAINT FK_MemberBenefits_Tier FOREIGN KEY (TierCode) REFERENCES MemberTiers(TierCode)
    );

    -- 银卡权益
    INSERT INTO MemberBenefits (TierCode, BenefitName, BenefitDesc, BenefitIcon, SortOrder) VALUES
    ('silver', N'9.5折优惠', N'全场商品享受9.5折专属折扣', 'fa-tag', 1),
    ('silver', N'积分加速', N'每消费1元获得1.2积分', 'fa-bolt', 2),
    ('silver', N'新品优先浏览', N'新品上架前24小时可提前预览', 'fa-eye', 3);

    -- 金卡权益
    INSERT INTO MemberBenefits (TierCode, BenefitName, BenefitDesc, BenefitIcon, SortOrder) VALUES
    ('gold', N'9折优惠', N'全场商品享受9折专属折扣', 'fa-tag', 1),
    ('gold', N'全场免运费', N'不限金额，所有订单免配送费', 'fa-truck', 2),
    ('gold', N'积分加速', N'每消费1元获得1.5积分', 'fa-bolt', 3),
    ('gold', N'新品优先浏览', N'新品上架前48小时可提前预览', 'fa-eye', 4),
    ('gold', N'专属包装', N'金卡专属礼盒包装', 'fa-gift', 5);

    -- 钻石权益
    INSERT INTO MemberBenefits (TierCode, BenefitName, BenefitDesc, BenefitIcon, SortOrder) VALUES
    ('diamond', N'8.5折优惠', N'全场商品享受8.5折专属折扣', 'fa-tag', 1),
    ('diamond', N'全场免运费', N'不限金额，所有订单免配送费', 'fa-truck', 2),
    ('diamond', N'积分加速', N'每消费1元获得2积分', 'fa-bolt', 3),
    ('diamond', N'优先发货', N'订单优先处理和发货', 'fa-rocket', 4),
    ('diamond', N'生日专属礼', N'生日月赠送专属定制小样', 'fa-cake', 5),
    ('diamond', N'新品抢先体验', N'新品上架前72小时可提前购买', 'fa-star', 6),
    ('diamond', N'专属包装', N'钻石会员专属奢华礼盒', 'fa-gift', 7);

    -- 黑金权益
    INSERT INTO MemberBenefits (TierCode, BenefitName, BenefitDesc, BenefitIcon, SortOrder) VALUES
    ('black', N'8折优惠', N'全场商品享受8折专属折扣', 'fa-tag', 1),
    ('black', N'全场免运费', N'不限金额，所有订单免配送费', 'fa-truck', 2),
    ('black', N'积分加速', N'每消费1元获得3积分', 'fa-bolt', 3),
    ('black', N'优先发货', N'订单最优先处理和发货', 'fa-rocket', 4),
    ('black', N'生日专属礼', N'生日月赠送专属正装香水', 'fa-cake', 5),
    ('black', N'新品抢先体验', N'新品上架前7天可提前购买', 'fa-star', 6),
    ('black', N'专属客服', N'1对1专属调香顾问服务', 'fa-headset', 7),
    ('black', N'限量新品优先购', N'限量版香水优先购买权', 'fa-crown', 8),
    ('black', N'专属包装', N'黑金会员专属定制奢华礼盒', 'fa-gift', 9),
    ('black', N'线下活动邀请', N'获邀参加品牌线下品香活动', 'fa-calendar-check', 10);
END
GO

-- 3. 用户等级记录表（记录等级变更历史）
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'UserTierHistory')
BEGIN
    CREATE TABLE UserTierHistory (
        HistoryID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL,
        OldTierCode VARCHAR(20) NULL,
        NewTierCode VARCHAR(20) NOT NULL,
        TotalSpent DECIMAL(10,2) NOT NULL DEFAULT 0,
        ChangeType VARCHAR(20) NOT NULL DEFAULT 'auto',
        ChangedAt DATETIME2 NOT NULL DEFAULT GETDATE()
    );
END
GO

PRINT 'V18 Member Tiers tables created successfully.';
GO
