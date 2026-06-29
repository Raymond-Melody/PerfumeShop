-- ============================================
-- V18 积分与奖励系统 - 数据库迁移
-- 包含: PointsLedger, PointsRules, PointsRedemption
-- 与现有 UserPoints/Users.Points 兼容共存
-- ============================================

-- 1. 积分账本（替代 SiteSettings 日志方式）
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PointsLedger]') AND type in (N'U'))
BEGIN
    CREATE TABLE PointsLedger (
        LedgerID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL,
        Points INT NOT NULL,                          -- 正数=获取, 负数=使用/过期
        PointType NVARCHAR(20) NOT NULL,              -- 'earn','redeem','expire','adjust'
        Source NVARCHAR(30) NOT NULL DEFAULT '',      -- 'purchase','signin','review','share','redeem_discount','redeem_coupon','redeem_sample','redeem_bottle','manual'
        ReferenceID INT NULL,                         -- 关联的 OrderID / ReviewID 等
        Description NVARCHAR(300) DEFAULT '',
        ExpiresAt DATETIME NULL,                      -- 获取的积分过期时间 (12个月滚动)
        IsExpired BIT NOT NULL DEFAULT 0,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        CONSTRAINT FK_PointsLedger_UserID FOREIGN KEY (UserID) REFERENCES Users(UserID)
    );

    -- 索引：按用户+时间查询账本
    CREATE INDEX IX_PointsLedger_UserID_CreatedAt ON PointsLedger(UserID, CreatedAt DESC);
    -- 索引：过期处理
    CREATE INDEX IX_PointsLedger_ExpiresAt ON PointsLedger(ExpiresAt) WHERE ExpiresAt IS NOT NULL AND IsExpired = 0;
END
GO

-- 2. 积分规则配置表
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PointsRules]') AND type in (N'U'))
BEGIN
    CREATE TABLE PointsRules (
        RuleID INT IDENTITY(1,1) PRIMARY KEY,
        RuleCode NVARCHAR(50) NOT NULL UNIQUE,        -- 规则唯一标识
        RuleName NVARCHAR(100) NOT NULL,              -- 规则名称
        RuleValue DECIMAL(10,2) NOT NULL,             -- 规则数值
        RuleUnit NVARCHAR(20) NOT NULL DEFAULT '',    -- 单位: 'points','rate','pct','days'
        IsEnabled BIT NOT NULL DEFAULT 1,
        SortOrder INT NOT NULL DEFAULT 0,
        Description NVARCHAR(200) DEFAULT '',
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        UpdatedAt DATETIME NOT NULL DEFAULT GETDATE()
    );

    -- 默认规则数据
    INSERT INTO PointsRules (RuleCode, RuleName, RuleValue, RuleUnit, SortOrder, Description) VALUES
    ('purchase_rate',       N'消费积分比例',   1,      'points', 1, N'每消费1元获得积分数'),
    ('signin_points',       N'签到积分',       5,      'points', 2, N'每日签到奖励积分'),
    ('review_points',       N'评价积分',       20,     'points', 3, N'发表有效评价奖励积分'),
    ('review_with_photo',   N'带图评价积分',   10,     'points', 4, N'带图评价额外奖励积分'),
    ('share_points',        N'分享积分',       10,     'points', 5, N'分享产品/订单获得积分'),
    ('referral_points',     N'推荐注册积分',   100,    'points', 6, N'推荐好友成功注册获得积分'),
    ('referral_purchase',   N'推荐消费积分',   50,     'points', 7, N'推荐好友首单获得积分'),
    ('redeem_discount_rate',N'积分抵扣比例',   100,    'rate',   10,N'每100积分可抵扣1元'),
    ('max_redeem_pct',      N'最大抵扣比例',   30,     'pct',    11,N'积分抵扣最高占订单金额百分比'),
    ('points_expire_months',N'积分有效期(月)', 12,     'days',   12,N'积分获取后有效期(月)，0=永不过期');
END
GO

-- 3. 积分兑换商品表
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PointsRedemption]') AND type in (N'U'))
BEGIN
    CREATE TABLE PointsRedemption (
        RedemptionID INT IDENTITY(1,1) PRIMARY KEY,
        ItemName NVARCHAR(100) NOT NULL,
        ItemType NVARCHAR(30) NOT NULL,               -- 'coupon','sample','bottle','discount'
        PointsCost INT NOT NULL,                      -- 所需积分
        Stock INT NOT NULL DEFAULT 0,                 -- 库存
        ImageURL NVARCHAR(300) DEFAULT '',
        RedemptionValue DECIMAL(10,2) NOT NULL DEFAULT 0, -- 兑换价值(元)
        MinUserLevel INT NOT NULL DEFAULT 0,          -- 最低会员等级要求(0=无限制)
        IsEnabled BIT NOT NULL DEFAULT 1,
        SortOrder INT NOT NULL DEFAULT 0,
        Description NVARCHAR(500) DEFAULT '',
        Terms NVARCHAR(500) DEFAULT '',               -- 使用条款
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        UpdatedAt DATETIME NOT NULL DEFAULT GETDATE()
    );

    -- 默认兑换数据
    INSERT INTO PointsRedemption (ItemName, ItemType, PointsCost, Stock, RedemptionValue, Description, SortOrder) VALUES
    (N'满100减10优惠券',   'coupon',  200,  999,  10,  N'全场通用，满100元可用', 1),
    (N'满200减30优惠券',   'coupon',  500,  500,  30,  N'全场通用，满200元可用', 2),
    (N'满500减100优惠券',  'coupon',  1000, 200,  100, N'全场通用，满500元可用', 3),
    (N'试用香水小样套装',  'sample',  300,  100,  25,  N'包含3款精选香调小样', 4),
    (N'限定香水小样礼盒',  'sample',  800,  50,   80,  N'6款热门香调小样，附赠闻香卡', 5),
    (N'经典方瓶瓶身',     'bottle',  1500, 30,   150, N'经典方瓶造型，可定制刻字', 6),
    (N'奢华水晶瓶身',     'bottle',  3000, 10,   300, N'限量水晶切割工艺瓶身', 7);
END
GO

-- 4. 扩展 Orders 表：积分相关字段（如不存在则添加）
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Orders') AND name = 'PointsEarned')
BEGIN
    ALTER TABLE Orders ADD PointsEarned INT NOT NULL DEFAULT 0;
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Orders') AND name = 'PointsRedeemed')
BEGIN
    ALTER TABLE Orders ADD PointsRedeemed INT NOT NULL DEFAULT 0;
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Orders') AND name = 'PointsDiscount')
BEGIN
    ALTER TABLE Orders ADD PointsDiscount DECIMAL(10,2) NOT NULL DEFAULT 0;
END
GO
