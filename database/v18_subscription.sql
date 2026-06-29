-- ============================================
-- V18 订阅制香氛盒 (Subscription Box)
-- 创建时间: 2026-06-30
-- ============================================

-- ============================================
-- 订阅计划表
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SubscriptionPlans]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[SubscriptionPlans] (
        PlanID INT IDENTITY(1,1) PRIMARY KEY,
        PlanName NVARCHAR(100) NOT NULL,              -- 计划名称
        Period VARCHAR(20) NOT NULL,                   -- monthly/quarterly/yearly
        Price DECIMAL(12,2) NOT NULL,                  -- 每期价格
        SampleCount INT NOT NULL DEFAULT 3,            -- 每期小样数量
        FullSizeCount INT NOT NULL DEFAULT 1,          -- 每期正装数量
        FreeShipping BIT NOT NULL DEFAULT 1,           -- 是否包邮
        CancellationFee DECIMAL(12,2) NOT NULL DEFAULT 0, -- 取消费用
        IsActive BIT NOT NULL DEFAULT 1,
        SortOrder INT NOT NULL DEFAULT 0,
        Description NVARCHAR(500) NULL,                -- 计划描述
        FeaturedImage VARCHAR(500) NULL,               -- 展示图
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE()
    );
END
GO

-- ============================================
-- 用户订阅表
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UserSubscriptions]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[UserSubscriptions] (
        SubscriptionID INT IDENTITY(1,1) PRIMARY KEY,
        UserID INT NOT NULL,
        PlanID INT NOT NULL,
        Status INT NOT NULL DEFAULT 0,                -- 0=活跃, 1=暂停, 2=已取消, 3=已过期
        StartDate DATETIME NOT NULL DEFAULT GETDATE(),
        NextDeliveryDate DATETIME NOT NULL,            -- 下次配送日期
        EndDate DATETIME NULL,                         -- 结束日期（取消/过期时设置）
        TotalDeliveries INT NOT NULL DEFAULT 0,       -- 已完成配送次数
        AutoRenew BIT NOT NULL DEFAULT 1,              -- 自动续费
        PauseNote NVARCHAR(200) NULL,                  -- 暂停原因
        CancelReason NVARCHAR(200) NULL,               -- 取消原因
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        UpdatedAt DATETIME NULL,
        FOREIGN KEY (UserID) REFERENCES [Users](UserID),
        FOREIGN KEY (PlanID) REFERENCES [SubscriptionPlans](PlanID)
    );

    CREATE INDEX IX_UserSubscriptions_User ON [UserSubscriptions](UserID, Status);
    CREATE INDEX IX_UserSubscriptions_NextDelivery ON [UserSubscriptions](NextDeliveryDate) WHERE Status = 0;
END
GO

-- ============================================
-- 订阅配送记录表
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SubscriptionDeliveries]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[SubscriptionDeliveries] (
        DeliveryID INT IDENTITY(1,1) PRIMARY KEY,
        SubscriptionID INT NOT NULL,
        DeliveryDate DATETIME NOT NULL,                -- 计划配送日期
        Status INT NOT NULL DEFAULT 0,                 -- 0=待配送, 1=已发货, 2=已签收, 3=已跳过, 4=已退回
        OrderID INT NULL,                              -- 关联订单ID
        Contents NVARCHAR(1000) NULL,                  -- 配送内容描述
        TrackingNumber VARCHAR(100) NULL,              -- 物流单号
        ShippedAt DATETIME NULL,                       -- 发货时间
        DeliveredAt DATETIME NULL,                     -- 签收时间
        SkippedAt DATETIME NULL,                       -- 跳过时间
        SkipReason NVARCHAR(200) NULL,                 -- 跳过原因
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (SubscriptionID) REFERENCES [UserSubscriptions](SubscriptionID)
    );

    CREATE INDEX IX_SubscriptionDeliveries_Sub ON [SubscriptionDeliveries](SubscriptionID);
    CREATE INDEX IX_SubscriptionDeliveries_Date ON [SubscriptionDeliveries](DeliveryDate);
END
GO

-- ============================================
-- 种子数据：三种订阅计划
-- ============================================
IF NOT EXISTS (SELECT 1 FROM SubscriptionPlans)
BEGIN
    INSERT INTO SubscriptionPlans (PlanName, Period, Price, SampleCount, FullSizeCount, FreeShipping, CancellationFee, Description, SortOrder)
    VALUES
    (N'月度探索盒', 'monthly', 199.00, 3, 1, 1, 0, N'每月配送3款精选小样+1款正装香水，AI根据您的偏好个性化选品。随时可取消。', 1),
    (N'季度臻选盒', 'quarterly', 549.00, 4, 2, 1, 0, N'每季度配送4款新品小样+2款正装香水，尊享8折优惠。赠送限定瓶型。', 2),
    (N'年度尊享盒', 'yearly', 1999.00, 5, 3, 1, 0, N'全年12次配送，每次5款独家小样+3款正装香水。专属调香师一对一服务。', 3);
END
GO
