-- ============================================
-- V18 限时活动引擎 (Flash Sale + Group Buy)
-- 创建时间: 2026-06-30
-- ============================================

-- ============================================
-- 秒杀活动表
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FlashSale]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[FlashSale] (
        FlashSaleID INT IDENTITY(1,1) PRIMARY KEY,
        ProductID INT NOT NULL,
        FlashPrice DECIMAL(12,2) NOT NULL,          -- 秒杀价
        Stock INT NOT NULL DEFAULT 0,                -- 秒杀库存（剩余）
        SoldCount INT NOT NULL DEFAULT 0,            -- 已售数量
        LimitPerUser INT NOT NULL DEFAULT 1,         -- 每人限购数量
        StartTime DATETIME NOT NULL,                 -- 开始时间
        EndTime DATETIME NOT NULL,                   -- 结束时间
        IsActive BIT NOT NULL DEFAULT 1,             -- 是否启用
        SortOrder INT NOT NULL DEFAULT 0,            -- 排序
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (ProductID) REFERENCES [Products](ProductID)
    );

    CREATE INDEX IX_FlashSale_Time ON [FlashSale](IsActive, StartTime, EndTime);
    CREATE INDEX IX_FlashSale_Product ON [FlashSale](ProductID);
END
GO

-- ============================================
-- 拼团活动计划表
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GroupBuyPlans]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[GroupBuyPlans] (
        PlanID INT IDENTITY(1,1) PRIMARY KEY,
        ProductID INT NOT NULL,
        TeamSize INT NOT NULL DEFAULT 2,             -- 成团人数 (2/3/5人)
        GroupPrice DECIMAL(12,2) NOT NULL,            -- 拼团价（每人）
        MinUnit INT NOT NULL DEFAULT 1,               -- 最低开团数
        MaxUnit INT NOT NULL DEFAULT 0,               -- 最高开团数（0=不限）
        StartTime DATETIME NOT NULL,
        EndTime DATETIME NOT NULL,
        DurationHours INT NOT NULL DEFAULT 24,        -- 成团有效期（小时），过期自动失败
        IsActive BIT NOT NULL DEFAULT 1,
        SortOrder INT NOT NULL DEFAULT 0,
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (ProductID) REFERENCES [Products](ProductID)
    );

    CREATE INDEX IX_GroupBuyPlans_Time ON [GroupBuyPlans](IsActive, StartTime, EndTime);
    CREATE INDEX IX_GroupBuyPlans_Product ON [GroupBuyPlans](ProductID);
END
GO

-- ============================================
-- 拼团订单/团表
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GroupBuyOrders]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[GroupBuyOrders] (
        GroupID INT IDENTITY(1,1) PRIMARY KEY,
        PlanID INT NOT NULL,
        GroupSN VARCHAR(20) NOT NULL,                 -- 团唯一编号 (如 GB20260630XXXX)
        InitiatorID INT NOT NULL,                     -- 团长（发起人UserID）
        Status INT NOT NULL DEFAULT 0,                -- 0=进行中, 1=已成团, 2=已失效/未成团, 3=已退款
        CurrentSize INT NOT NULL DEFAULT 1,           -- 当前参团人数
        TargetSize INT NOT NULL,                      -- 目标成团人数
        CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
        CompletedAt DATETIME NULL,                    -- 完成时间（成团或失败）
        FOREIGN KEY (PlanID) REFERENCES [GroupBuyPlans](PlanID)
    );

    CREATE INDEX IX_GroupBuyOrders_Plan ON [GroupBuyOrders](PlanID, Status);
    CREATE INDEX IX_GroupBuyOrders_SN ON [GroupBuyOrders](GroupSN);
END
GO

-- ============================================
-- 拼团参团记录表
-- ============================================
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GroupBuyParticipants]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[GroupBuyParticipants] (
        ParticipantID INT IDENTITY(1,1) PRIMARY KEY,
        GroupID INT NOT NULL,
        UserID INT NOT NULL,
        OrderID INT NULL,                             -- 关联的订单ID
        IsInitiator BIT NOT NULL DEFAULT 0,           -- 是否是团长
        Status INT NOT NULL DEFAULT 0,                -- 0=待支付, 1=已支付, 2=已退款
        JoinedAt DATETIME NOT NULL DEFAULT GETDATE(),
        FOREIGN KEY (GroupID) REFERENCES [GroupBuyOrders](GroupID)
    );

    CREATE INDEX IX_GroupBuyParticipants_Group ON [GroupBuyParticipants](GroupID);
    CREATE INDEX IX_GroupBuyParticipants_User ON [GroupBuyParticipants](UserID);
END
GO
