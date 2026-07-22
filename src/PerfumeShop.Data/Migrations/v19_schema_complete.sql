-- ============================================================================
-- V19 Schema Completion Migration Script
-- Adds V18 coupon/points/community fields to Orders and ProductReviews,
-- creates ReviewImages table.
-- Idempotent: safe to run multiple times.
-- ============================================================================

-- 1. Orders: add coupon-related columns (from v18_coupon_system.sql)
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'CouponCode')
    ALTER TABLE Orders ADD CouponCode NVARCHAR(50) NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'CouponDiscount')
    ALTER TABLE Orders ADD CouponDiscount DECIMAL(18,2) NULL;

-- 2. Orders: add points-related columns (from v18_points_system.sql)
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'PointsEarned')
    ALTER TABLE Orders ADD PointsEarned INT NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'PointsRedeemed')
    ALTER TABLE Orders ADD PointsRedeemed INT NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Orders') AND name = 'PointsDiscount')
    ALTER TABLE Orders ADD PointsDiscount DECIMAL(18,2) NULL;

-- 3. ProductReviews: add community UGC columns (from v18_community_ugc.sql)
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ProductReviews') AND name = 'Title')
    ALTER TABLE ProductReviews ADD Title NVARCHAR(200) NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ProductReviews') AND name = 'IsVerifiedPurchase')
    ALTER TABLE ProductReviews ADD IsVerifiedPurchase BIT NOT NULL DEFAULT 0;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ProductReviews') AND name = 'AIFeelingSummary')
    ALTER TABLE ProductReviews ADD AIFeelingSummary NVARCHAR(2000) NULL;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('ProductReviews') AND name = 'LikeCount')
    ALTER TABLE ProductReviews ADD LikeCount INT NOT NULL DEFAULT 0;

-- 4. Create ReviewImages table for review image attachments
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ReviewImages')
BEGIN
    CREATE TABLE ReviewImages (
        ImageId       INT           IDENTITY(1,1) NOT NULL,
        ReviewId      INT           NOT NULL,
        ImageUrl      NVARCHAR(500) NOT NULL,
        SortOrder     INT           NOT NULL DEFAULT 0,
        CreatedAt     DATETIME2     NOT NULL DEFAULT SYSUTCDATETIME(),
        CONSTRAINT PK_ReviewImages PRIMARY KEY CLUSTERED (ImageId),
        CONSTRAINT FK_ReviewImages_ProductReviews FOREIGN KEY (ReviewId)
            REFERENCES ProductReviews (ReviewId) ON DELETE CASCADE
    ) WITH (ONLINE = ON);
END;

-- 5. Index on ReviewImages.ReviewId for fast lookups
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('ReviewImages') AND name = 'IX_ReviewImages_ReviewId')
    CREATE NONCLUSTERED INDEX IX_ReviewImages_ReviewId ON ReviewImages (ReviewId) WITH (ONLINE = ON);
