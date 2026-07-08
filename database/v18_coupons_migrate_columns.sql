-- ============================================
-- V18 优惠券表列迁移：旧schema → 新schema
-- 旧表有: CouponCode, CouponID, CreatedAt, DiscountType, DiscountValue, 
--          EndDate, IsActive, MinPurchase, StartDate, UsageLimit, UsedCount
-- 需新增/重命名匹配V18代码
-- ============================================

-- 1. 新增缺失列（如不存在则添加）
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'CouponName')
    ALTER TABLE Coupons ADD CouponName NVARCHAR(100) NOT NULL DEFAULT '';

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'CouponType')
    ALTER TABLE Coupons ADD CouponType NVARCHAR(20) NOT NULL DEFAULT 'fixed';

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'MinSpend')
    ALTER TABLE Coupons ADD MinSpend DECIMAL(10,2) NOT NULL DEFAULT 0;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'MaxDiscount')
    ALTER TABLE Coupons ADD MaxDiscount DECIMAL(10,2) NOT NULL DEFAULT 0;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'ValidFrom')
    ALTER TABLE Coupons ADD ValidFrom DATETIME NOT NULL DEFAULT GETDATE();

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'ValidTo')
    ALTER TABLE Coupons ADD ValidTo DATETIME NOT NULL DEFAULT DATEADD(YEAR, 1, GETDATE());

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'TotalQty')
    ALTER TABLE Coupons ADD TotalQty INT NOT NULL DEFAULT 0;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'UsedQty')
    ALTER TABLE Coupons ADD UsedQty INT NOT NULL DEFAULT 0;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'FirstOrderOnly')
    ALTER TABLE Coupons ADD FirstOrderOnly BIT NOT NULL DEFAULT 0;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'ApplicableCategory')
    ALTER TABLE Coupons ADD ApplicableCategory NVARCHAR(50) NULL;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'ApplicableProductID')
    ALTER TABLE Coupons ADD ApplicableProductID INT NULL;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'Description')
    ALTER TABLE Coupons ADD Description NVARCHAR(500) DEFAULT '';

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'Terms')
    ALTER TABLE Coupons ADD Terms NVARCHAR(500) DEFAULT '';

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'IsPublic')
    ALTER TABLE Coupons ADD IsPublic BIT NOT NULL DEFAULT 1;

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'UpdatedAt')
    ALTER TABLE Coupons ADD UpdatedAt DATETIME NOT NULL DEFAULT GETDATE();

GO

-- 2. 从旧列名迁移数据到新列名
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'DiscountType')
BEGIN
    UPDATE Coupons SET CouponType = DiscountType WHERE CouponType = 'fixed' AND DiscountType IS NOT NULL AND DiscountType <> 'fixed';
    -- 如果 DiscountType 和 CouponType 不一致，以 DiscountType 为准（首次迁移）
    UPDATE Coupons SET CouponType = DiscountType WHERE CouponType = 'fixed' AND DiscountType <> 'fixed';
END

IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'MinPurchase')
    UPDATE Coupons SET MinSpend = MinPurchase WHERE MinSpend = 0 AND MinPurchase > 0;

IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'StartDate')
    UPDATE Coupons SET ValidFrom = StartDate WHERE ValidFrom <= '2025-01-01' AND StartDate IS NOT NULL;

IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'EndDate')
    UPDATE Coupons SET ValidTo = EndDate WHERE ValidTo >= '2027-01-01' AND EndDate IS NOT NULL;

IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'UsageLimit')
    UPDATE Coupons SET TotalQty = UsageLimit WHERE TotalQty = 0 AND UsageLimit > 0;

IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'Coupons') AND name = 'UsedCount')
    UPDATE Coupons SET UsedQty = UsedCount WHERE UsedQty = 0 AND UsedCount > 0;

GO

PRINT 'Coupons table migration completed successfully.';
