-- ============================================
-- V20 费用分摊修复 - 数据库结构迁移
-- 用途：给 Products 增加 Weight/Volume 列（运费按规则自动计算）、补 SiteSettings 默认值
-- 日期：2026-07-17
-- 环境：SQL Server 2017+
-- ============================================

-- 1. Products 增加 Weight 列（单位 kg，含包装）
IF COL_LENGTH('Products','Weight') IS NULL
BEGIN
    ALTER TABLE Products ADD Weight DECIMAL(9,3) NULL;
    PRINT '[OK] Products.Weight added';
END
ELSE
    PRINT '[SKIP] Products.Weight already exists';

-- 2. Products 增加 Volume 列（包装体积，单位 cm³）
IF COL_LENGTH('Products','Volume') IS NULL
BEGIN
    ALTER TABLE Products ADD Volume DECIMAL(12,3) NULL;
    PRINT '[OK] Products.Volume added';
END
ELSE
    PRINT '[SKIP] Products.Volume already exists';

-- 3. 给存量商品设置默认重量（0.5 kg / 500g），使运费计算有可用值
UPDATE Products SET Weight = 0.5 WHERE Weight IS NULL;
PRINT '[OK] Products.Weight defaults set to 0.5';

-- 4. 给存量商品设置默认体积（750 cm³ / 约一瓶香水包装）
UPDATE Products SET Volume = 750 WHERE Volume IS NULL;
PRINT '[OK] Products.Volume defaults set to 750';

-- 5. SiteSettings 补运费分摊默认配置（不存在才插入）
IF NOT EXISTS (SELECT 1 FROM SiteSettings WHERE SettingKey = 'ShippingDefaultUnitWeight')
    INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('ShippingDefaultUnitWeight', '0.5');
IF NOT EXISTS (SELECT 1 FROM SiteSettings WHERE SettingKey = 'ShippingDefaultUnitVolume')
    INSERT INTO SiteSettings (SettingKey, SettingValue) VALUES ('ShippingDefaultUnitVolume', '750');
PRINT '[OK] SiteSettings defaults seeded';

PRINT 'V20 费用分摊修复 - 数据库迁移完成';
