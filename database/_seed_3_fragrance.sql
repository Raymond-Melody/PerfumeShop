SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRAN;

/* ===================== 1. Suppliers ===================== */
DECLARE @sup TABLE(Nm nvarchar(200), P nvarchar(100), Ph nvarchar(60), Cat nvarchar(100));
INSERT INTO @sup VALUES
 (N'格拉斯香精供应商',N'李经理','13910000001',N'香精原料'),
 (N'国际精油贸易',N'王经理','13910000002',N'香精原料'),
 (N'晶莹玻璃器皿厂',N'张经理','13910000003',N'瓶身器具'),
 (N'雅致包装印刷',N'刘经理','13910000004',N'包装材料');
INSERT INTO Suppliers(SupplierName,ContactPerson,Phone,Category,IsActive,CreatedAt)
SELECT Nm,P,Ph,Cat,1,GETDATE() FROM @sup x WHERE NOT EXISTS(SELECT 1 FROM Suppliers s WHERE s.SupplierName=x.Nm);

/* ===================== 2. RawMaterialInventory（香精原料/基香原料） ===================== */
DECLARE @rm TABLE(Cd nvarchar(100), Nm nvarchar(400), Cat nvarchar(40), Stock float, Safety float, Unit nvarchar(40), Price decimal(18,4), SupNm nvarchar(200));
INSERT INTO @rm VALUES
 ('RM-BER',N'佛手柑精油','RAW',5000,500,N'ml',1.20,N'国际精油贸易'),
 ('RM-LEM',N'柠檬精油','RAW',5000,500,N'ml',0.90,N'国际精油贸易'),
 ('RM-ROSE',N'玫瑰精油','RAW',3000,300,N'ml',3.50,N'格拉斯香精供应商'),
 ('RM-JAS',N'茉莉精油','RAW',3000,300,N'ml',3.80,N'格拉斯香精供应商'),
 ('RM-SAN',N'檀香精油','RAW',2000,200,N'ml',5.00,N'格拉斯香精供应商'),
 ('RM-MUSK',N'白麝香香基','RAW',2000,200,N'ml',4.20,N'格拉斯香精供应商'),
 ('RM-VAN',N'香草香基','RAW',2500,250,N'ml',2.60,N'国际精油贸易'),
 ('RM-ALC',N'香水级酒精','BASE',50000,5000,N'ml',0.05,N'国际精油贸易'),
 ('RM-FIX',N'定香剂','BASE',3000,300,N'ml',1.80,N'国际精油贸易');
INSERT INTO RawMaterialInventory(ItemCode,ItemName,CategoryCode,StockQty,SafetyStock,Unit,UnitPrice,WeightedUnitCost,SupplierID,LastPurchaseDate,UpdatedAt)
SELECT x.Cd,x.Nm,x.Cat,x.Stock,x.Safety,x.Unit,x.Price,x.Price,s.SupplierID,GETDATE(),GETDATE()
FROM @rm x LEFT JOIN Suppliers s ON s.SupplierName=x.SupNm
WHERE NOT EXISTS(SELECT 1 FROM RawMaterialInventory r WHERE r.ItemCode=x.Cd);

/* ===================== 3. SupplierPrices ===================== */
INSERT INTO SupplierPrices(SupplierID,ItemCode,ItemName,UnitPrice,MinOrderQty,Unit,PriceType,IsActive,EffectiveDate,CreatedAt)
SELECT r.SupplierID, r.ItemCode, r.ItemName, r.UnitPrice, 100, r.Unit, N'采购价', 1, GETDATE(), GETDATE()
FROM RawMaterialInventory r
WHERE r.SupplierID IS NOT NULL AND NOT EXISTS(SELECT 1 FROM SupplierPrices sp WHERE sp.ItemCode=r.ItemCode AND sp.SupplierID=r.SupplierID);

/* ===================== 4. BaseNotes（基香，关联原料 MaterialID，供成本引擎） ===================== */
DECLARE @bn TABLE(Nm nvarchar(200), Price decimal(18,4), MatCode nvarchar(100));
INSERT INTO @bn VALUES
 (N'佛手柑基香',1.50,'RM-BER'),
 (N'玫瑰基香',4.00,'RM-ROSE'),
 (N'茉莉基香',4.30,'RM-JAS'),
 (N'檀香基香',5.60,'RM-SAN'),
 (N'麝香基香',4.80,'RM-MUSK'),
 (N'香草基香',3.00,'RM-VAN');
INSERT INTO BaseNotes(BaseNoteName,Description,Ingredients,IsActive,UnitPrice,MaterialID)
SELECT x.Nm, x.Nm, x.Nm, 1, x.Price, r.MaterialID
FROM @bn x LEFT JOIN RawMaterialInventory r ON r.ItemCode=x.MatCode
WHERE NOT EXISTS(SELECT 1 FROM BaseNotes b WHERE b.BaseNoteName=x.Nm);

/* ===================== 5. FragranceNotes（前/中/后调，香调主表） ===================== */
DECLARE @fn TABLE(Nm nvarchar(200), Typ nvarchar(40), Price decimal(18,4), IsBase int, Pct int);
INSERT INTO @fn VALUES
 (N'佛手柑',N'前调',8,0,15),(N'柠檬',N'前调',6,0,15),(N'甜橙',N'前调',6,0,12),(N'粉红胡椒',N'前调',10,0,8),
 (N'玫瑰',N'中调',18,0,20),(N'茉莉',N'中调',20,0,18),(N'依兰',N'中调',14,0,12),(N'鼠尾草',N'中调',12,0,10),
 (N'檀香',N'后调',22,1,20),(N'麝香',N'后调',20,1,18),(N'香草',N'后调',15,1,15),(N'广藿香',N'后调',16,1,12);
INSERT INTO FragranceNotes(NoteName,NoteType,PriceAddition,IsActive,IsBaseNote,RecommendedPercentage,Description)
SELECT x.Nm,x.Typ,x.Price,1,x.IsBase,x.Pct,x.Nm+N'香调'
FROM @fn x WHERE NOT EXISTS(SELECT 1 FROM FragranceNotes f WHERE f.NoteName=x.Nm);

/* ===================== 6. NoteInventory（每个香调一条库存，含加权成本） ===================== */
INSERT INTO NoteInventory(NoteID,StockQuantity,MinStockLevel,WeightedUnitCost,LastRestockDate,UpdatedAt)
SELECT f.NoteID, 2000, 200, CAST(f.PriceAddition*0.35 AS decimal(18,4)), GETDATE(), GETDATE()
FROM FragranceNotes f WHERE NOT EXISTS(SELECT 1 FROM NoteInventory ni WHERE ni.NoteID=f.NoteID);

/* ===================== 7. AccordProductions（香调生产记录，2 条已完成） ===================== */
INSERT INTO AccordProductions(NoteID,NoteName,BatchNo,PlannedQty,ActualQty,Status,WorkCenter,StartedAt,CompletedAt,CreatedAt,UpdatedAt)
SELECT TOP 2 f.NoteID, f.NoteName, 'ACC-'+RIGHT('000'+CAST(f.NoteID AS varchar(10)),3), 500, 500, 'Completed','SEMI',DATEADD(day,-3,GETDATE()),DATEADD(day,-2,GETDATE()),DATEADD(day,-3,GETDATE()),GETDATE()
FROM FragranceNotes f WHERE f.NoteType=N'中调' AND NOT EXISTS(SELECT 1 FROM AccordProductions ap WHERE ap.NoteID=f.NoteID)
ORDER BY f.NoteID;

COMMIT;
PRINT 'SEED_3_FRAGRANCE_OK';
SELECT 'Suppliers' t,COUNT(*) n FROM Suppliers UNION ALL SELECT 'RawMaterialInventory',COUNT(*) FROM RawMaterialInventory
UNION ALL SELECT 'SupplierPrices',COUNT(*) FROM SupplierPrices UNION ALL SELECT 'BaseNotes',COUNT(*) FROM BaseNotes
UNION ALL SELECT 'FragranceNotes',COUNT(*) FROM FragranceNotes UNION ALL SELECT 'NoteInventory',COUNT(*) FROM NoteInventory
UNION ALL SELECT 'AccordProductions',COUNT(*) FROM AccordProductions;
