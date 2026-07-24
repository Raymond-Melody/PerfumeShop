SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRAN;

/* ===================== 1. BottleStyles ===================== */
DECLARE @bs TABLE(Nm nvarchar(100), Add1 decimal(18,4), Stock decimal(18,4), Cost decimal(18,4), Cap int);
INSERT INTO @bs VALUES
 (N'经典圆瓶',0,500,2.50,50),(N'花语瓶',15,400,4.00,50),(N'方形瓶',10,400,3.50,50),(N'奢华瓶',35,300,8.00,100);
INSERT INTO BottleStyles(BottleName,Description,IsActive,PriceAddition,StockQty,SafetyStock,UnitCost,UnitPrice,CapacityML,UpdatedAt)
SELECT Nm,Nm,1,Add1,Stock,50,Cost,Cost*1.5,Cap,GETDATE() FROM @bs x WHERE NOT EXISTS(SELECT 1 FROM BottleStyles b WHERE b.BottleName=x.Nm);

/* ===================== 2. PackagingInventory ===================== */
DECLARE @pk TABLE(Cd nvarchar(100), Nm nvarchar(200), Stock decimal(18,4), Price decimal(18,4));
INSERT INTO @pk VALUES
 ('PKG-BOX',N'礼盒包装',1000,3.00),('PKG-BAG',N'手提袋',1000,1.00),('PKG-CARD',N'香卡',2000,0.50),('PKG-SPRAY',N'喷头',2000,0.80);
INSERT INTO PackagingInventory(ItemCode,ItemName,StockQty,SafetyStock,Unit,UnitPrice,UpdatedAt)
SELECT Cd,Nm,Stock,100,N'个',Price,GETDATE() FROM @pk x WHERE NOT EXISTS(SELECT 1 FROM PackagingInventory p WHERE p.ItemCode=x.Cd);

/* ===================== 3. Products（standard 6 + kol 2 + custom 1） ===================== */
DECLARE @prod TABLE(Nm nvarchar(200), Typ nvarchar(100), Price decimal(18,4), Cat nvarchar(100), Top1 nvarchar(200), Mid1 nvarchar(200), Base1 nvarchar(200), Bottle nvarchar(100), KOL int, Review nvarchar(40));
INSERT INTO @prod VALUES
 (N'晨曦柑橘淡香水',N'standard',199,N'柑橘调',N'佛手柑',N'玫瑰',N'檀香',N'经典圆瓶',NULL,'Approved'),
 (N'玫瑰绮梦香水',N'standard',299,N'花香调',N'柠檬',N'玫瑰',N'麝香',N'花语瓶',NULL,'Approved'),
 (N'茉莉倾城香水',N'standard',289,N'花香调',N'甜橙',N'茉莉',N'香草',N'花语瓶',NULL,'Approved'),
 (N'沉香木语香水',N'standard',359,N'木质调',N'佛手柑',N'依兰',N'檀香',N'方形瓶',NULL,'Approved'),
 (N'东方琥珀香水',N'standard',459,N'东方调',N'粉红胡椒',N'玫瑰',N'广藿香',N'奢华瓶',NULL,'Approved'),
 (N'清新之泉香水',N'standard',179,N'清新调',N'柠檬',N'鼠尾草',N'麝香',N'经典圆瓶',NULL,'Approved'),
 (N'星光限定联名香',N'kol',599,N'东方调',N'粉红胡椒',N'茉莉',N'香草',N'奢华瓶',1,'Approved'),
 (N'栀子花语联名香',N'kol',529,N'花香调',N'甜橙',N'茉莉',N'檀香',N'花语瓶',2,'Approved'),
 (N'专属定制调香',N'custom',399,N'花香调',N'佛手柑',N'玫瑰',N'檀香',N'经典圆瓶',NULL,'Approved');
INSERT INTO Products(ProductName,ProductType,BasePrice,Category,Description,ImageURL,IsActive,ReviewStatus,Weight,Volume,KOLID,Engravable,EngravingPrice,CreatedAt)
SELECT x.Nm,x.Typ,x.Price,x.Cat,x.Nm+N'，'+x.Cat+N'，前调'+x.Top1+N'/中调'+x.Mid1+N'/后调'+x.Base1,'/images/default-product.svg',1,x.Review,0.35,50,x.KOL,1,20,GETDATE()
FROM @prod x WHERE NOT EXISTS(SELECT 1 FROM Products p WHERE p.ProductName=x.Nm);

/* ===================== 4. ProductImages（主图） ===================== */
INSERT INTO ProductImages(ProductID,ImageURL,IsPrimary,SortOrder,CreatedAt)
SELECT p.ProductID,'/images/default-product.svg',1,0,GETDATE()
FROM Products p WHERE NOT EXISTS(SELECT 1 FROM ProductImages i WHERE i.ProductID=p.ProductID);

/* ===================== 5. ProductVolumePrices（每产品 × 每容量） ===================== */
INSERT INTO ProductVolumePrices(ProductID,VolumeID,Price)
SELECT p.ProductID, v.VolumeID, CAST(p.BasePrice*v.PriceMultiplier AS decimal(18,2))
FROM Products p CROSS JOIN Volumes v
WHERE NOT EXISTS(SELECT 1 FROM ProductVolumePrices pv WHERE pv.ProductID=p.ProductID AND pv.VolumeID=v.VolumeID);

/* ===================== 6. ProductNotes（产品↔香调，前/中/后各一） ===================== */
INSERT INTO ProductNotes(ProductID,NoteID)
SELECT p.ProductID, f.NoteID
FROM @prod x
JOIN Products p ON p.ProductName=x.Nm
JOIN FragranceNotes f ON f.NoteName IN (x.Top1,x.Mid1,x.Base1)
WHERE NOT EXISTS(SELECT 1 FROM ProductNotes pn WHERE pn.ProductID=p.ProductID AND pn.NoteID=f.NoteID);

/* ===================== 7. ProductNoteRatios（配比，前30/中40/后30） ===================== */
INSERT INTO ProductNoteRatios(ProductID,NoteID,NoteType,Percentage)
SELECT p.ProductID, f.NoteID, f.NoteType,
       CASE f.NoteType WHEN N'前调' THEN 30 WHEN N'中调' THEN 40 ELSE 30 END
FROM @prod x
JOIN Products p ON p.ProductName=x.Nm
JOIN FragranceNotes f ON f.NoteName IN (x.Top1,x.Mid1,x.Base1)
WHERE NOT EXISTS(SELECT 1 FROM ProductNoteRatios r WHERE r.ProductID=p.ProductID AND r.NoteID=f.NoteID);

/* ===================== 8. ProductBottleStyles（产品↔瓶身） ===================== */
INSERT INTO ProductBottleStyles(ProductID,BottleID,CustomPrice)
SELECT p.ProductID, b.BottleID, NULL
FROM @prod x
JOIN Products p ON p.ProductName=x.Nm
JOIN BottleStyles b ON b.BottleName=x.Bottle
WHERE NOT EXISTS(SELECT 1 FROM ProductBottleStyles pb WHERE pb.ProductID=p.ProductID AND pb.BottleID=b.BottleID);

/* ===================== 9. ProductInventory（standard 成品库存；发货扣此） ===================== */
INSERT INTO ProductInventory(ProductID,StockQty,SafetyStock,StockType,UnitCost,UpdatedAt)
SELECT p.ProductID, 100, 10, N'standard', CAST(p.BasePrice*0.4 AS decimal(18,2)), GETDATE()
FROM Products p WHERE p.ProductType='standard' AND NOT EXISTS(SELECT 1 FROM ProductInventory pi WHERE pi.ProductID=p.ProductID);

/* ===================== 10. Recipes + RecipeProducts(Published) + RecipeProductNotes（custom/kol） ===================== */
INSERT INTO Recipes(RecipeName,RecipeCode,ProductType,ReviewStatus,IsActive,CreatedBy,CreatedAt,UpdatedAt)
SELECT x.Nm+N'配方','RCP-'+CAST(p.ProductID AS varchar(10)),x.Typ,'Approved',1,N'tech_manager',GETDATE(),GETDATE()
FROM @prod x JOIN Products p ON p.ProductName=x.Nm
WHERE x.Typ IN ('custom','kol') AND NOT EXISTS(SELECT 1 FROM Recipes r WHERE r.RecipeCode='RCP-'+CAST(p.ProductID AS varchar(10)));

INSERT INTO RecipeProducts(ProductID,RecipeID,BatchSize,Status,PublishedAt,PublishedBy,CreatedAt)
SELECT p.ProductID, r.RecipeID, 100, 'Published', GETDATE(), N'tech_manager', GETDATE()
FROM @prod x JOIN Products p ON p.ProductName=x.Nm
JOIN Recipes r ON r.RecipeCode='RCP-'+CAST(p.ProductID AS varchar(10))
WHERE x.Typ IN ('custom','kol') AND NOT EXISTS(SELECT 1 FROM RecipeProducts rp WHERE rp.ProductID=p.ProductID AND rp.Status='Published');

INSERT INTO RecipeProductNotes(ProductRecipeID,NoteID,NoteName,Percentage,PlannedQty)
SELECT rp.ProductRecipeID, f.NoteID, f.NoteName,
       CASE f.NoteType WHEN N'前调' THEN 30 WHEN N'中调' THEN 40 ELSE 30 END,
       CASE f.NoteType WHEN N'前调' THEN 30 WHEN N'中调' THEN 40 ELSE 30 END
FROM @prod x
JOIN Products p ON p.ProductName=x.Nm
JOIN RecipeProducts rp ON rp.ProductID=p.ProductID AND rp.Status='Published'
JOIN FragranceNotes f ON f.NoteName IN (x.Top1,x.Mid1,x.Base1)
WHERE x.Typ IN ('custom','kol') AND NOT EXISTS(SELECT 1 FROM RecipeProductNotes rpn WHERE rpn.ProductRecipeID=rp.ProductRecipeID AND rpn.NoteID=f.NoteID);

COMMIT;
PRINT 'SEED_4_PRODUCTS_OK';
SELECT 'BottleStyles' t,COUNT(*) n FROM BottleStyles UNION ALL SELECT 'PackagingInventory',COUNT(*) FROM PackagingInventory
UNION ALL SELECT 'Products',COUNT(*) FROM Products UNION ALL SELECT 'ProductImages',COUNT(*) FROM ProductImages
UNION ALL SELECT 'ProductVolumePrices',COUNT(*) FROM ProductVolumePrices UNION ALL SELECT 'ProductNotes',COUNT(*) FROM ProductNotes
UNION ALL SELECT 'ProductNoteRatios',COUNT(*) FROM ProductNoteRatios UNION ALL SELECT 'ProductBottleStyles',COUNT(*) FROM ProductBottleStyles
UNION ALL SELECT 'ProductInventory',COUNT(*) FROM ProductInventory UNION ALL SELECT 'Recipes',COUNT(*) FROM Recipes
UNION ALL SELECT 'RecipeProducts',COUNT(*) FROM RecipeProducts UNION ALL SELECT 'RecipeProductNotes',COUNT(*) FROM RecipeProductNotes;
