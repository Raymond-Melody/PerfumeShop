<%
' ============================================
' V14.6 采购订单 - 数据库Schema迁移
' 从 purchase_orders.asp 提取
' 确保必要的表结构和字段存在
' ============================================

' ========== 确保必要字段存在 ==========
On Error Resume Next
conn.Execute "SELECT OrderType FROM PurchaseOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PurchaseOrders ADD OrderType NVARCHAR(20) DEFAULT 'RawMaterial'"
conn.Execute "SELECT CategoryCode FROM PurchaseOrders WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE PurchaseOrders ADD CategoryCode NVARCHAR(20) DEFAULT 'RAW'"
' ========== 分类-类型映射表 ==========
conn.Execute "SELECT TOP 1 1 FROM PurchaseCategories"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PurchaseCategories (CategoryID INT IDENTITY(1,1) PRIMARY KEY, CategoryCode NVARCHAR(20) NOT NULL UNIQUE, CategoryName NVARCHAR(50) NOT NULL, IconClass NVARCHAR(50), SortOrder INT DEFAULT 0, IsActive BIT DEFAULT 1)"
    ' 插入7个默认分类（与现有CategoryCode一致）
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('RAW','原材料','fas fa-flask',1)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('BASE','基香原料','fas fa-leaf',2)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('PACK','包装材料','fas fa-box',3)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('BOTTLE','瓶子包装','fas fa-wine-bottle',4)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('PRINTING','印刷品','fas fa-print',5)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('SPRAYHEAD','喷头配件','fas fa-spray-can',6)"
    conn.Execute "INSERT INTO PurchaseCategories (CategoryCode, CategoryName, IconClass, SortOrder) VALUES ('MARKET','营销物料','fas fa-ad',7)"
End If
conn.Execute "SELECT TOP 1 1 FROM PurchaseCategoryTypes"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PurchaseCategoryTypes (MapID INT IDENTITY(1,1) PRIMARY KEY, CategoryCode NVARCHAR(20) NOT NULL, OrderType NVARCHAR(30) NOT NULL, IsDefault BIT DEFAULT 0, FOREIGN KEY (CategoryCode) REFERENCES PurchaseCategories(CategoryCode))"
    ' 插入默认映射关系（与现有代码保持一致）
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('RAW','RawMaterial',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('BASE','RawMaterial',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('PACK','Packaging',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('BOTTLE','Bottle',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('PRINTING','Printing',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('SPRAYHEAD','SprayHead',1)"
    conn.Execute "INSERT INTO PurchaseCategoryTypes (CategoryCode, OrderType, IsDefault) VALUES ('MARKET','Packaging',1)"
End If
' V12: 状态变更日志表
conn.Execute "SELECT TOP 1 1 FROM PurchaseOrderStatusLog"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PurchaseOrderStatusLog (LogID INT IDENTITY(1,1) PRIMARY KEY, PurchaseID INT NOT NULL, FromStatus NVARCHAR(30), ToStatus NVARCHAR(30) NOT NULL, ChangedBy NVARCHAR(50), ChangedAt DATETIME DEFAULT GETDATE(), Remarks NVARCHAR(200))"
End If
On Error GoTo 0
%>
