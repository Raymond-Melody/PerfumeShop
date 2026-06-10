$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;")
$conn.Execute("CREATE TABLE OrderItems (OrderItemID INT IDENTITY(1,1) PRIMARY KEY, OrderID INT NOT NULL, ProductID INT NULL, Quantity INT DEFAULT 1, UnitPrice DECIMAL(10,2) DEFAULT 0, CreatedAt DATETIME DEFAULT GETDATE())")
Write-Host "OrderItems table created"
$conn.Close()
