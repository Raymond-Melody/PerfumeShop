<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<%
Call OpenConnection()

' V8：自动创建 PackagingInventory 和 BottleStyles 库存字段
On Error Resume Next
conn.Execute "SELECT TOP 1 1 FROM PackagingInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PackagingInventory (PackagingID INT IDENTITY(1,1) PRIMARY KEY, PackagingName NVARCHAR(200) NOT NULL, ItemCode NVARCHAR(100), SupplierID INT, PackagingType NVARCHAR(50), Unit NVARCHAR(30), StockQty DECIMAL(19,4) DEFAULT 0, SafetyStock DECIMAL(19,4) DEFAULT 0, UnitPrice DECIMAL(19,4) DEFAULT 0, LastPurchaseDate DATETIME, Status NVARCHAR(20) DEFAULT 'Active', UpdatedAt DATETIME DEFAULT GETDATE())"
End If
conn.Execute "SELECT StockQty FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD StockQty DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT SafetyStock FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD SafetyStock DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT UnitPrice FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD UnitPrice DECIMAL(19,4) DEFAULT 0"
conn.Execute "SELECT LastPurchaseDate FROM BottleStyles WHERE 1=0"
If Err.Number <> 0 Then Err.Clear : conn.Execute "ALTER TABLE BottleStyles ADD LastPurchaseDate DATETIME"
' V10: 自动创建收货与批次追溯表
conn.Execute "SELECT TOP 1 1 FROM PurchaseReceipts"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE PurchaseReceipts (ReceiptID INT IDENTITY(1,1) PRIMARY KEY, PurchaseID INT, ReceiptNo NVARCHAR(100), SupplierID INT, ReceivedBy NVARCHAR(100), ReceiptDate DATETIME, Status NVARCHAR(30) DEFAULT 'Complete', TotalReceivedQty FLOAT DEFAULT 0, Notes NVARCHAR(500), CreatedAt DATETIME DEFAULT GETDATE())"
conn.Execute "SELECT TOP 1 1 FROM PurchaseReceiptDetails"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE PurchaseReceiptDetails (ReceiptDetailID INT IDENTITY(1,1) PRIMARY KEY, ReceiptID INT, PurchaseDetailID INT, ReceivedQty FLOAT DEFAULT 0, AcceptedQty FLOAT DEFAULT 0, RejectedQty FLOAT DEFAULT 0, RejectReason NVARCHAR(500), UnitPrice DECIMAL(19,4) DEFAULT 0)"
conn.Execute "SELECT TOP 1 1 FROM PurchaseBatches"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE PurchaseBatches (BatchID INT IDENTITY(1,1) PRIMARY KEY, PurchaseDetailID INT, PurchaseID INT, BatchNo NVARCHAR(100), ItemType NVARCHAR(30), ItemCode NVARCHAR(50), ItemName NVARCHAR(200), UnitPrice DECIMAL(19,4) DEFAULT 0, Quantity FLOAT DEFAULT 0, ReceivedQty FLOAT DEFAULT 0, RemainingQty FLOAT DEFAULT 0, ReceivedDate DATETIME, SupplierID INT, CreatedAt DATETIME DEFAULT GETDATE())"
conn.Execute "SELECT TOP 1 1 FROM InventoryBatches"
If Err.Number <> 0 Then Err.Clear : conn.Execute "CREATE TABLE InventoryBatches (BatchID INT IDENTITY(1,1) PRIMARY KEY, ItemType NVARCHAR(30), ItemID INT, ItemCode NVARCHAR(50), ItemName NVARCHAR(200), BatchNo NVARCHAR(100), UnitCost DECIMAL(19,4) DEFAULT 0, StockQty FLOAT DEFAULT 0, UpdatedAt DATETIME, CreatedAt DATETIME DEFAULT GETDATE())"
On Error GoTo 0

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then
        SafeNum = 0
    Else
        On Error Resume Next
        SafeNum = val * 1.0
        If Err.Number <> 0 Then SafeNum = 0
        On Error GoTo 0
    End If
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = ""
    Else
        SafeSQL = Replace(str, "'", "''")
    End If
End Function

' ========== 状态常量和映射 ==========
Dim action, purchaseId, msg, msgType
action = Request.QueryString("action")
purchaseId = Request.QueryString("purchase_id")
msg = Request.QueryString("msg")
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

' ========== POST 处理：提交收货 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim postAction
    postAction = Request.Form("action")
    
    If postAction = "create_receipt" Then
        ' CSRF验证
        If Not ValidateCSRFToken() Then
            msg = "安全令牌验证失败，请刷新页面后重试"
            msgType = "error"
            Call CloseConnection()
            Response.End
        End If
        
        Dim recPurchaseID, recSupplierID, recItemCount, recNotes, recDetailCount
        recPurchaseID = SafeNum(Request.Form("purchase_id"))
        recSupplierID = SafeNum(Request.Form("supplier_id"))
        recDetailCount = SafeNum(Request.Form("detail_count"))
        recNotes = Trim(Request.Form("notes"))
        
        ' V8：读取采购订单类型
        Dim recOrderType : recOrderType = "RawMaterial"
        On Error Resume Next
        Dim rsOT : Set rsOT = conn.Execute("SELECT OrderType FROM PurchaseOrders WHERE PurchaseID=" & recPurchaseID)
        If Not rsOT Is Nothing Then If Not rsOT.EOF Then recOrderType = rsOT("OrderType") & ""
        If rsOT Is Nothing Then Err.Clear
        If Not rsOT Is Nothing Then rsOT.Close : Set rsOT = Nothing
        On Error GoTo 0
        If recOrderType = "" Then recOrderType = "RawMaterial"
        
        If recPurchaseID > 0 And recDetailCount > 0 Then
            ' 生成收货单号
            Dim recNo, recReceiptID
            recNo = "RE" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
            
            ' 计算收货总量
            Dim recTotalQty, recAllAccepted
            recTotalQty = 0
            recAllAccepted = True
            
            Dim ri
            For ri = 1 To recDetailCount
                Dim rAcceptedQty
                rAcceptedQty = SafeNum(Request.Form("accepted_qty_" & ri))
                recTotalQty = recTotalQty + rAcceptedQty
                If rAcceptedQty <= 0 Then recAllAccepted = False
            Next
            
            ' 确定收货状态（PurchaseReceipts使用Complete/Partial，PurchaseOrders使用Received/PartialReceived）
            Dim recStatus, poStatus
            If recAllAccepted Then
                recStatus = "Complete"
                poStatus = "Received"
            Else
                recStatus = "Partial"
                poStatus = "PartialReceived"
            End If
            
            On Error Resume Next
            Err.Clear
            Call BeginTransaction()
            If Err.Number <> 0 Then Err.Clear
            
            ' 创建收货单
            Dim sqlReceipt
            sqlReceipt = "INSERT INTO PurchaseReceipts (PurchaseID, ReceiptNo, SupplierID, ReceivedBy, ReceiptDate, Status, TotalReceivedQty, Notes, CreatedAt) VALUES (" & _
                recPurchaseID & ", '" & recNo & "', " & recSupplierID & ", '" & SafeSQL(Session("AdminRealName")) & "', GETDATE(), '" & recStatus & "', " & recTotalQty & ", "
            
            If recNotes <> "" Then
                sqlReceipt = sqlReceipt & "'" & SafeSQL(recNotes) & "'"
            Else
                sqlReceipt = sqlReceipt & "Null"
            End If
            sqlReceipt = sqlReceipt & ", GETDATE())"
            
            conn.Execute sqlReceipt
            
            If Err.Number <> 0 Then
                msg = "创建收货单失败: " & Err.Description
                msgType = "error"
                Call RollbackTransaction()
                Err.Clear
            Else
                ' 获取收货单ID
                Dim rsRecID
                Set rsRecID = conn.Execute("SELECT SCOPE_IDENTITY()")
                recReceiptID = 0
                If Not rsRecID Is Nothing Then
                    If Not rsRecID.EOF Then
                        recReceiptID = rsRecID(0)
                        If IsNull(recReceiptID) Then recReceiptID = 0
                    End If
                    rsRecID.Close
                End If
                Set rsRecID = Nothing
                
                If recReceiptID > 0 Then
                    Dim rj, rDetailID, rAccepted, rRejected, rReason, rMatName, rMatCode, rMatUnit, rMatPrice, rMatSupplier
                    Dim anyError
                    anyError = False
                    
                    For rj = 1 To recDetailCount
                        rDetailID = SafeNum(Request.Form("detail_id_" & rj))
                        rAccepted = SafeNum(Request.Form("accepted_qty_" & rj))
                        rRejected = SafeNum(Request.Form("rejected_qty_" & rj))
                        rReason = Trim(Request.Form("reject_reason_" & rj))
                        rMatName = Trim(Request.Form("item_name_" & rj))
                        rMatCode = Trim(Request.Form("item_code_" & rj))
                        rMatUnit = Trim(Request.Form("unit_" & rj))
                        rMatPrice = SafeNum(Request.Form("unit_price_" & rj))
                        rMatSupplier = SafeNum(Request.Form("supplier_id_" & rj))
                        
                        ' 写入收货明细
                        Dim sqlRecDetail
                        sqlRecDetail = "INSERT INTO PurchaseReceiptDetails (ReceiptID, PurchaseDetailID, ReceivedQty, AcceptedQty, RejectedQty, RejectReason, UnitPrice) VALUES (" & _
                            recReceiptID & ", " & rDetailID & ", " & (rAccepted + rRejected) & ", " & rAccepted & ", " & rRejected & ", "
                        If rReason <> "" Then
                            sqlRecDetail = sqlRecDetail & "'" & SafeSQL(rReason) & "'"
                        Else
                            sqlRecDetail = sqlRecDetail & "Null"
                        End If
                        sqlRecDetail = sqlRecDetail & ", " & rMatPrice & ")"
                        
                        conn.Execute sqlRecDetail
                        
                        If Err.Number <> 0 Then
                            anyError = True
                            Err.Clear
                        Else
                            ' V8：按采购类型联动对应库存表
                            If rAccepted > 0 And rMatName <> "" Then
                                Dim rsMat, matID
                                matID = 0
                                
                                If recOrderType = "Packaging" Then
                                    ' ===== 包装物库存更新 =====
                                    Set rsMat = conn.Execute("SELECT PackagingID FROM PackagingInventory WHERE PackagingName='" & SafeSQL(rMatName) & "'")
                                    If Not rsMat Is Nothing And Not rsMat.EOF Then
                                        matID = rsMat("PackagingID")
                                    End If
                                    If Not rsMat Is Nothing Then rsMat.Close
                                    Set rsMat = Nothing
                                    
                                    If matID > 0 Then
                                        conn.Execute "UPDATE PackagingInventory SET StockQty = StockQty + " & rAccepted & ", UnitPrice = " & rMatPrice & ", LastPurchaseDate = GETDATE(), UpdatedAt = GETDATE() WHERE PackagingID=" & matID
                                        If Err.Number <> 0 Then anyError = True : Err.Clear
                                    Else
                                        conn.Execute "INSERT INTO PackagingInventory (PackagingName, ItemCode, SupplierID, PackagingType, Unit, StockQty, UnitPrice, LastPurchaseDate, UpdatedAt) VALUES ('" & _
                                            SafeSQL(rMatName) & "', " & IIf(rMatCode <> "", "'" & SafeSQL(rMatCode) & "'", "Null") & ", " & rMatSupplier & ", 'General', '" & SafeSQL(rMatUnit) & "', " & rAccepted & ", " & rMatPrice & ", GETDATE(), GETDATE())"
                                        If Err.Number = 0 Then
                                            Set rsMat = conn.Execute("SELECT SCOPE_IDENTITY()")
                                            If Not rsMat Is Nothing And Not rsMat.EOF Then matID = CLng(rsMat(0))
                                            rsMat.Close : Set rsMat = Nothing
                                        Else
                                            anyError = True : Err.Clear
                                        End If
                                    End If
                                    
                                ElseIf recOrderType = "Bottle" Then
                                    ' ===== 瓶子库存更新 =====
                                    Set rsMat = conn.Execute("SELECT BottleID FROM BottleStyles WHERE BottleName='" & SafeSQL(rMatName) & "'")
                                    If Not rsMat Is Nothing And Not rsMat.EOF Then
                                        matID = rsMat("BottleID")
                                    End If
                                    If Not rsMat Is Nothing Then rsMat.Close
                                    Set rsMat = Nothing
                                    
                                    If matID > 0 Then
                                        conn.Execute "UPDATE BottleStyles SET StockQty = ISNULL(StockQty,0) + " & rAccepted & ", UnitPrice = " & rMatPrice & ", LastPurchaseDate = GETDATE() WHERE BottleID=" & matID
                                        If Err.Number <> 0 Then anyError = True : Err.Clear
                                    Else
                                        conn.Execute "INSERT INTO BottleStyles (BottleName, BottleCode, SupplierID, StockQty, UnitPrice, LastPurchaseDate) VALUES ('" & _
                                            SafeSQL(rMatName) & "', " & IIf(rMatCode <> "", "'" & SafeSQL(rMatCode) & "'", "Null") & ", " & rMatSupplier & ", " & rAccepted & ", " & rMatPrice & ", GETDATE())"
                                        If Err.Number = 0 Then
                                            Set rsMat = conn.Execute("SELECT SCOPE_IDENTITY()")
                                            If Not rsMat Is Nothing And Not rsMat.EOF Then matID = CLng(rsMat(0))
                                            rsMat.Close : Set rsMat = Nothing
                                        Else
                                            anyError = True : Err.Clear
                                        End If
                                    End If
                                    
                                ElseIf recOrderType = "Printing" Then
                                    ' ===== 印刷品库存更新 =====
                                    Set rsMat = conn.Execute("SELECT PrintingID FROM PrintingInventory WHERE ItemName='" & SafeSQL(rMatName) & "'")
                                    If Not rsMat Is Nothing And Not rsMat.EOF Then
                                        matID = rsMat("PrintingID")
                                    End If
                                    If Not rsMat Is Nothing Then rsMat.Close
                                    Set rsMat = Nothing
                                    
                                    If matID > 0 Then
                                        conn.Execute "UPDATE PrintingInventory SET StockQty = StockQty + " & rAccepted & ", UnitPrice = " & rMatPrice & ", LastPurchaseDate = GETDATE(), UpdatedAt = GETDATE() WHERE PrintingID=" & matID
                                        If Err.Number <> 0 Then anyError = True : Err.Clear
                                    Else
                                        conn.Execute "INSERT INTO PrintingInventory (ItemName, ItemCode, SupplierID, PrintingType, Unit, StockQty, UnitPrice, LastPurchaseDate, UpdatedAt) VALUES ('" & _
                                            SafeSQL(rMatName) & "', " & IIf(rMatCode <> "", "'" & SafeSQL(rMatCode) & "'", "Null") & ", " & rMatSupplier & ", 'Manual', '" & SafeSQL(rMatUnit) & "', " & rAccepted & ", " & rMatPrice & ", GETDATE(), GETDATE())"
                                        If Err.Number = 0 Then
                                            Set rsMat = conn.Execute("SELECT SCOPE_IDENTITY()")
                                            If Not rsMat Is Nothing And Not rsMat.EOF Then matID = CLng(rsMat(0))
                                            rsMat.Close : Set rsMat = Nothing
                                        Else
                                            anyError = True : Err.Clear
                                        End If
                                    End If
                                    
                                ElseIf recOrderType = "SprayHead" Then
                                    ' ===== 喷头库存更新 =====
                                    Set rsMat = conn.Execute("SELECT SprayHeadID FROM SprayHeadInventory WHERE ItemName='" & SafeSQL(rMatName) & "'")
                                    If Not rsMat Is Nothing And Not rsMat.EOF Then
                                        matID = rsMat("SprayHeadID")
                                    End If
                                    If Not rsMat Is Nothing Then rsMat.Close
                                    Set rsMat = Nothing
                                    
                                    If matID > 0 Then
                                        conn.Execute "UPDATE SprayHeadInventory SET StockQty = StockQty + " & rAccepted & ", UnitPrice = " & rMatPrice & ", LastPurchaseDate = GETDATE(), UpdatedAt = GETDATE() WHERE SprayHeadID=" & matID
                                        If Err.Number <> 0 Then anyError = True : Err.Clear
                                    Else
                                        conn.Execute "INSERT INTO SprayHeadInventory (ItemName, ItemCode, SupplierID, SprayType, Unit, StockQty, UnitPrice, LastPurchaseDate, UpdatedAt) VALUES ('" & _
                                            SafeSQL(rMatName) & "', " & IIf(rMatCode <> "", "'" & SafeSQL(rMatCode) & "'", "Null") & ", " & rMatSupplier & ", 'Mist', '" & SafeSQL(rMatUnit) & "', " & rAccepted & ", " & rMatPrice & ", GETDATE(), GETDATE())"
                                        If Err.Number = 0 Then
                                            Set rsMat = conn.Execute("SELECT SCOPE_IDENTITY()")
                                            If Not rsMat Is Nothing And Not rsMat.EOF Then matID = CLng(rsMat(0))
                                            rsMat.Close : Set rsMat = Nothing
                                        Else
                                            anyError = True : Err.Clear
                                        End If
                                    End If
                                    
                                Else
                                    ' ===== 原料库存更新 =====
                                    Set rsMat = conn.Execute("SELECT MaterialID FROM RawMaterialInventory WHERE ItemName='" & SafeSQL(rMatName) & "'")
                                    If Not rsMat Is Nothing And Not rsMat.EOF Then
                                        matID = rsMat("MaterialID")
                                    End If
                                    If Not rsMat Is Nothing Then rsMat.Close
                                    Set rsMat = Nothing
                                    
                                    If matID > 0 Then
                                        conn.Execute "UPDATE RawMaterialInventory SET StockQty = StockQty + " & rAccepted & ", UnitPrice = " & rMatPrice & ", LastPurchaseDate = GETDATE(), UpdatedAt = GETDATE() WHERE MaterialID=" & matID
                                        If Err.Number <> 0 Then anyError = True : Err.Clear
                                    Else
                                        conn.Execute "INSERT INTO RawMaterialInventory (ItemName, ItemCode, SupplierID, CategoryCode, Unit, StockQty, UnitPrice, LastPurchaseDate, UpdatedAt) VALUES ('" & _
                                            SafeSQL(rMatName) & "', " & IIf(rMatCode <> "", "'" & SafeSQL(rMatCode) & "'", "Null") & ", " & rMatSupplier & ", 'RAW_MATERIAL', '" & SafeSQL(rMatUnit) & "', " & rAccepted & ", " & rMatPrice & ", GETDATE(), GETDATE())"
                                        If Err.Number = 0 Then
                                            Set rsMat = conn.Execute("SELECT SCOPE_IDENTITY()")
                                            If Not rsMat Is Nothing And Not rsMat.EOF Then matID = CLng(rsMat(0))
                                            rsMat.Close : Set rsMat = Nothing
                                        Else
                                            anyError = True : Err.Clear
                                        End If
                                    End If
                                End If
                                
                                ' V10：批次记录与加权成本计算
                                If matID > 0 And Not anyError Then
                                    Dim oldStock, oldCost, newCost
                                    oldStock = 0 : oldCost = 0
                                    
                                    ' 读取当前库存加权成本及库存量（扣除本次入库后的旧值）
                                    On Error Resume Next
                                    If recOrderType = "RawMaterial" Then
                                        Dim rsWC : Set rsWC = conn.Execute("SELECT ISNULL(StockQty,0) - " & rAccepted & " as OldStock, ISNULL(WeightedUnitCost,0) as OldCost FROM RawMaterialInventory WHERE MaterialID=" & matID)
                                        If Not rsWC Is Nothing And Not rsWC.EOF Then
                                            oldStock = CDbl("0" & rsWC("OldStock")) : oldCost = CDbl("0" & rsWC("OldCost"))
                                            rsWC.Close
                                        End If : Set rsWC = Nothing
                                    ElseIf recOrderType = "Packaging" Then
                                        Dim rsWP : Set rsWP = conn.Execute("SELECT ISNULL(StockQty,0) - " & rAccepted & " as OldStock, ISNULL(WeightedUnitCost,0) as OldCost FROM PackagingInventory WHERE PackagingID=" & matID)
                                        If Not rsWP Is Nothing And Not rsWP.EOF Then
                                            oldStock = CDbl("0" & rsWP("OldStock")) : oldCost = CDbl("0" & rsWP("OldCost"))
                                            rsWP.Close
                                        End If : Set rsWP = Nothing
                                    ElseIf recOrderType = "Bottle" Then
                                        Dim rsWB : Set rsWB = conn.Execute("SELECT ISNULL(StockQty,0) - " & rAccepted & " as OldStock, ISNULL(WeightedUnitCost,0) as OldCost FROM BottleStyles WHERE BottleID=" & matID)
                                        If Not rsWB Is Nothing And Not rsWB.EOF Then
                                            oldStock = CDbl("0" & rsWB("OldStock")) : oldCost = CDbl("0" & rsWB("OldCost"))
                                            rsWB.Close
                                        End If : Set rsWB = Nothing
                                    ElseIf recOrderType = "Printing" Then
                                        Dim rsWPr : Set rsWPr = conn.Execute("SELECT ISNULL(StockQty,0) - " & rAccepted & " as OldStock, ISNULL(WeightedUnitCost,0) as OldCost FROM PrintingInventory WHERE PrintingID=" & matID)
                                        If Not rsWPr Is Nothing And Not rsWPr.EOF Then
                                            oldStock = CDbl("0" & rsWPr("OldStock")) : oldCost = CDbl("0" & rsWPr("OldCost"))
                                            rsWPr.Close
                                        End If : Set rsWPr = Nothing
                                    ElseIf recOrderType = "SprayHead" Then
                                        Dim rsWS : Set rsWS = conn.Execute("SELECT ISNULL(StockQty,0) - " & rAccepted & " as OldStock, ISNULL(WeightedUnitCost,0) as OldCost FROM SprayHeadInventory WHERE SprayHeadID=" & matID)
                                        If Not rsWS Is Nothing And Not rsWS.EOF Then
                                            oldStock = CDbl("0" & rsWS("OldStock")) : oldCost = CDbl("0" & rsWS("OldCost"))
                                            rsWS.Close
                                        End If : Set rsWS = Nothing
                                    End If
                                    If Err.Number <> 0 Then Err.Clear
                                    On Error GoTo 0
                                    
                                    ' 计算加权平均成本
                                    If oldStock < 0 Then oldStock = 0
                                    If (oldStock + rAccepted) > 0 Then
                                        newCost = (oldStock * oldCost + rAccepted * rMatPrice) / (oldStock + rAccepted)
                                    Else
                                        newCost = rMatPrice
                                    End If
                                    
                                    ' 写入 PurchaseBatches
                                    conn.Execute "INSERT INTO PurchaseBatches (PurchaseDetailID, PurchaseID, BatchNo, ItemType, ItemCode, ItemName, UnitPrice, Quantity, ReceivedQty, RemainingQty, ReceivedDate, SupplierID) VALUES (" & _
                                        rDetailID & ", " & recPurchaseID & ", '" & recNo & "-" & rj & "', '" & recOrderType & "', '" & SafeSQL(rMatCode) & "', '" & SafeSQL(rMatName) & "', " & rMatPrice & ", " & rAccepted & ", " & rAccepted & ", " & rAccepted & ", GETDATE(), " & rMatSupplier & ")"
                                    If Err.Number <> 0 Then anyError = True : Err.Clear
                                    
                                    ' 写入 InventoryBatches（批次实际采购成本，非加权平均）
                                    ' UnitCost 使用实际采购单价 rMatPrice，保留批次差异化成本
                                    ' 加权平均成本 newCost 存储于库存主表 WeightedUnitCost
                                    If Not anyError Then
                                        conn.Execute "INSERT INTO InventoryBatches (ItemType, ItemID, ItemCode, ItemName, BatchNo, UnitCost, StockQty, UpdatedAt, CreatedAt) VALUES ('" & _
                                            recOrderType & "', " & matID & ", '" & SafeSQL(rMatCode) & "', '" & SafeSQL(rMatName) & "', '" & recNo & "-" & rj & "', " & rMatPrice & ", " & rAccepted & ", GETDATE(), GETDATE())"
                                        If Err.Number <> 0 Then anyError = True : Err.Clear
                                    End If
                                    
                                    ' 更新库存表的加权成本
                                    If Not anyError Then
                                        If recOrderType = "RawMaterial" Then
                                            conn.Execute "UPDATE RawMaterialInventory SET WeightedUnitCost = " & newCost & " WHERE MaterialID=" & matID
                                        ElseIf recOrderType = "Packaging" Then
                                            conn.Execute "UPDATE PackagingInventory SET WeightedUnitCost = " & newCost & " WHERE PackagingID=" & matID
                                        ElseIf recOrderType = "Bottle" Then
                                            conn.Execute "UPDATE BottleStyles SET WeightedUnitCost = " & newCost & " WHERE BottleID=" & matID
                                        ElseIf recOrderType = "Printing" Then
                                            conn.Execute "UPDATE PrintingInventory SET WeightedUnitCost = " & newCost & " WHERE PrintingID=" & matID
                                        ElseIf recOrderType = "SprayHead" Then
                                            conn.Execute "UPDATE SprayHeadInventory SET WeightedUnitCost = " & newCost & " WHERE SprayHeadID=" & matID
                                        End If
                                        If Err.Number <> 0 Then anyError = True : Err.Clear
                                    End If
                                End If
                                
                                ' 记录库存流水
                                If matID > 0 And Not anyError Then
                                    conn.Execute "INSERT INTO InventoryTransactions (NoteID, MaterialID, Quantity, TransactionType, TransactionDirection, UnitCost, Notes, CreatedBy, CreatedAt) VALUES (" & _
                                        "0, " & matID & ", " & rAccepted & ", '采购入库', 'IN', " & rMatPrice & ", '收货单" & recNo & "', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
                                End If
                            End If
                        End If
                    Next
                    
                    If Not anyError Then
                        ' 更新采购订单明细的收货数量
                        For rj = 1 To recDetailCount
                            rDetailID = SafeNum(Request.Form("detail_id_" & rj))
                            rAccepted = SafeNum(Request.Form("accepted_qty_" & rj))
                            If rDetailID > 0 Then
                                conn.Execute "UPDATE PurchaseOrderDetails SET ReceivedQty = ReceivedQty + " & rAccepted & " WHERE DetailID=" & rDetailID
                                If Err.Number <> 0 Then
                                    anyError = True
                                    Err.Clear
                                End If
                            End If
                        Next
                        
                        If Not anyError Then
                            ' 更新采购订单状态
                            conn.Execute "UPDATE PurchaseOrders SET Status='" & SafeSQL(poStatus) & "', UpdatedAt= GETDATE() WHERE PurchaseID=" & recPurchaseID
                            If Err.Number <> 0 Then
                                anyError = True
                                Err.Clear
                            End If
                        End If
                    End If
                    
                    If Not anyError Then
                        Call CommitTransaction()
                        Response.Redirect "receiving.asp?msg=收货成功！收货单号：" & recNo
                        Response.End
                    Else
                        Call RollbackTransaction()
                        msg = "收货失败，数据已回滚，请重试"
                        msgType = "error"
                    End If
                Else
                    Call RollbackTransaction()
                    msg = "获取收货单ID失败"
                    msgType = "error"
                End If
            End If
            On Error GoTo 0
        Else
            msg = "参数错误"
            msgType = "error"
        End If
    End If
End If

' 获取收货统计
Dim statsPending, statsToday, statsTotalReceipts
statsPending = 0
statsToday = 0
statsTotalReceipts = 0
On Error Resume Next
Dim rsStats
Set rsStats = conn.Execute("SELECT COUNT(*) FROM PurchaseReceipts")
If Not rsStats Is Nothing And Not rsStats.EOF Then statsTotalReceipts = CLng(rsStats(0))
rsStats.Close
Set rsStats = Nothing

Set rsStats = conn.Execute("SELECT COUNT(*) FROM PurchaseReceipts WHERE Status='Partial'")
If Not rsStats Is Nothing And Not rsStats.EOF Then statsPending = CLng(rsStats(0))
rsStats.Close
Set rsStats = Nothing

Set rsStats = conn.Execute("SELECT COUNT(*) FROM PurchaseReceipts WHERE CAST(ReceiptDate AS DATE)=CAST(GETDATE() AS DATE)")
If Not rsStats Is Nothing And Not rsStats.EOF Then statsToday = CLng(rsStats(0))
rsStats.Close
Set rsStats = Nothing
On Error GoTo 0

' ========== 生成采购单号辅助函数 ==========
Function IIF(cond, tVal, fVal)
    If cond Then IIF = tVal Else IIF = fVal
End Function
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>采购收货管理 - 采购管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        /* 暗色主题基础 */
        body {
            background: #1a1a2e;
            color: #e0e0e0;
        }
        
        /* 消息提示 */
        .message { padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; font-weight: 500; }
        .message.success { background: rgba(39,174,96,0.15); color: #27ae60; border: 1px solid rgba(39,174,96,0.3); }
        .message.error { background: rgba(231,76,60,0.15); color: #e74c3c; border: 1px solid rgba(231,76,60,0.3); }
        
        /* 统计卡片 */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
            margin-bottom: 24px;
        }
        .stat-card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.05);
            text-align: center;
        }
        .stat-card .num { font-size: 28px; font-weight: 700; color: #FF9800; display: block; }
        .stat-card .label { font-size: 13px; color: #888; margin-top: 4px; }
        
        /* 卡片 */
        .card {
            background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%);
            border-radius: 10px;
            border: 1px solid rgba(255,255,255,0.05);
            margin-bottom: 20px;
            overflow: hidden;
        }
        .card-header {
            padding: 14px 20px;
            background: rgba(255,152,0,0.06);
            border-bottom: 1px solid rgba(255,255,255,0.05);
            font-weight: 600;
            font-size: 15px;
            color: #fff;
        }
        .card-body { padding: 16px 20px; }
        
        /* 表格 */
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px 12px; background: rgba(255,152,0,0.05); border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 12px; color: #888; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.05); font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.02); }
        
        /* 状态标签 */
        .status { padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .status-draft { background: rgba(158,158,158,0.2); color: #9e9e9e; }
        .status-submitted { background: rgba(255,152,0,0.2); color: #FF9800; }
        .status-ordered { background: rgba(33,150,243,0.2); color: #2196F3; }
        .status-received { background: rgba(76,175,80,0.2); color: #4CAF50; }
        .status-partial { background: rgba(255,193,7,0.2); color: #FFC107; }
        .status-completed { background: rgba(156,39,176,0.2); color: #CE93D8; }
        
        /* 表单 */
        input[type="text"], input[type="number"], textarea, select { width: 100%; padding: 9px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.08); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        input[type="text"]:focus, input[type="number"]:focus, textarea:focus { border-color: #FF9800; outline: none; }
        
        /* 辅助类 */
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }
        .grid-4 { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr; gap: 8px; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .mb-1 { margin-bottom: 8px; }
        .mb-2 { margin-bottom: 16px; }
        .mt-2 { margin-top: 16px; }
        
        .detail-header { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr; gap: 8px; padding: 8px 12px; background: rgba(255,255,255,0.03); font-size: 13px; color: #888; font-weight: 600; margin-bottom: 4px; border-radius: 6px; }
        .detail-row { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr; gap: 8px; padding: 8px 12px; border-bottom: 1px solid rgba(255,255,255,0.05); align-items: center; }
        .detail-row input { width: 100%; padding: 6px 8px; }
        
        .section-title {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 15px;
            color: #fff;
            display: flex;
            align-items: center;
        }
        .section-title i { margin-right: 10px; color: #FF9800; }
        
        .empty-row {
            text-align: center;
            color: #666;
            padding: 40px;
        }
    </style>
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-clipboard-check"></i> 采购收货管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>收货入库</span>
            </div>
        </div>
    
    <% If msg <> "" Then %>
    <div class="message <%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
    <% End If %>
    
    <!-- 统计卡片 -->
    <div class="stats-grid">
        <div class="stat-card"><span class="num"><%=statsTotalReceipts%></span><span class="label">总收货单</span></div>
        <div class="stat-card"><span class="num"><%=statsToday%></span><span class="label">今日收货</span></div>
        <div class="stat-card"><span class="num"><%=statsPending%></span><span class="label">待完成收货</span></div>
    </div>
    
    <%
    ' ========== 收货录入页面 ==========
    If action = "receive" And purchaseId <> "" And IsNumeric(purchaseId) Then
        Dim rsPO
        Set rsPO = conn.Execute("SELECT po.PurchaseID, po.PurchaseNo, po.SupplierID, po.OrderType, po.CategoryCode, CAST(ISNULL(po.TotalAmount,0) AS FLOAT) as TotalAmount, po.Status, po.CreatedBy, po.ExpectedDate, po.Remarks, s.SupplierName, s.ContactPerson FROM PurchaseOrders po LEFT JOIN Suppliers s ON po.SupplierID = s.SupplierID WHERE po.PurchaseID=" & CLng(purchaseId))
        If Not rsPO Is Nothing And Not rsPO.EOF Then
            poStatus = rsPO("Status")
    %>
    <div class="card">
        <div class="card-header">收货录入 - <%=rsPO("PurchaseNo")%></div>
        <div class="card-body">
            <form method="post" id="receiptForm">
                <input type="hidden" name="action" value="create_receipt">
                <%=GetCSRFTokenField()%>
                <input type="hidden" name="purchase_id" value="<%=purchaseId%>">
                <input type="hidden" name="supplier_id" value="<%=rsPO("SupplierID")%>">
                
                <div class="grid-3 mb-2">
                    <div><label class="text-muted">采购单号</label><div style="font-weight:600"><%=rsPO("PurchaseNo")%></div></div>
                    <div><label class="text-muted">采购类型</label><div style="font-weight:600">
                        <%
                            Dim formOT : formOT = rsPO("OrderType") & ""
                            If formOT = "Packaging" Then
                                Response.Write "包装物"
                            ElseIf formOT = "Bottle" Then
                                Response.Write "瓶子"
                            ElseIf formOT = "Printing" Then
                                Response.Write "印刷品"
                            ElseIf formOT = "SprayHead" Then
                                Response.Write "喷头"
                            Else
                                Response.Write "原料"
                            End If
                        %>
                    </div></div>
                    <div><label class="text-muted">供应商</label><div style="font-weight:600"><%=rsPO("SupplierName")%></div></div>
                    <div><label class="text-muted">联系人</label><div><%=rsPO("ContactPerson")%></div></div>
                </div>
                
                <h2>收货明细</h2>
                <div class="detail-header">
                    <span>物料名称</span><span>采购数量</span><span>合格数量</span><span>不合格数量</span>
                </div>
                
                <% 
                Dim rsPOD, detailIdx
                detailIdx = 0
                Set rsPOD = conn.Execute("SELECT DetailID, PurchaseID, ItemName, ItemCode, Specification, Unit, Quantity, CAST(ISNULL(UnitPrice,0) AS FLOAT) as UnitPrice, CAST(ISNULL(TotalPrice,0) AS FLOAT) as TotalPrice, ReceivedQty FROM PurchaseOrderDetails WHERE PurchaseID=" & CLng(purchaseId) & " ORDER BY DetailID")
                If Not rsPOD Is Nothing Then
                    Do While Not rsPOD.EOF
                        detailIdx = detailIdx + 1
                        Dim dimPurchasedQty, dimReceivedQty, dimRemaining
                        dimPurchasedQty = SafeNum(rsPOD("Quantity"))
                        dimReceivedQty = SafeNum(rsPOD("ReceivedQty"))
                        dimRemaining = dimPurchasedQty - dimReceivedQty
                        If dimRemaining < 0 Then dimRemaining = 0
                %>
                <div class="detail-row">
                    <div>
                        <input type="hidden" name="detail_id_<%=detailIdx%>" value="<%=rsPOD("DetailID")%>">
                        <input type="hidden" name="item_name_<%=detailIdx%>" value="<%=Server.HTMLEncode(rsPOD("ItemName"))%>">
                        <input type="hidden" name="item_code_<%=detailIdx%>" value="<%=Server.HTMLEncode(rsPOD("ItemCode") & "")%>">
                        <input type="hidden" name="unit_<%=detailIdx%>" value="<%=Server.HTMLEncode(rsPOD("Unit") & "")%>">
                        <input type="hidden" name="unit_price_<%=detailIdx%>" value="<%=rsPOD("UnitPrice")%>">
                        <input type="hidden" name="supplier_id_<%=detailIdx%>" value="<%=rsPO("SupplierID")%>">
                        <strong><%=Server.HTMLEncode(rsPOD("ItemName"))%></strong>
                        <div class="text-muted" style="font-size:12px">单价: <%=FormatNumber(CDbl("0" & rsPOD("UnitPrice")),2)%> | 剩余: <%=dimRemaining%></div>
                    </div>
                    <div class="text-center"><%=dimPurchasedQty%></div>
                    <div><input type="number" name="accepted_qty_<%=detailIdx%>" value="<%=dimRemaining%>" min="0" max="<%=dimRemaining%>" step="0.01" style="margin-bottom:4px" onchange="updateTotal()"></div>
                    <div><input type="number" name="rejected_qty_<%=detailIdx%>" value="0" min="0" max="<%=dimRemaining%>" step="0.01">
                        <input type="text" name="reject_reason_<%=detailIdx%>" placeholder="不合格原因" style="font-size:12px; margin-top:2px"></div>
                </div>
                <%
                        rsPOD.MoveNext
                    Loop
                    rsPOD.Close
                End If
                Set rsPOD = Nothing
                %>
                <input type="hidden" name="detail_count" value="<%=detailIdx%>">
                
                <div class="mt-2 grid-2">
                    <div>
                        <label class="text-muted">备注</label>
                        <textarea name="notes" rows="3" placeholder="收货备注..."></textarea>
                    </div>
                    <div style="text-align:right; padding-top: 20px;">
                        <span class="text-muted">收货合计: </span>
                        <strong id="totalDisplay" style="font-size:18px; color:#FF9800">0</strong>
                    </div>
                </div>
                
                <div class="mt-2" style="display:flex; gap:12px; justify-content:flex-end">
                    <a href="receiving.asp" class="btn btn-outline">返回列表</a>
                    <button type="submit" class="btn btn-success" onclick="return confirm('确认提交收货？')">确认收货</button>
                </div>
            </form>
        </div>
    </div>
    <script>
    function updateTotal() {
        var total = 0;
        var inputs = document.querySelectorAll('[name^="accepted_qty_"]');
        inputs.forEach(function(inp) { total += parseFloat(inp.value) || 0; });
        document.getElementById('totalDisplay').textContent = total.toFixed(2);
    }
    updateTotal();
    </script>
    <%
        Else
            Response.Write "<div class='message error'>采购订单不存在</div>"
        End If
        If Not rsPO Is Nothing Then rsPO.Close
        Set rsPO = Nothing
    %>
    
    <%
    ' ========== 历史收货记录 ==========
    Else
    %>
    <div class="card">
        <div class="card-header">收货记录</div>
        <div class="card-body" style="overflow-x:auto;">
            <table>
                <thead>
                    <tr>
                        <th>收货单号</th><th>采购单号</th><th>供应商</th><th>收货日期</th><th>收货数量</th><th>批次</th><th>状态</th><th>操作</th>
                    </tr>
                </thead>
                <tbody>
                <%
                Dim sqlReceipts, rsReceipts
                sqlReceipts = "SELECT pr.*, po.PurchaseNo, s.SupplierName, (SELECT COUNT(*) FROM PurchaseBatches WHERE PurchaseID=pr.PurchaseID) AS BatchCount FROM (PurchaseReceipts pr LEFT JOIN PurchaseOrders po ON pr.PurchaseID=po.PurchaseID) LEFT JOIN Suppliers s ON pr.SupplierID=s.SupplierID ORDER BY pr.ReceiptID DESC"
                Set rsReceipts = conn.Execute(sqlReceipts)
                If Not rsReceipts Is Nothing Then
                    Dim recCount
                    recCount = 0
                    Do While Not rsReceipts.EOF And recCount < 50
                        recCount = recCount + 1
                        Dim recStatusLabel, recStatusClass, recSC
                        recSC = UCase(Trim(rsReceipts("Status") & ""))
                        Select Case recSC
                            Case "COMPLETE", "COMPLETED"
                                recStatusLabel = "收货完成"
                                recStatusClass = "status-received"
                            Case "PARTIAL"
                                recStatusLabel = "部分收货"
                                recStatusClass = "status-partial"
                            Case "RECEIVED"
                                recStatusLabel = "已收货"
                                recStatusClass = "status-received"
                            Case "CANCELLED"
                                recStatusLabel = "已取消"
                                recStatusClass = "status-draft"
                            Case "PENDING"
                                recStatusLabel = "待处理"
                                recStatusClass = "status-submitted"
                            Case Else
                                recStatusLabel = rsReceipts("Status") & ""
                                recStatusClass = "status-draft"
                        End Select
                %>
                    <tr>
                        <td><strong><%=rsReceipts("ReceiptNo")%></strong></td>
                        <td><%=rsReceipts("PurchaseNo") & ""%></td>
                        <td><%=rsReceipts("SupplierName") & ""%></td>
                        <td><%=rsReceipts("ReceiptDate") & ""%></td>
                        <td><%=rsReceipts("TotalReceivedQty")%></td>
                        <td><span class="status <%=recStatusClass%>"><%=Server.HTMLEncode(recStatusLabel)%></span></td>
                        <td>
                            <% If SafeNum(rsReceipts("BatchCount")) > 0 Then %>
                            <a href="javascript:void(0)" class="btn btn-outline btn-sm" onclick="showBatchDetail(<%=rsReceipts("ReceiptID")%>)"><i class="fas fa-layer-group"></i> 批次 (<%=rsReceipts("BatchCount")%>)</a>
                            <% Else %>
                            <span class="text-muted" style="font-size:12px;">无批次</span>
                            <% End If %>
                        </td>
                    </tr>
                <%
                        rsReceipts.MoveNext
                    Loop
                    rsReceipts.Close
                    If recCount = 0 Then
                %>
                    <tr><td colspan="7" class="text-center text-muted" style="padding:40px">暂无收货记录</td></tr>
                <% End If %>
                <% End If
                Set rsReceipts = Nothing %>
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- 待收货采购订单 -->
    <div class="card">
        <div class="card-header">待收货采购订单</div>
        <div class="card-body" style="overflow-x:auto;">
            <table>
                <thead>
                    <tr>
                        <th>采购单号</th><th>类型</th><th>供应商</th><th>分类</th><th>订单日期</th><th>金额</th><th>状态</th><th>操作</th>
                    </tr>
                </thead>
                <tbody>
                <%
                Dim sqlPendingPO, rsPendingPO
                sqlPendingPO = "SELECT po.*, s.SupplierName, pc.CategoryName FROM (PurchaseOrders po LEFT JOIN Suppliers s ON po.SupplierID=s.SupplierID) LEFT JOIN PurchaseCategories pc ON po.CategoryCode=pc.CategoryCode WHERE po.Status IN ('Submitted','Ordered','FinanceApproved','PartialReceived') ORDER BY po.PurchaseID DESC"
                Set rsPendingPO = conn.Execute(sqlPendingPO)
                If Not rsPendingPO Is Nothing Then
                    Dim pendingCount
                    pendingCount = 0
                    Do While Not rsPendingPO.EOF And pendingCount < 50
                        pendingCount = pendingCount + 1
                %>
                    <tr>
                        <td><strong><%=rsPendingPO("PurchaseNo")%></strong></td>
                        <td>
                            <%
                                Dim poOT : poOT = rsPendingPO("OrderType") & ""
                                If poOT = "Packaging" Then
                                    Response.Write "<span style='background:rgba(156,39,176,0.2);color:#CE93D8;padding:3px 10px;border-radius:12px;font-size:12px;'>包装物</span>"
                                ElseIf poOT = "Bottle" Then
                                    Response.Write "<span style='background:rgba(33,150,243,0.2);color:#64B5F6;padding:3px 10px;border-radius:12px;font-size:12px;'>瓶子</span>"
                                ElseIf poOT = "Printing" Then
                                    Response.Write "<span style='background:rgba(76,175,80,0.2);color:#81C784;padding:3px 10px;border-radius:12px;font-size:12px;'>印刷品</span>"
                                ElseIf poOT = "SprayHead" Then
                                    Response.Write "<span style='background:rgba(255,152,0,0.2);color:#FFB74D;padding:3px 10px;border-radius:12px;font-size:12px;'>喷头</span>"
                                Else
                                    Response.Write "<span style='background:rgba(255,152,0,0.2);color:#FFB74D;padding:3px 10px;border-radius:12px;font-size:12px;'>原料</span>"
                                End If
                            %>
                        </td>
                        <td><%=rsPendingPO("SupplierName") & ""%></td>
                        <td><%=rsPendingPO("CategoryName") & ""%></td>
                        <td><%=rsPendingPO("OrderDate") & ""%></td>
                        <td class="text-right"><%=FormatNumber(CDbl("0" & rsPendingPO("TotalAmount")), 2)%></td>
                        <td>
                            <%
                            Dim pendSC, pendStatusLabel, pendStatusClass
                            pendSC = UCase(Trim(rsPendingPO("Status") & ""))
                            Select Case pendSC
                                Case "PARTIALRECEIVED"
                                    pendStatusLabel = "已部分收货"
                                    pendStatusClass = "status-partial"
                                Case "ORDERED"
                                    pendStatusLabel = "已下单"
                                    pendStatusClass = "status-submitted"
                                Case "SUBMITTED"
                                    pendStatusLabel = "已提交"
                                    pendStatusClass = "status-submitted"
                                Case "FINANCEAPPROVED"
                                    pendStatusLabel = "财务已审批"
                                    pendStatusClass = "status-submitted"
                                Case "APPROVED"
                                    pendStatusLabel = "已审批"
                                    pendStatusClass = "status-submitted"
                                Case "RECEIVED"
                                    pendStatusLabel = "已收货"
                                    pendStatusClass = "status-received"
                                Case "PENDING"
                                    pendStatusLabel = "待处理"
                                    pendStatusClass = "status-submitted"
                                Case Else
                                    pendStatusLabel = rsPendingPO("Status") & ""
                                    pendStatusClass = "status-submitted"
                            End Select
                            %>
                            <span class="status <%=pendStatusClass%>"><%=Server.HTMLEncode(pendStatusLabel)%></span>
                        </td>
                        <td>
                            <a href="receiving.asp?action=receive&purchase_id=<%=rsPendingPO("PurchaseID")%>" class="btn btn-primary btn-sm">收货</a>
                        </td>
                    </tr>
                <%
                        rsPendingPO.MoveNext
                    Loop
                    rsPendingPO.Close
                    If pendingCount = 0 Then
                %>
                    <tr><td colspan="8" class="text-center text-muted" style="padding:40px">暂无待收货采购订单</td></tr>
                <% End If %>
                <% End If
                Set rsPendingPO = Nothing %>
                </tbody>
            </table>
        </div>
    </div>
    <% End If %>
</div>

<!-- 批次详情弹窗 -->
<div id="batchModal" class="modal" style="display:none; position:fixed; z-index:1000; left:0; top:0; width:100%; height:100%; overflow:auto; background:rgba(0,0,0,0.7);">
    <div class="modal-content" style="background:linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); margin:10% auto; padding:0; border:1px solid rgba(255,255,255,0.1); border-radius:12px; width:90%; max-width:700px;">
        <div class="modal-header" style="padding:14px 20px; background:rgba(255,152,0,0.08); border-bottom:1px solid rgba(255,255,255,0.06); display:flex; justify-content:space-between; align-items:center;">
            <h3 style="margin:0; color:#e0e0e0; font-size:16px;"><i class="fas fa-layer-group" style="color:#FF9800;"></i> 批次追溯</h3>
            <button onclick="closeBatchModal()" style="background:none; border:none; color:#888; font-size:20px; cursor:pointer;">&times;</button>
        </div>
        <div class="modal-body" id="batchModalBody" style="padding:16px 20px; max-height:500px; overflow-y:auto;">
            <div style="text-align:center; color:#888; padding:20px;">加载中...</div>
        </div>
    </div>
</div>

<script>
function showBatchDetail(receiptId) {
    var modal = document.getElementById('batchModal');
    var body = document.getElementById('batchModalBody');
    body.innerHTML = '<div style="text-align:center;color:#888;padding:20px;"><i class="fas fa-spinner fa-spin"></i> 加载批次数据...</div>';
    modal.style.display = 'block';
    
    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'batch_detail.asp?receipt_id=' + receiptId, true);
    xhr.onload = function() {
        if (xhr.status === 200) {
            body.innerHTML = xhr.responseText;
        } else {
            body.innerHTML = '<div style="text-align:center;color:#e74c3c;padding:20px;">加载批次失败</div>';
        }
    };
    xhr.onerror = function() {
        body.innerHTML = '<div style="text-align:center;color:#e74c3c;padding:20px;">网络错误</div>';
    };
    xhr.send();
}

function closeBatchModal() {
    document.getElementById('batchModal').style.display = 'none';
}

window.onclick = function(event) {
    var modal = document.getElementById('batchModal');
    if (event.target === modal) {
        modal.style.display = 'none';
    }
};
</script>

</body>
</html>
<%
Call CloseConnection()
%>
