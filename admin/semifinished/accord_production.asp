<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Function IIF(cond, tVal, fVal)
    If cond Then IIF = tVal Else IIF = fVal
End Function

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

Dim action, msg, msgType
action = Trim(Request.Form("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    
    If action = "start_production" Then
        Dim prodAccordID, prodPlannedQty, prodNotes
        prodAccordID = SafeNum(Request.Form("accord_recipe_id"))
        prodPlannedQty = SafeNum(Request.Form("planned_qty"))
        prodNotes = Trim(Request.Form("notes"))
        
        If prodAccordID > 0 And prodPlannedQty > 0 Then
            Dim prodBatchNo
            prodBatchNo = "ACP" & Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2)
            
            Dim prodRecipeID, prodNoteID, prodBatchSize
            prodRecipeID = 0 : prodNoteID = 0 : prodBatchSize = 100
            Dim rsAccord
            Set rsAccord = conn.Execute("SELECT RecipeID, NoteID, BatchSize FROM RecipeAccords WHERE AccordRecipeID=" & prodAccordID & " AND Status='Published'")
            If Not rsAccord Is Nothing Then
                If Not rsAccord.EOF Then
                    prodRecipeID = SafeNum(rsAccord("RecipeID"))
                    prodNoteID = SafeNum(rsAccord("NoteID"))
                    prodBatchSize = SafeNum(rsAccord("BatchSize"))
                End If
                rsAccord.Close
            End If
            Set rsAccord = Nothing
            
            If prodNoteID > 0 Then
                Dim stockOK, insufficientMaterials
                stockOK = True : insufficientMaterials = ""
                
                Dim rsMaterials
                Set rsMaterials = conn.Execute("SELECT ram.MaterialID, ram.MaterialName, ram.Percentage, ram.PlannedQty, rmi.StockQty FROM (RecipeAccordMaterials ram LEFT JOIN RawMaterialInventory rmi ON ram.MaterialID=rmi.MaterialID) WHERE ram.AccordRecipeID=" & prodAccordID)
                If Not rsMaterials Is Nothing Then
                    Do While Not rsMaterials.EOF
                        Dim matID, matName, matPct, matNeed, matStock
                        matID = SafeNum(rsMaterials("MaterialID"))
                        matName = CStr(rsMaterials("MaterialName") & "")
                        matPct = SafeNum(rsMaterials("Percentage"))
                        matStock = SafeNum(rsMaterials("StockQty"))
                        matNeed = (matPct / prodBatchSize) * prodPlannedQty
                        If matNeed > matStock Then
                            stockOK = False
                            insufficientMaterials = insufficientMaterials & matName & "(需要" & FormatNumber(matNeed,1) & "g,库存" & FormatNumber(matStock,1) & "g) "
                        End If
                        rsMaterials.MoveNext
                    Loop
                    rsMaterials.Close
                End If
                Set rsMaterials = Nothing
                
                If Not stockOK Then
                    msg = "原材料库存不足: " & insufficientMaterials
                    msgType = "error"
                Else
                    On Error Resume Next
                    Err.Clear
                    Call BeginTransaction()
                    
                    conn.Execute "INSERT INTO AccordProductions (BatchNo, AccordRecipeID, PlannedQty, Status, StartedAt, WorkCenter, Notes, CreatedAt, UpdatedAt, NoteID, NoteName) VALUES ('" & _
                        prodBatchNo & "', " & prodAccordID & ", " & prodPlannedQty & ", 'InProgress', GETDATE(), 'SEMI', '" & SafeSQL(prodNotes) & "', GETDATE(), GETDATE(), " & prodNoteID & ", (SELECT NoteName FROM FragranceNotes WHERE NoteID=" & prodNoteID & "))"
                    
                    If Err.Number <> 0 Then
                        msg = "创建生产单失败: " & Err.Description
                        msgType = "error"
                        Call RollbackTransaction()
                        Err.Clear
                    Else
                        Dim newProdID
                        Set rsAccord = conn.Execute("SELECT SCOPE_IDENTITY()")
                        newProdID = 0
                        If Not rsAccord Is Nothing Then
                            If Not rsAccord.EOF Then newProdID = CLng(rsAccord(0))
                            rsAccord.Close
                        End If
                        Set rsAccord = Nothing
                        
                        If newProdID > 0 Then
                            Dim anyError : anyError = False
                            
                            Set rsMaterials = conn.Execute("SELECT MaterialID, MaterialName, Percentage, PlannedQty FROM RecipeAccordMaterials WHERE AccordRecipeID=" & prodAccordID)
                            If Not rsMaterials Is Nothing Then
                                Do While Not rsMaterials.EOF
                                    matID = SafeNum(rsMaterials("MaterialID"))
                                    matName = CStr(rsMaterials("MaterialName") & "")
                                    matPct = SafeNum(rsMaterials("Percentage"))
                                    matNeed = (matPct / prodBatchSize) * prodPlannedQty
                                    
                                    conn.Execute "UPDATE RawMaterialInventory SET StockQty = StockQty - " & matNeed & ", UpdatedAt = GETDATE() WHERE MaterialID=" & matID
                                    If Err.Number <> 0 Then anyError = True : Err.Clear
                                    
                                    conn.Execute "INSERT INTO InventoryTransactions (NoteID, MaterialID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES (" & _
                                        prodNoteID & ", " & matID & ", " & matNeed & ", '香调生产消耗', 'OUT', '批次" & prodBatchNo & "消耗[" & SafeSQL(matName) & "]', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
                                    If Err.Number <> 0 Then anyError = True : Err.Clear
                                    
                                    conn.Execute "INSERT INTO AccordProductionDetails (ProductionID, MaterialID, MaterialName, PlannedQty, ActualQty) VALUES (" & _
                                        newProdID & ", " & matID & ", '" & SafeSQL(matName) & "', " & matNeed & ", " & matNeed & ")"
                                    If Err.Number <> 0 Then anyError = True : Err.Clear
                                    
                                    rsMaterials.MoveNext
                                Loop
                                rsMaterials.Close
                            End If
                            Set rsMaterials = Nothing
                            
                            If Not anyError Then
                                Dim noteName
                                noteName = CStr(GetScalar("SELECT NoteName FROM FragranceNotes WHERE NoteID=" & prodNoteID) & "")
                                
                                Dim rsNI
                                Set rsNI = conn.Execute("SELECT InventoryID FROM NoteInventory WHERE NoteID=" & prodNoteID)
                                If Not rsNI Is Nothing Then
                                    If Not rsNI.EOF Then
                                        conn.Execute "UPDATE NoteInventory SET StockQuantity = StockQuantity + " & prodPlannedQty & ", UpdatedAt = GETDATE() WHERE NoteID=" & prodNoteID
                                    Else
                                        conn.Execute "INSERT INTO NoteInventory (NoteID, StockQuantity, MinStockLevel, UpdatedAt) VALUES (" & prodNoteID & ", " & prodPlannedQty & ", 50, GETDATE())"
                                    End If
                                    rsNI.Close
                                End If
                                Set rsNI = Nothing
                                
                                If Err.Number <> 0 Then anyError = True : Err.Clear
                                
                                conn.Execute "INSERT INTO InventoryTransactions (NoteID, Quantity, TransactionType, TransactionDirection, Notes, CreatedBy, CreatedAt) VALUES (" & _
                                    prodNoteID & ", " & prodPlannedQty & ", '香调生产产出', 'IN', '批次" & prodBatchNo & "产出香调[" & SafeSQL(noteName) & "]', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE())"
                                
                                conn.Execute "UPDATE AccordProductions SET Status='Completed', ActualQty=" & prodPlannedQty & ", CompletedAt=GETDATE(), UpdatedAt=GETDATE(), NoteName='" & SafeSQL(noteName) & "' WHERE ProductionID=" & newProdID
                            End If
                            
                            If Not anyError Then
                                Call CommitTransaction()
                                Response.Redirect "accord_production.asp?msg=香调生产完成！批次：" & prodBatchNo
                                Response.End
                            Else
                                Call RollbackTransaction()
                                msg = "香调生产失败，数据已回滚"
                                msgType = "error"
                            End If
                        Else
                            Call RollbackTransaction()
                            msg = "获取生产单ID失败"
                            msgType = "error"
                        End If
                    End If
                    On Error GoTo 0
                End If
            Else
                msg = "无效的香调配方"
                msgType = "error"
            End If
        Else
            msg = "请选择配方和数量"
            msgType = "error"
        End If
    
    ElseIf action = "qc_report" Then
        Dim qcProdID, qcResult, qcNotes
        qcProdID = SafeNum(Request.Form("production_id"))
        qcResult = Trim(Request.Form("qc_result"))
        qcNotes = Trim(Request.Form("qc_notes"))
        
        If qcProdID > 0 And qcResult <> "" Then
            Dim qcBatchNo
            qcBatchNo = CStr(GetScalar("SELECT BatchNo FROM AccordProductions WHERE ProductionID=" & qcProdID) & "")
            
            conn.Execute "INSERT INTO AccordQCReports (ProductionID, BatchNo, QCResult, TestDate, TesterID, TesterName, Notes, CreatedAt) VALUES (" & _
                qcProdID & ", '" & SafeSQL(qcBatchNo) & "', '" & SafeSQL(qcResult) & "', GETDATE(), 0, '" & SafeSQL(Session("AdminUsername")) & "', '" & SafeSQL(qcNotes) & "', GETDATE())"
            
            If qcResult = "Pass" Then
                conn.Execute "UPDATE AccordProductions SET Status='QC' WHERE ProductionID=" & qcProdID
                Response.Redirect "accord_production.asp?msg=质检通过！香调已可用"
            Else
                conn.Execute "UPDATE AccordProductions SET Status='QC_Fail' WHERE ProductionID=" & qcProdID
                Response.Redirect "accord_production.asp?msg=质检未通过，需返工处理"
            End If
            Response.End
        End If
    End If
End If

' ========== 统计 ==========
Dim statInProgress, statCompleted, statLowStock
statInProgress = 0 : statCompleted = 0 : statLowStock = 0
On Error Resume Next
statInProgress = SafeNum(GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status='InProgress'"))
statCompleted = SafeNum(GetScalar("SELECT COUNT(*) FROM AccordProductions WHERE Status IN ('Completed','QC') AND CompletedAt >= CAST(GETDATE() AS DATE)"))
statLowStock = SafeNum(GetScalar("SELECT COUNT(*) FROM RawMaterialInventory WHERE StockQty <= SafetyStock AND SafetyStock > 0"))
On Error GoTo 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>Accord生产 - 半成品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --card: #16213e; --border: #2a2a4a; --text: #e0e0e0; --accent: #2196F3; --success: #4CAF50; --warning: #FF9800; --danger: #f44336; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        
        .page-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #2196F3; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 32px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 13px; color: #888; display: block; margin-top: 5px; }
        .stat-card.inprogress .num { color: #2196F3; }
        .stat-card.completed .num { color: #4CAF50; }
        .stat-card.warning .num { color: #f44336; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(33,150,243,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(33,150,243,0.15); color: #64b5f6; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-progress { background: rgba(33,150,243,0.15); color: #64b5f6; }
        .status-completed { background: rgba(76,175,80,0.15); color: #81c784; }
        .status-qc { background: rgba(156,39,176,0.15); color: #ce93d8; }
        .status-fail { background: rgba(244,67,54,0.15); color: #e57373; }
        
        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #e57373; border: 1px solid rgba(244,67,54,0.3); }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); width: 90%; max-width: 550px; margin: 80px auto; padding: 30px; border-radius: 15px; border: 1px solid rgba(255,255,255,0.06); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .modal-header h3 { margin: 0; font-size: 18px; color: #e0e0e0; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .modal-close:hover { color: #fff; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 25px; }
        
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; margin-bottom: 6px; font-weight: 600; color: #e0e0e0; font-size: 13px; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 10px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        .form-group input:focus, .form-group select:focus, .form-group textarea:focus { outline: none; border-color: #2196F3; }
        
        .security-badge { display: inline-block; padding: 2px 8px; background: rgba(255,152,0,0.15); color: #ffb74d; border-radius: 4px; font-size: 11px; margin-left: 6px; }
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .text-right { text-align: right; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-cogs"></i> Accord生产 <span class="security-badge">原料→香调</span></h2>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card inprogress">
                <span class="num"><%=statInProgress%></span>
                <span class="label">生产中</span>
            </div>
            <div class="stat-card completed">
                <span class="num"><%=statCompleted%></span>
                <span class="label">今日完成</span>
            </div>
            <div class="stat-card warning">
                <span class="num" style="color:<%=IIF(statLowStock>0,"#f44336","#4CAF50")%>"><%=statLowStock%></span>
                <span class="label">原料低库存预警</span>
            </div>
        </div>
        
        <!-- 已发布香调配方 -->
        <div class="card">
            <div class="card-header">
                已发布香调配方（原材料→香调）
                <span class="text-muted" style="font-weight:normal;">仅显示技术中心已发布配方</span>
            </div>
            <div class="card-body">
                <table>
                    <thead><tr><th>配方名称</th><th>产出香调</th><th>批量(ml)</th><th>原材料数</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim rsAccords, acRowCount : acRowCount = 0
                    Set rsAccords = conn.Execute("SELECT ra.AccordRecipeID, ra.BatchSize, ra.RecipeName, ra.NoteID, fn.NoteName, (SELECT COUNT(*) FROM RecipeAccordMaterials WHERE AccordRecipeID=ra.AccordRecipeID) AS MatCount FROM RecipeAccords ra LEFT JOIN FragranceNotes fn ON ra.NoteID=fn.NoteID WHERE ra.Status='Published' ORDER BY ra.PublishedAt DESC")
                    If Not rsAccords Is Nothing Then
                        Do While Not rsAccords.EOF
                            acRowCount = acRowCount + 1
                            Dim acID : acID = rsAccords("AccordRecipeID")
                            Dim acNoteName : acNoteName = CStr(rsAccords("NoteName") & "")
                    %>
                        <tr>
                            <td><strong><%=rsAccords("RecipeName") & ""%></strong></td>
                            <td><%=acNoteName%></td>
                            <td><%=rsAccords("BatchSize")%></td>
                            <td><%=rsAccords("MatCount")%></td>
                            <td>
                                <button class="btn btn-primary btn-sm" onclick="openProductionModal(<%=acID%>,'<%=Server.HTMLEncode(acNoteName)%>',<%=rsAccords("BatchSize")%>)">开始生产</button>
                            </td>
                        </tr>
                    <%
                            rsAccords.MoveNext
                        Loop
                        rsAccords.Close
                    End If
                    Set rsAccords = Nothing
                    If acRowCount = 0 Then
                    %>
                        <tr><td colspan="5" class="text-center text-muted" style="padding:40px;">暂无已发布的香调配方，请等待技术中心发布</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- 生产记录 -->
        <div class="card">
            <div class="card-header">Accord生产记录</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>批次号</th><th>香调</th><th>计划量</th><th>实际量</th><th>状态</th><th>时间</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim rsProd, prodRowCount : prodRowCount = 0
                    Set rsProd = conn.Execute("SELECT TOP 30 ProductionID, BatchNo, NoteName, PlannedQty, ActualQty, Status, StartedAt, CompletedAt FROM AccordProductions ORDER BY ProductionID DESC")
                    If Not rsProd Is Nothing Then
                        Do While Not rsProd.EOF
                            prodRowCount = prodRowCount + 1
                            Dim pStatus : pStatus = CStr(rsProd("Status") & "")
                            Dim pStatusClass
                            If pStatus = "InProgress" Then
                                pStatusClass = "status-progress"
                            ElseIf pStatus = "Completed" Then
                                pStatusClass = "status-completed"
                            ElseIf pStatus = "QC" Then
                                pStatusClass = "status-qc"
                            Else
                                pStatusClass = "status-fail"
                            End If
                    %>
                        <tr>
                            <td><strong><%=rsProd("BatchNo") & ""%></strong></td>
                            <td><%=rsProd("NoteName") & ""%></td>
                            <td><%=rsProd("PlannedQty")%></td>
                            <td><%=IIF(IsNull(rsProd("ActualQty")) Or rsProd("ActualQty")="","-",rsProd("ActualQty"))%></td>
                            <td><span class="status-badge <%=pStatusClass%>"><%=pStatus%></span></td>
                            <td style="color:#888;font-size:13px;"><%=IIF(IsNull(rsProd("StartedAt")) Or rsProd("StartedAt")="","-",rsProd("StartedAt"))%></td>
                            <td>
                                <% If pStatus = "Completed" Then %>
                                <button class="btn btn-success btn-sm" onclick="openQCModal(<%=rsProd("ProductionID")%>,'<%=rsProd("BatchNo") & ""%>')">质检</button>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rsProd.MoveNext
                        Loop
                        rsProd.Close
                    End If
                    Set rsProd = Nothing
                    If prodRowCount = 0 Then
                    %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:40px;">暂无生产记录</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 开始生产弹窗 -->
    <div id="productionModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>开始Accord生产 - <span id="modalNoteName"></span></h3>
                <button class="modal-close" onclick="closeModal('productionModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="start_production">
                <input type="hidden" name="accord_recipe_id" id="modalAccordID">
                <div class="form-group">
                    <label>计划生产量 (ml)</label>
                    <input type="number" name="planned_qty" id="modalQty" required min="1" placeholder="输入生产量">
                </div>
                <div class="form-group">
                    <label>备注</label>
                    <textarea name="notes" rows="2" placeholder="生产备注（可选）"></textarea>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('productionModal')">取消</button>
                    <button type="submit" class="btn btn-primary">确认开始生产</button>
                </div>
            </form>
        </div>
    </div>
    
    <!-- 质检弹窗 -->
    <div id="qcModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>质量检验 - <span id="qcBatchNo"></span></h3>
                <button class="modal-close" onclick="closeModal('qcModal')">&times;</button>
            </div>
            <form method="post">
                <input type="hidden" name="action" value="qc_report">
                <input type="hidden" name="production_id" id="qcProdID">
                <div class="form-group">
                    <label>检验结果</label>
                    <select name="qc_result" required>
                        <option value="">请选择</option>
                        <option value="Pass">合格 (Pass)</option>
                        <option value="Fail">不合格 (Fail)</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>检验备注</label>
                    <textarea name="qc_notes" rows="2" placeholder="检验备注（可选）"></textarea>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('qcModal')">取消</button>
                    <button type="submit" class="btn btn-success">提交检验结果</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    function openProductionModal(accordID, noteName, batchSize) {
        document.getElementById('modalAccordID').value = accordID;
        document.getElementById('modalNoteName').innerText = noteName;
        document.getElementById('modalQty').value = batchSize;
        document.getElementById('productionModal').style.display = 'block';
    }
    function openQCModal(prodID, batchNo) {
        document.getElementById('qcProdID').value = prodID;
        document.getElementById('qcBatchNo').innerText = batchNo;
        document.getElementById('qcModal').style.display = 'block';
    }
    function closeModal(id) {
        document.getElementById(id).style.display = 'none';
    }
    window.onclick = function(event) {
        if (event.target.classList.contains('modal')) event.target.style.display = 'none';
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
