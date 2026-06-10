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

' 确保RawMaterialInventory表有WeightedUnitCost字段
On Error Resume Next
Dim rsWCCheck
Set rsWCCheck = conn.Execute("SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='RawMaterialInventory' AND COLUMN_NAME='WeightedUnitCost'")
If Not rsWCCheck Is Nothing Then
    If Not rsWCCheck.EOF Then
        If CLng(rsWCCheck(0)) = 0 Then
            conn.Execute "ALTER TABLE RawMaterialInventory ADD WeightedUnitCost DECIMAL(18,6) DEFAULT 0"
            Err.Clear
        End If
    End If
    rsWCCheck.Close
End If
Set rsWCCheck = Nothing
Err.Clear
On Error GoTo 0

' 检查WeightedUnitCost列是否实际存在
Dim hasWCostCol : hasWCostCol = False
On Error Resume Next
Dim rsColTest
Set rsColTest = conn.Execute("SELECT COL_LENGTH('RawMaterialInventory','WeightedUnitCost')")
If Err.Number = 0 And Not rsColTest Is Nothing Then
    If Not rsColTest.EOF Then
        If Not IsNull(rsColTest(0)) Then hasWCostCol = True
    End If
    rsColTest.Close
End If
Set rsColTest = Nothing
Err.Clear
On Error GoTo 0

Dim wcostExpr
If hasWCostCol Then
    wcostExpr = "ISNULL(WeightedUnitCost, UnitPrice)"
Else
    wcostExpr = "UnitPrice"
End If

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
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

' ========== 预构建原料选项 ==========
Dim matOptionsHTML
matOptionsHTML = ""
Dim rsMO
Set rsMO = conn.Execute("SELECT MaterialID, ItemName, StockQty, Unit, " & wcostExpr & " AS WCost FROM RawMaterialInventory WHERE StockQty > 0 ORDER BY ItemName ASC")
If Not rsMO Is Nothing Then
    Do While Not rsMO.EOF
        matOptionsHTML = matOptionsHTML & "<option value='" & rsMO("MaterialID") & "' data-wcost='" & FormatNumber(SafeNum(rsMO("WCost")),2) & "'>" & Server.HTMLEncode(rsMO("ItemName")) & " (库存:" & rsMO("StockQty") & rsMO("Unit") & ")</option>"
        rsMO.MoveNext
    Loop
    rsMO.Close
End If
Set rsMO = Nothing
Dim action, msg, msgType
action = Trim(Request.Form("action"))
msg = Trim(Request.QueryString("msg"))
msgType = "success"
If InStr(msg, "失败") > 0 Or InStr(msg, "错误") > 0 Then msgType = "error"

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "create_outbound" Then
        Dim outType, outRefID, outRefType, outNotes, outDetailCount
        outType = Trim(Request.Form("outbound_type"))
        outRefID = SafeNum(Request.Form("reference_id"))
        outRefType = Trim(Request.Form("reference_type"))
        outNotes = Trim(Request.Form("notes"))
        outDetailCount = SafeNum(Request.Form("detail_count"))
        
        If outType <> "" And outDetailCount > 0 Then
            Dim outNo
            outNo = "OUT" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & Right("0" & Hour(Now),2) & Right("0" & Minute(Now),2) & Right("0" & Second(Now),2)
            
            On Error Resume Next
            Err.Clear
            Call BeginTransaction()
            
            Dim sqlOut
            sqlOut = "INSERT INTO MaterialOutbound (OutboundNo, OutboundType, ReferenceID, ReferenceType, RequestedBy, OutboundDate, Status, Notes, CreatedAt) VALUES ('" & _
                outNo & "', '" & SafeSQL(outType) & "', " & outRefID & ", '" & SafeSQL(outRefType) & "', '" & SafeSQL(Session("AdminUsername")) & "', GETDATE(), 'Confirmed', "
            If outNotes <> "" Then
                sqlOut = sqlOut & "'" & SafeSQL(outNotes) & "'"
            Else
                sqlOut = sqlOut & "Null"
            End If
            sqlOut = sqlOut & ", GETDATE())"
            
            conn.Execute sqlOut
            
            If Err.Number <> 0 Then
                msg = "创建出库单失败: " & Err.Description
                msgType = "error"
                Call RollbackTransaction()
                Err.Clear
            Else
                Dim outID
                Set rsOut = conn.Execute("SELECT SCOPE_IDENTITY()")
                outID = 0
                If Not rsOut Is Nothing Then
                    If Not rsOut.EOF Then outID = CLng(rsOut(0))
                    rsOut.Close
                End If
                Set rsOut = Nothing
                
                If outID > 0 Then
                    Dim i, anyOutError : anyOutError = False
                    For i = 1 To outDetailCount
                        Dim mID, mQty, mPrice
                        mID = SafeNum(Request.Form("material_id_" & i))
                        mQty = SafeNum(Request.Form("qty_" & i))
                        mPrice = SafeNum(Request.Form("price_" & i))
                        
                        If mID > 0 And mQty > 0 Then
                            ' V10: 使用加权平均成本（WeightedUnitCost）替代末次采购价（UnitPrice）
                            If mPrice <= 0 Then
                                mPrice = SafeNum(GetScalar("SELECT " & wcostExpr & " FROM RawMaterialInventory WHERE MaterialID=" & mID))
                            End If
                            If mPrice <= 0 Then mPrice = SafeNum(GetScalar("SELECT ISNULL(UnitPrice, 0) FROM RawMaterialInventory WHERE MaterialID=" & mID))
                            conn.Execute "INSERT INTO MaterialOutboundDetails (OutboundID, MaterialID, RequestedQty, ActualQty, UnitPrice, TotalAmount) VALUES (" & _
                                outID & "," & mID & "," & mQty & "," & mQty & "," & mPrice & "," & (mQty*mPrice) & ")"
                            If Err.Number <> 0 Then anyOutError = True : Err.Clear
                            
                            conn.Execute "UPDATE RawMaterialInventory SET StockQty=StockQty-" & mQty & ", UpdatedAt=GETDATE() WHERE MaterialID=" & mID
                            If Err.Number <> 0 Then anyOutError = True : Err.Clear
                        End If
                    Next
                    
                    If Not anyOutError Then
                        Call CommitTransaction()
                        Response.Redirect "material_outbound.asp?msg=出库成功！单号：" & outNo
                        Response.End
                    Else
                        Call RollbackTransaction()
                        msg = "出库处理失败，数据已回滚"
                        msgType = "error"
                    End If
                Else
                    Call RollbackTransaction()
                    msg = "创建出库单失败"
                    msgType = "error"
                End If
            End If
            On Error GoTo 0
        Else
            msg = "请填写出库信息"
            msgType = "error"
        End If
    End If
End If

' ========== 统计 ==========
Dim moTotal, moToday, moPending
moTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM MaterialOutbound"))
moToday = SafeNum(GetScalar("SELECT COUNT(*) FROM MaterialOutbound WHERE CAST(CreatedAt AS DATE)=CAST(GETDATE() AS DATE)"))
moPending = SafeNum(GetScalar("SELECT COUNT(*) FROM MaterialOutbound WHERE Status='Pending'"))

' ========== 出库记录 ==========
Dim rsOutbound
Set rsOutbound = conn.Execute("SELECT TOP 40 mo.*, (SELECT COUNT(*) FROM MaterialOutboundDetails WHERE OutboundID=mo.OutboundID) AS DetailCount FROM MaterialOutbound mo ORDER BY mo.CreatedAt DESC")

' ========== 可用原料列表 ==========
Dim rsAvailMat
Set rsAvailMat = conn.Execute("SELECT MaterialID, ItemName, StockQty, Unit, " & wcostExpr & " AS WCost FROM RawMaterialInventory WHERE StockQty > 0 ORDER BY ItemName ASC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>原料出库 - 半成品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #9C27B0; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #9C27B0; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(156,39,176,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; display: flex; justify-content: space-between; align-items: center; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(156,39,176,0.15); color: #ce93d8; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-confirmed { background: rgba(76,175,80,0.15); color: #81c784; }
        .status-pending { background: rgba(255,152,0,0.15); color: #ffb74d; }
        

        .alert { padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .alert-success { background: rgba(76,175,80,0.15); color: #81c784; border: 1px solid rgba(76,175,80,0.3); }
        .alert-error { background: rgba(244,67,54,0.15); color: #e57373; border: 1px solid rgba(244,67,54,0.3); }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); z-index: 1000; }
        .modal-content { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); width: 90%; max-width: 650px; margin: 60px auto; padding: 30px; border-radius: 15px; border: 1px solid rgba(255,255,255,0.06); max-height: 80vh; overflow-y: auto; }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.06); position: sticky; top: 0; background: linear-gradient(135deg, #2d2d44, #1e1e32); z-index: 1; }
        .modal-header h3 { margin: 0; font-size: 18px; }
        .modal-close { background: none; border: none; color: #888; font-size: 24px; cursor: pointer; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; margin-top: 25px; }
        
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; margin-bottom: 6px; font-weight: 600; color: #e0e0e0; font-size: 13px; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 10px 12px; background: #2d2d44; border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 14px; }
        .form-group input:focus, .form-group select:focus { outline: none; border-color: #9C27B0; }
        
        .detail-row { display: flex; gap: 10px; align-items: flex-end; padding: 10px; background: rgba(0,0,0,0.1); border-radius: 8px; margin-bottom: 8px; }
        .detail-row .form-group { flex: 1; margin-bottom: 0; }
        .detail-row .btn-remove { background: #f44336; color: #fff; padding: 10px 12px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
        .text-right { text-align: right; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-truck-loading"></i> 原料出库管理</h2>
            <button class="btn btn-primary" onclick="openOutboundModal()"><i class="fas fa-plus"></i> 新建出库单</button>
        </div>
        
        <% If msg <> "" Then %>
        <div class="alert alert-<%=msgType%>"><%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#9C27B0;"><%=moTotal%></span><span class="label">总出库单</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=moToday%></span><span class="label">今日出库</span></div>
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=moPending%></span><span class="label">待处理</span></div>
        </div>
        
        <!-- 出库记录 -->
        <div class="card">
            <div class="card-header">出库记录</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>出库单号</th><th>类型</th><th>明细数</th><th>申请人</th><th>出库日期</th><th>状态</th><th>备注</th></tr></thead>
                    <tbody>
                    <%
                    Dim moRowCount : moRowCount = 0
                    If Not rsOutbound Is Nothing Then
                        Do While Not rsOutbound.EOF
                            moRowCount = moRowCount + 1
                    %>
                        <tr>
                            <td><strong><%=rsOutbound("OutboundNo") & ""%></strong></td>
                            <td><%=rsOutbound("OutboundType") & ""%></td>
                            <td><%=rsOutbound("DetailCount")%></td>
                            <td><%=rsOutbound("RequestedBy") & ""%></td>
                            <td class="text-muted"><%=IIF(IsNull(rsOutbound("OutboundDate")) Or rsOutbound("OutboundDate")="","-",Left(rsOutbound("OutboundDate"),10))%></td>
                            <td><span class="status-badge <%=IIF(rsOutbound("Status")&""="Confirmed","status-confirmed","status-pending")%>"><%=rsOutbound("Status") & ""%></span></td>
                            <td class="text-muted"><%=IIF(Len(rsOutbound("Notes")&"")>20,Left(rsOutbound("Notes")&"",20)&"...",rsOutbound("Notes")&"")%></td>
                        </tr>
                    <%
                            rsOutbound.MoveNext
                        Loop
                        rsOutbound.Close
                    End If
                    Set rsOutbound = Nothing
                    If moRowCount = 0 Then
                    %>
                        <tr><td colspan="7" class="text-center text-muted" style="padding:40px;">暂无出库记录</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- 新建出库单弹窗 -->
    <div id="outboundModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>新建原料出库单</h3>
                <button class="modal-close" onclick="closeModal('outboundModal')">&times;</button>
            </div>
            <form method="post" id="outboundForm">
                <input type="hidden" name="action" value="create_outbound">
                <div class="form-group">
                    <label>出库类型</label>
                    <select name="outbound_type" required>
                        <option value="">请选择</option>
                        <option value="生产领用">生产领用</option>
                        <option value="研发领用">研发领用</option>
                        <option value="退货出库">退货出库</option>
                        <option value="其他出库">其他出库</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>关联参考</label>
                    <input type="text" name="reference_type" placeholder="关联类型（如：生产单号）">
                </div>
                <div class="form-group">
                    <label>参考ID</label>
                    <input type="number" name="reference_id" value="0">
                </div>
                <div id="detailContainer">
                    <input type="hidden" name="detail_count" id="detailCount" value="1">
                    <div class="detail-row" id="detail_1">
                        <div class="form-group">
                            <label>原料</label>
                            <select name="material_id_1" required onchange="autoFillPrice(this, 'price_1')">
                                <option value="">选择原料</option>
                                <%
                                If Not rsAvailMat Is Nothing Then
                                    Do While Not rsAvailMat.EOF
                                %>
                                <option value="<%=rsAvailMat("MaterialID")%>" data-wcost="<%=FormatNumber(SafeNum(rsAvailMat("WCost")),2)%>"><%=Server.HTMLEncode(rsAvailMat("ItemName") & "")%> (库存:<%=FormatNumber(SafeNum(rsAvailMat("StockQty")),1)%><%=rsAvailMat("Unit")&""%>) - 成本:¥<%=FormatNumber(SafeNum(rsAvailMat("WCost")),2)%></option>
                                <%
                                        rsAvailMat.MoveNext
                                    Loop
                                    rsAvailMat.Close
                                End If
                                Set rsAvailMat = Nothing
                                %>
                            </select>
                        </div>
                        <div class="form-group"><label>数量</label><input type="number" name="qty_1" step="0.1" required></div>
                        <div class="form-group"><label>单价(加权成本)</label><input type="number" name="price_1" step="0.01" value="0" placeholder="自动填充加权成本"></div>
                        <button type="button" class="btn-remove" onclick="removeDetail(1)" style="display:none;">-</button>
                    </div>
                </div>
                <button type="button" class="btn btn-sm btn--neutral" onclick="addDetail()"><i class="fas fa-plus"></i> 添加明细行</button>
                <div class="form-group" style="margin-top:15px;"><label>备注</label><textarea name="notes" rows="2"></textarea></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn--neutral" onclick="closeModal('outboundModal')">取消</button>
                    <button type="submit" class="btn btn-primary">确认出库</button>
                </div>
            </form>
        </div>
    </div>
    
    <script>
    var detailIdx = 1;
    function autoFillPrice(selectEl, priceFieldName) {
        var opt = selectEl.options[selectEl.selectedIndex];
        var wcost = opt.getAttribute('data-wcost');
        if (wcost && parseFloat(wcost) > 0) {
            var priceInput = document.querySelector('[name="' + priceFieldName + '"]');
            if (priceInput) priceInput.value = wcost;
        }
    }
    function addDetail() {
        detailIdx++;
        document.getElementById('detailCount').value = detailIdx;
        var container = document.getElementById('detailContainer');
        var div = document.createElement('div');
        div.className = 'detail-row';
        div.id = 'detail_' + detailIdx;
        div.innerHTML = '<div class="form-group"><label>原料</label><select name="material_id_' + detailIdx + '" required onchange="autoFillPrice(this, \'price_' + detailIdx + '\')"><option value="">选择原料</option>' +
            '<%=matOptionsHTML%>' +
            '</select></div>' +
            '<div class="form-group"><label>数量</label><input type="number" name="qty_' + detailIdx + '" step="0.1" required></div>' +
            '<div class="form-group"><label>单价(加权成本)</label><input type="number" name="price_' + detailIdx + '" step="0.01" value="0" placeholder="自动填充加权成本"></div>' +
            '<button type="button" class="btn-remove" onclick="removeDetail(' + detailIdx + ')">-</button>';
        container.appendChild(div);
        if (detailIdx > 1) {
            document.querySelector('#detail_1 .btn-remove').style.display = 'inline-block';
        }
    }
    function removeDetail(idx) {
        var el = document.getElementById('detail_' + idx);
        if (el) el.remove();
        var remaining = document.querySelectorAll('.detail-row').length;
        document.getElementById('detailCount').value = remaining;
    }
    function openOutboundModal() { document.getElementById('outboundModal').style.display = 'block'; }
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }
    window.onclick = function(event) { if (event.target.classList.contains('modal')) event.target.style.display = 'none'; }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
