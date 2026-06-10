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

Function GetScalar(sql)
    Dim rs, val : val = 0
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                val = rs(0)
                rs.Close
            End If
        End If
    Else
        Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

' 自动创建 PackagingInventory 表
On Error Resume Next
conn.Execute "SELECT TOP 1 1 FROM PackagingInventory"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE PackagingInventory (" & _
        "PackagingID INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "PackagingName NVARCHAR(100) NOT NULL, " & _
        "PackagingType NVARCHAR(30) DEFAULT 'Box', " & _
        "Description NVARCHAR(MAX), " & _
        "StockQty INT DEFAULT 0, " & _
        "SafetyStock INT DEFAULT 0, " & _
        "UnitCost DECIMAL(19,4) DEFAULT 0, " & _
        "IsActive BIT DEFAULT 1, " & _
        "UpdatedAt DATETIME)"
    If Err.Number <> 0 Then Err.Clear
End If
Err.Clear
On Error GoTo 0

' POST处理
Dim pAction, pkgId
pAction = Request.Form("action")
If pAction = "add" Then
    Dim pName, pType, pDesc : pName = Replace(Request.Form("pkgName"),"'","''")
    pType = Replace(Request.Form("pkgType"),"'","''")
    pDesc = Replace(Request.Form("pkgDesc"),"'","''")
    If pName <> "" Then
        conn.Execute "INSERT INTO PackagingInventory (PackagingName, PackagingType, Description, IsActive, UpdatedAt) VALUES ('" & pName & "','" & pType & "','" & pDesc & "',1,GETDATE())"
        Response.Redirect "packaging_inventory.asp?msg=包装物已添加"
        Response.End
    End If
ElseIf pAction = "restock" Then
    pkgId = Request.Form("pkgId")
    Dim pQty : pQty = SafeNum(Request.Form("qty"))
    If IsNumeric(pkgId) And pQty > 0 Then
        conn.Execute "UPDATE PackagingInventory SET StockQty = ISNULL(StockQty,0) + " & pQty & ", UpdatedAt = GETDATE() WHERE PackagingID = " & CLng(pkgId)
        Response.Redirect "packaging_inventory.asp?msg=入库成功"
        Response.End
    End If
ElseIf pAction = "update_safety" Then
    pkgId = Request.Form("pkgId")
    Dim pSafety : pSafety = SafeNum(Request.Form("safetyStock"))
    If IsNumeric(pkgId) Then
        conn.Execute "UPDATE PackagingInventory SET SafetyStock = " & pSafety & ", UpdatedAt = GETDATE() WHERE PackagingID = " & CLng(pkgId)
        Response.Redirect "packaging_inventory.asp?msg=安全库存已更新"
        Response.End
    End If
End If

' 统计
Dim pgTotal, pgActive, pgLow, pgTotalQty, pgValue
pgTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory"))
pgActive = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory WHERE IsActive=1"))
pgLow = SafeNum(GetScalar("SELECT COUNT(*) FROM PackagingInventory WHERE ISNULL(StockQty,0) <= ISNULL(SafetyStock,0) AND ISNULL(SafetyStock,0) > 0"))
pgTotalQty = SafeNum(GetScalar("SELECT ISNULL(SUM(ISNULL(StockQty,0)),0) FROM PackagingInventory"))
pgValue = SafeNum(GetScalar("SELECT ISNULL(SUM(ISNULL(StockQty,0) * ISNULL(UnitCost,0)),0) FROM PackagingInventory"))

Dim filterType : filterType = Request.QueryString("type")
Dim whereSql : whereSql = ""
If filterType <> "" Then whereSql = " WHERE PackagingType = '" & Replace(filterType,"'","''") & "'"

Dim rsPkg
On Error Resume Next
Set rsPkg = conn.Execute("SELECT * FROM PackagingInventory" & whereSql & " ORDER BY PackagingType, PackagingName")
If Err.Number <> 0 Then
    Err.Clear
    Set rsPkg = Nothing
End If
On Error GoTo 0
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>包装物库存 - 产品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #00BCD4; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: var(--accent); }
        .filter-tabs { display: flex; gap: 6px; }
        .filter-tab { padding: 6px 16px; border-radius: 20px; font-size: 12px; color: #888; text-decoration: none; border: 1px solid rgba(255,255,255,0.1); transition: all 0.2s; }
        .filter-tab:hover { color: #e0e0e0; border-color: rgba(255,255,255,0.25); }
        .filter-tab.active { background: rgba(0,188,212,0.15); color: #4dd0e1; border-color: rgba(0,188,212,0.3); }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 16px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 11px; color: #888; margin-bottom: 6px; }
        .stat-card .stat-value { font-size: 24px; font-weight: 700; color: var(--accent); }
        .stat-card .stat-sub { font-size: 11px; color: #666; margin-top: 4px; }
        .data-table { width: 100%; border-collapse: collapse; background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.06); }
        .data-table th { background: rgba(0,188,212,0.12); color: #4dd0e1; font-weight: 600; padding: 12px 14px; text-align: left; font-size: 13px; }
        .data-table td { padding: 10px 14px; color: #e0e0e0; font-size: 13px; border-bottom: 1px solid rgba(255,255,255,0.04); }
        .data-table tr:hover td { background: rgba(255,255,255,0.03); }
        .stock-bar { height: 6px; border-radius: 3px; background: rgba(255,255,255,0.1); margin-top: 4px; overflow: hidden; }
        .stock-bar-fill { height: 100%; border-radius: 3px; transition: width 0.3s; }
        .stock-normal { background: #4CAF50; }
        .stock-low { background: #FF9800; }
        .stock-zero { background: #f44336; }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-box { background: rgba(0,188,212,0.12); color: #4dd0e1; }
        .badge-bag { background: rgba(156,39,176,0.12); color: #ba68c8; }
        .badge-label { background: rgba(255,152,0,0.12); color: #ffb74d; }
        .badge-other { background: rgba(158,158,158,0.12); color: #bdbdbd; }
        .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 999; align-items: center; justify-content: center; }
        .modal.active { display: flex; }
        .modal-content { background: #2d2d44; border-radius: 12px; padding: 24px; width: 420px; max-width: 90vw; border: 1px solid rgba(255,255,255,0.1); }
        .modal-title { font-size: 18px; color: #e0e0e0; margin-bottom: 16px; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input, .form-group select { width: 100%; padding: 9px 12px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; }
        .form-group input:focus, .form-group select:focus { outline: none; border-color: var(--accent); }
        .form-group select option { background: #2d2d44; color: #e0e0e0; }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
</style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-box-open"></i> 包装物库存管理</h2>
            <div class="filter-tabs">
                <a href="packaging_inventory.asp" class="filter-tab <%=IIf(filterType="","active","")%>">全部</a>
                <a href="packaging_inventory.asp?type=Box" class="filter-tab <%=IIf(filterType="Box","active","")%>"><i class="fas fa-box"></i> 盒子</a>
                <a href="packaging_inventory.asp?type=Bag" class="filter-tab <%=IIf(filterType="Bag","active","")%>"><i class="fas fa-shopping-bag"></i> 袋子</a>
                <a href="packaging_inventory.asp?type=Label" class="filter-tab <%=IIf(filterType="Label","active","")%>"><i class="fas fa-tag"></i> 标签</a>
                <a href="packaging_inventory.asp?type=Ribbon" class="filter-tab <%=IIf(filterType="Ribbon","active","")%>"><i class="fas fa-ribbon"></i> 丝带</a>
                <a href="packaging_inventory.asp?type=Card" class="filter-tab <%=IIf(filterType="Card","active","")%>"><i class="fas fa-id-card"></i> 卡片</a>
                <a href="packaging_inventory.asp?type=Other" class="filter-tab <%=IIf(filterType="Other","active","")%>"><i class="fas fa-ellipsis-h"></i> 其他</a>
            </div>
        </div>

        <% Dim pMsg : pMsg = Request.QueryString("msg")
        If pMsg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(pMsg)%></div>
        <% End If %>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-box-open"></i> 包装物种类</div>
                <div class="stat-value"><%=pgActive%></div>
                <div class="stat-sub">总计 <%=pgTotal%> 种</div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-cubes"></i> 库存总数</div>
                <div class="stat-value" style="color:#4CAF50;"><%=pgTotalQty%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-exclamation-triangle"></i> 低库存预警</div>
                <div class="stat-value <%=IIf(pgLow>0,"style='color:#f44336'","style='color:#4CAF50'")%>"><%=pgLow%></div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fas fa-dollar-sign"></i> 库存价值</div>
                <div class="stat-value" style="color:#4CAF50;">¥<%=FormatNumber(pgValue,2)%></div>
            </div>
        </div>

        <div style="margin-bottom:15px;">
            <button class="btn btn-primary" onclick="document.getElementById('addModal').classList.add('active')"><i class="fas fa-plus"></i> 新增包装物</button>
        </div>

        <table class="data-table">
            <thead>
                <tr>
                    <th>名称</th>
                    <th>类型</th>
                    <th>描述</th>
                    <th>库存量</th>
                    <th>安全库存</th>
                    <th>单价(元)</th>
                    <th>价值</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsPkg Is Nothing Then
                    Do While Not rsPkg.EOF
                        Dim pId, pName2, pType2, pDesc2, pQty2, pSafety2, pCost2, pActive2
                        pId = rsPkg("PackagingID")
                        pName2 = rsPkg("PackagingName") & ""
                        pType2 = rsPkg("PackagingType") & ""
                        pDesc2 = rsPkg("Description") & ""
                        pQty2 = SafeNum(rsPkg("StockQty"))
                        pSafety2 = SafeNum(rsPkg("SafetyStock"))
                        pCost2 = SafeNum(rsPkg("UnitCost"))
                        pActive2 = rsPkg("IsActive")
                        Dim pStockClass, pStockPct, pTypeBadge, pValue2
                        pValue2 = pQty2 * pCost2
                        If pQty2 <= 0 Then
                            pStockClass = "stock-zero"
                            pStockPct = 0
                        ElseIf pSafety2 > 0 And pQty2 <= pSafety2 Then
                            pStockClass = "stock-low"
                            pStockPct = Int((pQty2/(pSafety2*2))*100)
                        Else
                            pStockClass = "stock-normal"
                            pStockPct = 100
                        End If
                        If pStockPct > 100 Then pStockPct = 100
                        Select Case pType2
                            Case "Box" : pTypeBadge = "badge-box"
                            Case "Bag" : pTypeBadge = "badge-bag"
                            Case "Label" : pTypeBadge = "badge-label"
                            Case Else : pTypeBadge = "badge-other"
                        End Select
                %>
                <tr>
                    <td><strong><%=Server.HTMLEncode(pName2)%></strong></td>
                    <td><span class="badge <%=pTypeBadge%>"><%=Server.HTMLEncode(pType2)%></span></td>
                    <td style="color:#888;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%=IIf(pDesc2="","-",Server.HTMLEncode(pDesc2))%></td>
                    <td>
                        <%=pQty2%>
                        <div class="stock-bar"><div class="stock-bar-fill <%=pStockClass%>" style="width:<%=pStockPct%>%;"></div></div>
                    </td>
                    <td><%=IIf(pSafety2=0,"-",pSafety2)%></td>
                    <td>¥<%=FormatNumber(pCost2,2)%></td>
                    <td style="color:<%=IIf(pValue2>1000,"#81c784","#888")%>;">¥<%=FormatNumber(pValue2,2)%></td>
                    <td>
                        <button class="btn btn-sm btn-outline" onclick="openRestock(<%=pId%>,'<%=Server.HTMLEncode(Replace(pName2,"'","\'"))%>')"><i class="fas fa-plus"></i> 入库</button>
                        <button class="btn btn-sm btn-outline" onclick="openSafety(<%=pId%>,<%=pSafety2%>,'<%=Server.HTMLEncode(Replace(pName2,"'","\'"))%>')"><i class="fas fa-shield-alt"></i></button>
                    </td>
                </tr>
                <%      rsPkg.MoveNext
                    Loop
                    rsPkg.Close : Set rsPkg = Nothing
                End If %>
            </tbody>
        </table>
    </div>

    <!-- 新增 Modal -->
    <div id="addModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-plus-circle"></i> 新增包装物</div>
            <form method="post">
                <input type="hidden" name="action" value="add">
                <div class="form-group">
                    <label>名称 *</label>
                    <input type="text" name="pkgName" required>
                </div>
                <div class="form-group">
                    <label>类型</label>
                    <select name="pkgType">
                        <option value="Box">盒子</option>
                        <option value="Bag">袋子</option>
                        <option value="Label">标签/贴纸</option>
                        <option value="Ribbon">丝带</option>
                        <option value="Card">卡片</option>
                        <option value="Other">其他</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>描述</label>
                    <input type="text" name="pkgDesc">
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="document.getElementById('addModal').classList.remove('active')">取消</button>
                    <button type="submit" class="btn btn-primary">添加</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 入库 Modal -->
    <div id="restockModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-plus-circle"></i> 包装物入库 — <span id="restockName"></span></div>
            <form method="post">
                <input type="hidden" name="action" value="restock">
                <input type="hidden" name="pkgId" id="restockId">
                <div class="form-group">
                    <label>入库数量</label>
                    <input type="number" name="qty" min="1" value="1" required>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="document.getElementById('restockModal').classList.remove('active')">取消</button>
                    <button type="submit" class="btn btn-primary">确认入库</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 安全库存 Modal -->
    <div id="safetyModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-shield-alt"></i> 安全库存 — <span id="safetyName"></span></div>
            <form method="post">
                <input type="hidden" name="action" value="update_safety">
                <input type="hidden" name="pkgId" id="safetyId">
                <div class="form-group">
                    <label>安全库存阈值</label>
                    <input type="number" name="safetyStock" id="safetyStockVal" min="0" required>
                </div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="document.getElementById('safetyModal').classList.remove('active')">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    function openRestock(id, name) {
        document.getElementById('restockId').value = id;
        document.getElementById('restockName').textContent = name;
        document.getElementById('restockModal').classList.add('active');
    }
    function openSafety(id, val, name) {
        document.getElementById('safetyId').value = id;
        document.getElementById('safetyStockVal').value = val;
        document.getElementById('safetyName').textContent = name;
        document.getElementById('safetyModal').classList.add('active');
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
