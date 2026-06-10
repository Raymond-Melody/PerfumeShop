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
    Set rs = Nothing : GetScalar = val
End Function

' 自动创建 ShippingCompanies 表
On Error Resume Next
conn.Execute "SELECT TOP 1 1 FROM ShippingCompanies"
If Err.Number <> 0 Then
    Err.Clear
    conn.Execute "CREATE TABLE ShippingCompanies (" & _
        "CompanyID INT IDENTITY(1,1) NOT NULL PRIMARY KEY, " & _
        "CompanyName NVARCHAR(100) NOT NULL, " & _
        "ContactPerson NVARCHAR(50), " & _
        "ContactPhone NVARCHAR(20), " & _
        "Website NVARCHAR(200), " & _
        "IsActive BIT DEFAULT 1, " & _
        "Notes NVARCHAR(MAX), " & _
        "CreatedAt DATETIME DEFAULT GETDATE(), " & _
        "UpdatedAt DATETIME)"
End If
On Error GoTo 0

' POST
Dim scAction : scAction = Request.Form("action")
If scAction = "add" Then
    Dim scName, scContact, scPhone, scWeb, scNotes
    scName = Replace(Request.Form("companyName"),"'","''")
    scContact = Replace(Request.Form("contactPerson"),"'","''")
    scPhone = Replace(Request.Form("contactPhone"),"'","''")
    scWeb = Replace(Request.Form("website"),"'","''")
    scNotes = Replace(Request.Form("notes"),"'","''")
    If scName <> "" Then
        conn.Execute "INSERT INTO ShippingCompanies (CompanyName, ContactPerson, ContactPhone, Website, Notes, IsActive) VALUES ('" & scName & "','" & scContact & "','" & scPhone & "','" & scWeb & "','" & scNotes & "',1)"
        Response.Redirect "shipping_companies.asp?msg=物流公司已添加"
        Response.End
    End If
ElseIf scAction = "toggle" Then
    Dim scId : scId = Request.Form("companyId")
    If IsNumeric(scId) Then
        conn.Execute "UPDATE ShippingCompanies SET IsActive = CASE WHEN IsActive=1 THEN 0 ELSE 1 END, UpdatedAt=GETDATE() WHERE CompanyID=" & CLng(scId)
        Response.Redirect "shipping_companies.asp?msg=状态已更新"
        Response.End
    End If
ElseIf scAction = "update" Then
    scId = Request.Form("companyId")
    scContact = Replace(Request.Form("contactPerson"),"'","''")
    scPhone = Replace(Request.Form("contactPhone"),"'","''")
    scNotes = Replace(Request.Form("notes"),"'","''")
    If IsNumeric(scId) Then
        conn.Execute "UPDATE ShippingCompanies SET ContactPerson='" & scContact & "', ContactPhone='" & scPhone & "', Notes='" & scNotes & "', UpdatedAt=GETDATE() WHERE CompanyID=" & CLng(scId)
        Response.Redirect "shipping_companies.asp?msg=信息已更新"
        Response.End
    End If
End If

Dim scTotal, scActive
scTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM ShippingCompanies"))
scActive = SafeNum(GetScalar("SELECT COUNT(*) FROM ShippingCompanies WHERE IsActive=1"))

Dim rsSC
Set rsSC = conn.Execute("SELECT * FROM ShippingCompanies ORDER BY CompanyName")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>物流公司管理 - 物流管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #00BCD4; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: var(--accent); }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 10px; padding: 16px; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .stat-label { font-size: 11px; color: #888; margin-bottom: 6px; }
        .stat-card .stat-value { font-size: 24px; font-weight: 700; color: var(--accent); }
        .card-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 15px; }
        .company-card { background: linear-gradient(135deg, #2d2d44, #1e1e32); border-radius: 12px; padding: 20px; border: 1px solid rgba(255,255,255,0.06); transition: all 0.2s; position: relative; }
        .company-card:hover { border-color: rgba(0,188,212,0.3); }
        .company-card.inactive { opacity: 0.5; }
        .company-card .cc-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px; }
        .company-card .cc-name { font-size: 16px; font-weight: 600; color: #e0e0e0; }
        .company-card .cc-info { font-size: 12px; color: #888; line-height: 1.8; }
        .company-card .cc-info i { width: 16px; color: var(--accent); margin-right: 6px; }
        .company-card .cc-actions { margin-top: 12px; display: flex; gap: 8px; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 8px; font-size: 10px; font-weight: 500; }
        .badge-active { background: rgba(76,175,80,0.15); color: #81c784; }
        .badge-inactive { background: rgba(158,158,158,0.15); color: #9e9e9e; }
        .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 999; align-items: center; justify-content: center; }
        .modal.active { display: flex; }
        .modal-content { background: #2d2d44; border-radius: 12px; padding: 24px; width: 450px; max-width: 90vw; border: 1px solid rgba(255,255,255,0.1); }
        .modal-title { font-size: 18px; color: #e0e0e0; margin-bottom: 16px; }
        .form-group { margin-bottom: 14px; }
        .form-group label { display: block; font-size: 12px; color: #888; margin-bottom: 5px; }
        .form-group input, .form-group textarea { width: 100%; padding: 9px 12px; background: var(--input-bg); border: 1px solid rgba(255,255,255,0.12); border-radius: 6px; color: #e0e0e0; font-size: 13px; }
        .form-group textarea { resize: vertical; min-height: 60px; }
        .form-group input:focus, .form-group textarea:focus { outline: none; border-color: var(--accent); }
        .modal-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px; }
        .alert-msg { padding: 10px 16px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .alert-success { background: rgba(76,175,80,0.12); color: #81c784; border: 1px solid rgba(76,175,80,0.2); }
</style>
</head>
<body>
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-building"></i> 物流公司管理</h2>
            <button class="btn btn-primary" onclick="document.getElementById('addModal').classList.add('active')"><i class="fas fa-plus"></i> 添加公司</button>
        </div>

        <% Dim scMsg : scMsg = Request.QueryString("msg")
        If scMsg <> "" Then %>
        <div class="alert-msg alert-success"><i class="fas fa-check-circle"></i> <%=Server.HTMLEncode(scMsg)%></div>
        <% End If %>

        <div class="stats-grid">
            <div class="stat-card"><div class="stat-label"><i class="fas fa-building"></i> 物流公司</div><div class="stat-value"><%=scTotal%></div></div>
            <div class="stat-card"><div class="stat-label"><i class="fas fa-check-circle"></i> 合作中</div><div class="stat-value"><%=scActive%></div></div>
        </div>

        <div class="card-grid">
            <% If Not rsSC Is Nothing Then
                Do While Not rsSC.EOF
                    Dim cId, cName, cContact, cPhone, cWeb, cActive, cNotes, cCreated
                    cId = rsSC("CompanyID")
                    cName = rsSC("CompanyName") & ""
                    cContact = rsSC("ContactPerson") & ""
                    cPhone = rsSC("ContactPhone") & ""
                    cWeb = rsSC("Website") & ""
                    cActive = rsSC("IsActive")
                    cNotes = rsSC("Notes") & ""
            %>
            <div class="company-card <%=IIf(cActive,"","inactive")%>">
                <div class="cc-header">
                    <span class="cc-name"><%=Server.HTMLEncode(cName)%></span>
                    <span class="badge <%=IIf(cActive,"badge-active","badge-inactive")%>"><%=IIf(cActive,"合作中","已停用")%></span>
                </div>
                <div class="cc-info">
                    <% If cContact <> "" Then %><div><i class="fas fa-user"></i> <%=Server.HTMLEncode(cContact)%></div><% End If %>
                    <% If cPhone <> "" Then %><div><i class="fas fa-phone"></i> <%=Server.HTMLEncode(cPhone)%></div><% End If %>
                    <% If cWeb <> "" Then %><div><i class="fas fa-globe"></i> <%=Server.HTMLEncode(cWeb)%></div><% End If %>
                    <% If cNotes <> "" Then %><div><i class="fas fa-sticky-note"></i> <%=Server.HTMLEncode(cNotes)%></div><% End If %>
                </div>
                <div class="cc-actions">
                    <button class="btn btn-outline" style="font-size:11px;" onclick="openEdit(<%=cId%>,'<%=Server.HTMLEncode(Replace(cContact,"'","\'"))%>','<%=Server.HTMLEncode(Replace(cPhone,"'","\'"))%>','<%=Server.HTMLEncode(Replace(cNotes,"'","\'"))%>')"><i class="fas fa-edit"></i> 编辑</button>
                    <form method="post" style="display:inline;">
                        <input type="hidden" name="action" value="toggle">
                        <input type="hidden" name="companyId" value="<%=cId%>">
                        <button type="submit" class="btn btn-danger" style="font-size:11px;"><i class="fas <%=IIf(cActive,"fa-ban","fa-check")%>"></i> <%=IIf(cActive,"停用","启用")%></button>
                    </form>
                </div>
            </div>
            <%      rsSC.MoveNext
                Loop
                rsSC.Close : Set rsSC = Nothing
            End If %>
        </div>
    </div>

    <!-- 添加 Modal -->
    <div id="addModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-plus-circle"></i> 添加物流公司</div>
            <form method="post">
                <input type="hidden" name="action" value="add">
                <div class="form-group"><label>公司名称 *</label><input type="text" name="companyName" required></div>
                <div class="form-group"><label>联系人</label><input type="text" name="contactPerson"></div>
                <div class="form-group"><label>联系电话</label><input type="text" name="contactPhone"></div>
                <div class="form-group"><label>网址</label><input type="text" name="website" placeholder="https://..."></div>
                <div class="form-group"><label>备注</label><textarea name="notes"></textarea></div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="document.getElementById('addModal').classList.remove('active')">取消</button>
                    <button type="submit" class="btn btn-primary">添加</button>
                </div>
            </form>
        </div>
    </div>

    <!-- 编辑 Modal -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <div class="modal-title"><i class="fas fa-edit"></i> 编辑联系信息</div>
            <form method="post">
                <input type="hidden" name="action" value="update">
                <input type="hidden" name="companyId" id="editCompanyId">
                <div class="form-group"><label>联系人</label><input type="text" name="contactPerson" id="editContact"></div>
                <div class="form-group"><label>联系电话</label><input type="text" name="contactPhone" id="editPhone"></div>
                <div class="form-group"><label>备注</label><textarea name="notes" id="editNotes"></textarea></div>
                <div class="modal-actions">
                    <button type="button" class="btn btn-outline" onclick="document.getElementById('editModal').classList.remove('active')">取消</button>
                    <button type="submit" class="btn btn-primary">保存</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    function openEdit(id, contact, phone, notes) {
        document.getElementById('editCompanyId').value = id;
        document.getElementById('editContact').value = contact;
        document.getElementById('editPhone').value = phone;
        document.getElementById('editNotes').value = notes;
        document.getElementById('editModal').classList.add('active');
    }
    </script>
</body>
</html>
<%
Call CloseConnection()
%>
