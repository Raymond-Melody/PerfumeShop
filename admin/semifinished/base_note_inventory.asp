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
            If Not rs.EOF Then val = rs(0)
            If IsNull(val) Then val = 0
            rs.Close
        End If
    Else : Err.Clear
    End If
    Set rs = Nothing
    GetScalar = val
End Function

Dim bnTotal, bnActive, bnWithNotes
bnTotal = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes"))
bnActive = SafeNum(GetScalar("SELECT COUNT(*) FROM BaseNotes WHERE IsActive=1"))
bnWithNotes = SafeNum(GetScalar("SELECT COUNT(DISTINCT bn.BaseNoteID) FROM BaseNotes bn INNER JOIN FragranceNotes fn ON bn.BaseNoteID=fn.BaseNoteID"))

Dim rsBaseNotes
Set rsBaseNotes = conn.Execute("SELECT bn.*, (SELECT COUNT(*) FROM FragranceNotes WHERE BaseNoteID=bn.BaseNoteID) AS NoteCount, (SELECT STRING_AGG(NoteName,', ') FROM FragranceNotes WHERE BaseNoteID=bn.BaseNoteID) AS NoteNames FROM BaseNotes bn ORDER BY bn.IsActive DESC, bn.BaseNoteName ASC")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>基香库存 - 半成品生产中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        :root { --bg: #1a1a2e; --text: #e0e0e0; --accent: #2196F3; --input-bg: #2d2d44; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1a2e; color: var(--text); min-height: 100vh; }
        .main-content { padding: 30px; margin-left: 260px; }
        .page-header { margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .page-title { font-size: 24px; color: #e0e0e0; display: flex; align-items: center; gap: 10px; }
        .page-title i { color: #2196F3; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 20px; border-radius: 12px; text-align: center; border: 1px solid rgba(255,255,255,0.06); }
        .stat-card .num { font-size: 28px; font-weight: bold; display: block; }
        .stat-card .label { font-size: 12px; color: #888; display: block; margin-top: 5px; }
        
        .card { background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); border-radius: 12px; margin-bottom: 25px; border: 1px solid rgba(255,255,255,0.06); }
        .card-header { padding: 14px 20px; background: rgba(33,150,243,0.08); border-bottom: 1px solid rgba(255,255,255,0.06); font-weight: 600; font-size: 15px; color: #e0e0e0; }
        .card-body { padding: 16px 20px; overflow-x: auto; }
        
        table { width: 100%; border-collapse: collapse; }
        th { background: rgba(33,150,243,0.15); color: #64b5f6; padding: 12px 15px; text-align: left; font-weight: 600; font-size: 13px; }
        td { padding: 12px 15px; border-bottom: 1px solid rgba(255,255,255,0.04); color: #e0e0e0; font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        
        .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; }
        .status-active { background: rgba(76,175,80,0.15); color: #81c784; }
        .status-inactive { background: rgba(244,67,54,0.15); color: #e57373; }
        
        .ingredient-tags { display: flex; flex-wrap: wrap; gap: 4px; }
        .ingredient-tag { background: rgba(33,150,243,0.12); color: #64b5f6; padding: 2px 10px; border-radius: 10px; font-size: 12px; }
        
        .text-center { text-align: center; }
        .text-muted { color: #888; font-size: 13px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-database"></i> 基香库存</h2>
        </div>
        
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card"><span class="num" style="color:#2196F3;"><%=bnTotal%></span><span class="label">基香总数</span></div>
            <div class="stat-card"><span class="num" style="color:#4CAF50;"><%=bnActive%></span><span class="label">已激活</span></div>
            <div class="stat-card"><span class="num" style="color:#FF9800;"><%=bnWithNotes%></span><span class="label">已关联香调</span></div>
        </div>
        
        <!-- 基香库存列表 -->
        <div class="card">
            <div class="card-header">基香库存清单</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>基香名称</th><th>成分</th><th>关联香调</th><th>状态</th><th>描述</th></tr></thead>
                    <tbody>
                    <%
                    If Not rsBaseNotes Is Nothing Then
                        Dim bnRowCount : bnRowCount = 0
                        Do While Not rsBaseNotes.EOF
                            bnRowCount = bnRowCount + 1
                            Dim bnID, bnName, bnActiveFlag, bnCount, bnNoteNames, bnDesc, bnIngredients
                            bnID = rsBaseNotes("BaseNoteID")
                            bnName = CStr(rsBaseNotes("BaseNoteName") & "")
                            bnActiveFlag = SafeNum(rsBaseNotes("IsActive"))
                            bnCount = SafeNum(rsBaseNotes("NoteCount"))
                            bnNoteNames = CStr(rsBaseNotes("NoteNames") & "")
                            bnDesc = CStr(rsBaseNotes("Description") & "")
                            bnIngredients = CStr(rsBaseNotes("Ingredients") & "")
                    %>
                        <tr>
                            <td><strong><%=Server.HTMLEncode(bnName)%></strong></td>
                            <td>
                                <div class="ingredient-tags">
                                <%
                                    Dim bnIngArr, bnIngItem
                                    bnIngArr = Split(bnIngredients, ",")
                                    For Each bnIngItem In bnIngArr
                                        bnIngItem = Trim(bnIngItem)
                                        If bnIngItem <> "" Then
                                %>
                                    <span class="ingredient-tag"><%=Server.HTMLEncode(bnIngItem)%></span>
                                <%      End If
                                    Next
                                %>
                                </div>
                            </td>
                            <td>
                                <% If bnCount > 0 Then %>
                                <span style="color:#64b5f6;"><%=bnCount%>个: <%=Server.HTMLEncode(bnNoteNames)%></span>
                                <% Else %>
                                <span class="text-muted">未关联</span>
                                <% End If %>
                            </td>
                            <td><span class="status-badge <%=IIF(bnActiveFlag=1,"status-active","status-inactive")%>"><%=IIF(bnActiveFlag=1,"激活","停用")%></span></td>
                            <td class="text-muted"><%=IIF(bnDesc<>"",Server.HTMLEncode(bnDesc),"-")%></td>
                        </tr>
                    <%
                            rsBaseNotes.MoveNext
                        Loop
                        rsBaseNotes.Close
                    End If
                    Set rsBaseNotes = Nothing
                    If bnRowCount = 0 Then
                    %>
                        <tr><td colspan="5" class="text-center text-muted" style="padding:40px;">暂无基香数据</td></tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
