<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
' ============================================
' V21 集中站点参数配置中心
' 统一 CRUD SiteSettings（含 FEATURE_* 开关），带 CSRF + 操作级权限 + 审计
' 仅 SUPER_ADMIN 可访问（由 includes/auth.asp 的 VerifyModuleAccess("system") 保证）
' ============================================
Call OpenConnection()

Function SafeVal(v)
    If IsNull(v) Then SafeVal = "" Else SafeVal = CStr(v)
End Function

Dim msg, msgType
msg = "" : msgType = "success"

' ========== POST 处理 ==========
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not ValidateCSRFToken() Then
        msg = "安全验证失败，请刷新后重试" : msgType = "error"
    Else
        Call RequirePermissionOrDie("system", "edit")
        Dim postAct : postAct = Request.Form("action")

        If postAct = "save_setting" Then
            Dim sKey, sVal
            sKey = Trim(Request.Form("settingKey"))
            sVal = Trim(Request.Form("settingValue"))
            If sKey <> "" Then
                Dim existCnt
                existCnt = GetScalar("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey = '" & SafeSQL(sKey) & "'")
                If CLng("0" & existCnt) > 0 Then
                    ExecuteNonQuery "UPDATE SiteSettings SET SettingValue = '" & SafeSQL(sVal) & "', UpdatedAt = GETDATE() WHERE SettingKey = '" & SafeSQL(sKey) & "'"
                Else
                    ExecuteNonQuery "INSERT INTO SiteSettings (SettingKey, SettingValue, UpdatedAt) VALUES ('" & SafeSQL(sKey) & "', '" & SafeSQL(sVal) & "', GETDATE())"
                End If
                Call LogAdminAction("修改站点参数", "system", "SiteSettings", sKey, "值=" & sVal)
                msg = "参数 [" & sKey & "] 已保存"
            Else
                msg = "参数键不能为空" : msgType = "error"
            End If

        ElseIf postAct = "toggle_feature" Then
            Dim fKey, fOn
            fKey = Trim(Request.Form("featureKey"))
            fOn = IIf(Request.Form("featureOn") = "1", "1", "0")
            If fKey <> "" Then
                Dim fFull : fFull = fKey
                If Left(UCase(fKey), 8) <> "FEATURE_" Then fFull = "FEATURE_" & fKey
                Dim fExist : fExist = GetScalar("SELECT COUNT(*) FROM SiteSettings WHERE SettingKey='" & SafeSQL(fFull) & "'")
                If CLng("0" & fExist) > 0 Then
                    ExecuteNonQuery "UPDATE SiteSettings SET SettingValue='" & fOn & "', UpdatedAt=GETDATE() WHERE SettingKey='" & SafeSQL(fFull) & "'"
                Else
                    ExecuteNonQuery "INSERT INTO SiteSettings (SettingKey, SettingValue, UpdatedAt) VALUES ('" & SafeSQL(fFull) & "', '" & fOn & "', GETDATE())"
                End If
                Call LogAdminAction("切换功能开关", "system", "SiteSettings", fFull, "开=" & fOn)
                msg = "开关 [" & fFull & "] 已更新为 " & IIf(fOn = "1", "开启", "关闭")
            End If
        End If
    End If
End If

' ========== 参数分组定义 ==========
Function GroupOf(k)
    Dim uk : uk = UCase(k)
    If Left(uk, 8) = "FEATURE_" Then
        GroupOf = "功能开关"
    ElseIf InStr(uk, "PROMOTION") > 0 Or InStr(uk, "COUPON") > 0 Or InStr(uk, "POINTS") > 0 Then
        GroupOf = "营销促销"
    ElseIf InStr(uk, "SHIPPING") > 0 Or InStr(uk, "PLATFORMFEE") > 0 Or InStr(uk, "FREESHIP") > 0 Then
        GroupOf = "运费/平台费"
    ElseIf InStr(uk, "COST") > 0 Or InStr(uk, "PERCENT") > 0 Then
        GroupOf = "成本/配比"
    ElseIf InStr(uk, "ENABLE") > 0 Or InStr(uk, "PAY") > 0 Or InStr(uk, "ALIPAY") > 0 Or InStr(uk, "WECHAT") > 0 Then
        GroupOf = "支付/开关"
    ElseIf InStr(uk, "SHOW") > 0 Or InStr(uk, "SECTION") > 0 Then
        GroupOf = "首页展示"
    Else
        GroupOf = "其他"
    End If
End Function

' 读取全部设置
Dim rsAll
Set rsAll = ExecuteQuery("SELECT SettingKey, SettingValue, UpdatedAt FROM SiteSettings ORDER BY SettingKey")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>站点参数配置中心 - 系统管理</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body { background:#1a1a2e; color:#e0e0e0; }
        .main-content { padding:30px; margin-left:260px; }
        .page-title { font-size:24px; margin-bottom:20px; display:flex; align-items:center; gap:10px; }
        .page-title i { color:#00bcd4; }
        .alert { padding:12px 16px; border-radius:8px; margin-bottom:16px; }
        .alert-success { background:rgba(76,175,80,0.15); color:#81c784; }
        .alert-error { background:rgba(244,67,54,0.15); color:#ef9a9a; }
        .card { background:linear-gradient(135deg,#2d2d44,#1e1e32); border-radius:12px; margin-bottom:20px; border:1px solid rgba(255,255,255,0.06); }
        .card-header { padding:14px 20px; font-weight:600; border-bottom:1px solid rgba(255,255,255,0.06); color:#00bcd4; }
        .card-body { padding:16px 20px; overflow-x:auto; }
        table { width:100%; border-collapse:collapse; }
        th { text-align:left; padding:10px 12px; background:rgba(0,188,212,0.12); color:#4dd0e1; font-size:13px; }
        td { padding:8px 12px; border-bottom:1px solid rgba(255,255,255,0.04); font-size:13px; }
        input[type=text] { width:100%; padding:7px 10px; background:#1a1a2e; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#e0e0e0; }
        .btn { padding:6px 14px; border:none; border-radius:6px; background:#00bcd4; color:#fff; cursor:pointer; font-size:12px; }
        .add-form { display:flex; gap:10px; flex-wrap:wrap; align-items:center; }
        .add-form input { width:220px; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    <div class="main-content">
        <h2 class="page-title"><i class="fas fa-sliders-h"></i> 站点参数配置中心</h2>
        <% If msg <> "" Then %><div class="alert alert-<%= msgType %>"><%= Server.HTMLEncode(msg) %></div><% End If %>

        <!-- 新增参数 -->
        <div class="card">
            <div class="card-header">新增 / 快速修改参数</div>
            <div class="card-body">
                <form method="post" class="add-form">
                    <%= GetCSRFTokenField() %>
                    <input type="hidden" name="action" value="save_setting">
                    <input type="text" name="settingKey" placeholder="参数键 (如 Promotion_FirstOrder)" required>
                    <input type="text" name="settingValue" placeholder="参数值">
                    <button type="submit" class="btn"><i class="fas fa-save"></i> 保存</button>
                    <span style="color:#888;font-size:12px;">提示：功能开关请用 FEATURE_ 前缀，值 1/0</span>
                </form>
            </div>
        </div>

        <!-- 现有参数（分组） -->
        <div class="card">
            <div class="card-header">现有站点参数（<%= GroupOf("FEATURE_x") %> 等分组）</div>
            <div class="card-body">
                <table>
                    <thead><tr><th>分组</th><th>参数键</th><th>当前值</th><th>更新时间</th><th>操作</th></tr></thead>
                    <tbody>
                    <%
                    Dim rowN : rowN = 0
                    If Not rsAll Is Nothing Then
                        Do While Not rsAll.EOF
                            rowN = rowN + 1
                            Dim k, v
                            k = SafeVal(rsAll("SettingKey"))
                            v = SafeVal(rsAll("SettingValue"))
                    %>
                        <tr>
                            <td style="color:#888;"><%= GroupOf(k) %></td>
                            <td><strong><%= Server.HTMLEncode(k) %></strong></td>
                            <td>
                                <form method="post" style="display:flex;gap:6px;">
                                    <%= GetCSRFTokenField() %>
                                    <input type="hidden" name="action" value="save_setting">
                                    <input type="hidden" name="settingKey" value="<%= Server.HTMLEncode(k) %>">
                                    <input type="text" name="settingValue" value="<%= Server.HTMLEncode(v) %>">
                                    <button type="submit" class="btn">保存</button>
                                </form>
                            </td>
                            <td style="color:#888;"><%= IIf(IsNull(rsAll("UpdatedAt")), "-", Left(SafeVal(rsAll("UpdatedAt")), 19)) %></td>
                            <td style="color:#666;">—</td>
                        </tr>
                    <%
                            rsAll.MoveNext
                        Loop
                        rsAll.Close
                    End If
                    Set rsAll = Nothing
                    If rowN = 0 Then %>
                        <tr><td colspan="5" style="text-align:center;padding:30px;color:#888;">暂无参数</td></tr>
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
