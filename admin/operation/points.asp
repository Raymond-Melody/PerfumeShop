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

' 处理积分调整
Dim action
action = Request.Form("action")

If action = "adjust_points" Then
    ' 验证CSRF令牌
    If Not ValidateCSRFToken() Then
        Response.Redirect "points.asp?error=安全验证失败"
        Response.End
    End If
    
    If Session("AdminRoleCode") = "OP_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then
        Dim userId, pointsChange, reason
        userId = Request.Form("userId")
        pointsChange = CInt(Request.Form("pointsChange"))
        reason = SafeSQL(Request.Form("reason"))
        
        ' 更新用户积分
        Dim updateResult
        updateResult = ExecuteNonQuery("UPDATE Users SET Points = Points + " & pointsChange & " WHERE UserID = " & userId)
        
        If updateResult Then
            ' 记录积分变动 - 优先使用ALTER TABLE添加的字段，失败则用原始字段
            Dim insertSql, insertOk
            Dim adminName
            adminName = SafeSQL(Session("AdminUserName") & "")
            
            ' 先尝试新字段名
            insertSql = "INSERT INTO PointTransactions (UserID, PointsChange, Reason, CreatedBy, CreatedAt) VALUES (" & _
                userId & ", " & pointsChange & ", '" & reason & "', '" & adminName & "', GETDATE())"
            insertOk = ExecuteNonQuery(insertSql)
            
            ' 如果失败，尝试原始字段名，同时设置PointsChange
            If Not insertOk Then
                insertSql = "INSERT INTO PointTransactions (UserID, Points, PointsChange, [Description], TransactionType, CreatedAt) VALUES (" & _
                    userId & ", " & pointsChange & ", " & pointsChange & ", '" & reason & "', 'ManualAdjust', GETDATE())"
                insertOk = ExecuteNonQuery(insertSql)
            End If
            
            ' 如果还失败，用最简字段
            If Not insertOk Then
                insertSql = "INSERT INTO PointTransactions (UserID, Points, CreatedAt) VALUES (" & _
                    userId & ", " & pointsChange & ", GETDATE())"
                insertOk = ExecuteNonQuery(insertSql)
            End If
            
            If Not insertOk Then
                ' 如果插入失败，回滚用户积分
                ExecuteNonQuery "UPDATE Users SET Points = Points - " & pointsChange & " WHERE UserID = " & userId
                Response.Redirect "points.asp?error=积分记录失败: " & Server.URLEncode(Session("LastDBError"))
                Response.End
            End If
            
            Call LogAdminAction("调整用户积分", "operation", "Users", userId, reason)
            Response.Redirect "points.asp?msg=积分调整成功"
        Else
            Response.Redirect "points.asp?error=积分更新失败: " & Server.URLEncode(Session("LastDBError"))
        End If
    Else
        Response.Redirect "points.asp?error=权限不足"
    End If
End If

' 预先获取所有统计数据（必须在打开Recordset前执行，Access不支持多活动记录集）
Dim totalPoints, todayTransactions, participantCount
totalPoints = GetScalar("SELECT SUM(Points) FROM Users")
If IsNull(totalPoints) Or totalPoints = "" Then totalPoints = 0 Else totalPoints = CDbl(totalPoints)

todayTransactions = GetScalar("SELECT COUNT(*) FROM PointTransactions WHERE CAST(CreatedAt AS DATE) = CAST(GETDATE() AS DATE)")
If IsNull(todayTransactions) Or todayTransactions = "" Then todayTransactions = 0

participantCount = GetScalar("SELECT COUNT(DISTINCT UserID) FROM PointTransactions")
If IsNull(participantCount) Or participantCount = "" Then participantCount = 0
participantCount = CLng("0" & participantCount)

' 预先获取用户列表（存入数组，关闭记录集后再用）
Dim userList()
Dim userCount
userCount = 0
Dim rsUsersTemp
Set rsUsersTemp = ExecuteQuery("SELECT UserID, Username, FullName FROM Users ORDER BY Username")
If Not rsUsersTemp Is Nothing Then
    Dim tempArr()
    ReDim tempArr(100, 2)
    Do While Not rsUsersTemp.EOF
        If userCount > UBound(tempArr, 1) Then
            ReDim Preserve tempArr(userCount + 50, 2)
        End If
        tempArr(userCount, 0) = rsUsersTemp("UserID")
        tempArr(userCount, 1) = rsUsersTemp("Username") & ""
        tempArr(userCount, 2) = rsUsersTemp("FullName") & ""
        userCount = userCount + 1
        rsUsersTemp.MoveNext
    Loop
    rsUsersTemp.Close
    Set rsUsersTemp = Nothing
End If

' 最后打开交易记录Recordset（保持打开状态用于HTML渲染）
Dim rsTransactions
Set rsTransactions = ExecuteQuery(_
    "SELECT TOP 50 pt.TransactionID, pt.UserID, pt.OrderID, " & _
    "IIF(ISNULL(pt.PointsChange), pt.Points, pt.PointsChange) AS PointsChange, " & _
    "IIF(ISNULL(pt.Reason), pt.Description, pt.Reason) AS Reason, " & _
    "pt.CreatedBy, pt.CreatedAt, u.Username, u.FullName AS RealName " & _
    "FROM PointTransactions pt " & _
    "LEFT JOIN Users u ON pt.UserID = u.UserID " & _
    "ORDER BY pt.CreatedAt DESC")

Call LogAdminAction("查看积分管理", "operation", "PointTransactions", "", "")
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>积分管理 - 运营管理中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="css/operation-dark.css">
        <link rel="stylesheet" href="/css/design-tokens.css">
        <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        .stats-cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); text-align: center; }
        .stat-card i { font-size: 36px; color: #667eea; margin-bottom: 10px; }
        .stat-card h3 { font-size: 32px; margin: 10px 0; color: #333; }
        .stat-card p { color: #666; margin: 0; }
        .adjust-form { background: white; padding: 25px; border-radius: 12px; margin-bottom: 25px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .form-row { display: flex; gap: 15px; margin-bottom: 15px; }
        .form-group { flex: 1; }
        .form-group label { display: block; margin-bottom: 8px; color: #555; font-weight: 500; }
        .form-group input, .form-group select { width: 100%; padding: 12px 15px; border: 2px solid #e0e0e0; border-radius: 8px; }
        .transactions-table { width: 100%; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
        .transactions-table th { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; text-align: left; }
        .transactions-table td { padding: 12px 15px; border-bottom: 1px solid #f0f0f0; }
        .points-positive { color: #4CAF50; font-weight: bold; }
        .points-negative { color: #f44336; font-weight: bold; }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-coins"></i> 积分管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">运营中心</a> / <span>积分管理</span>
            </div>
        </div>
        
        <div class="stats-cards">
            <div class="stat-card">
                <i class="fas fa-coins"></i>
                <h3><%= FormatNumber(totalPoints, 0) %></h3>
                <p>总发放积分</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-exchange-alt"></i>
                <h3><%= todayTransactions %></h3>
                <p>今日交易</p>
            </div>
            <div class="stat-card">
                <i class="fas fa-users"></i>
                <h3><%= participantCount %></h3>
                <p>参与用户</p>
            </div>
        </div>
        
        <% If Session("AdminRoleCode") = "OP_MANAGER" OR Session("AdminRoleCode") = "SUPER_ADMIN" Then %>
        <div class="adjust-form">
            <h3 style="margin-bottom: 20px;"><i class="fas fa-sliders-h"></i> 积分调整</h3>
            <form method="post" action="points.asp">
                <%= GetCSRFTokenField() %>
                <input type="hidden" name="action" value="adjust_points">
                <div class="form-row">
                    <div class="form-group">
                        <label>用户</label>
                        <select name="userId" required>
                            <option value="">选择用户</option>
                            <% 
                            Dim ui
                            For ui = 0 To userCount - 1
                            %>
                            <option value="<%= tempArr(ui, 0) %>"><%= Server.HTMLEncode(tempArr(ui, 1)) %> (<%= Server.HTMLEncode(tempArr(ui, 2)) %>)</option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>积分变动</label>
                        <input type="number" name="pointsChange" required placeholder="正数增加，负数扣减">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group" style="flex: 2;">
                        <label>调整原因</label>
                        <input type="text" name="reason" required placeholder="请输入积分调整原因">
                    </div>
                    <div class="form-group" style="display: flex; align-items: flex-end;">
                        <button type="submit" class="admin-btn admin-btn-primary"><i class="fas fa-save"></i> 确认调整</button>
                    </div>
                </div>
            </form>
        </div>
        <% End If %>
        
        <h3 style="margin-bottom: 15px;"><i class="fas fa-history"></i> 最近50条积分交易记录</h3>
        <table class="transactions-table">
            <thead>
                <tr>
                    <th>时间</th>
                    <th>用户</th>
                    <th>变动</th>
                    <th>原因</th>
                    <th>操作人</th>
                </tr>
            </thead>
            <tbody>
                <% If Not rsTransactions Is Nothing Then %>
                <% Do While Not rsTransactions.EOF %>
                <tr>
                    <td><%= rsTransactions("CreatedAt") %></td>
                    <td><%= rsTransactions("Username") %></td>
                    <td>
                        <% If rsTransactions("PointsChange") >= 0 Then %>
                        <span class="points-positive">+<%= rsTransactions("PointsChange") %></span>
                        <% Else %>
                        <span class="points-negative"><%= rsTransactions("PointsChange") %></span>
                        <% End If %>
                    </td>
                    <td><%= rsTransactions("Reason") %></td>
                    <td>
                        <% 
                        Dim opName
                        If Not IsNull(rsTransactions("CreatedBy")) And rsTransactions("CreatedBy") <> "" Then
                            ' CreatedBy 存储的是 AdminUserName，直接显示
                            opName = rsTransactions("CreatedBy")
                        Else
                            opName = "系统"
                        End If
                        Response.Write Server.HTMLEncode(opName)
                        %>
                    </td>
                </tr>
                <% rsTransactions.MoveNext %>
                <% Loop %>
                <% rsTransactions.Close %>
                <% End If %>
            </tbody>
        </table>
    </div>
</body>
</html>
<%
Call CloseConnection()
%>
