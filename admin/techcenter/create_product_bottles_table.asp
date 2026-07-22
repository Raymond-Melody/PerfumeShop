<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/connection.asp"-->
<%
Call OpenConnection()

' ========== 检查表是否存在的函数 ==========
Function TableExists(tableName)
    Dim rsCheck, sqlCheck
    On Error Resume Next
    sqlCheck = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='" & tableName & "'"
    Set rsCheck = conn.Execute(sqlCheck)
    If Err.Number <> 0 Then
        TableExists = False
        Err.Clear
    Else
        If Not rsCheck.EOF Then
            TableExists = (CLng(rsCheck(0)) > 0)
        Else
            TableExists = False
        End If
        rsCheck.Close
        Set rsCheck = Nothing
    End If
    On Error GoTo 0
End Function

' ========== 创建表的函数 ==========
Function CreateTable(sql, tableName)
    On Error Resume Next
    conn.Execute sql
    If Err.Number <> 0 Then
        Session("LastDBError") = "创建表 " & tableName & " 失败: " & Err.Description
        CreateTable = False
        Err.Clear
    Else
        CreateTable = True
    End If
    On Error GoTo 0
End Function

' ========== 页面输出 ==========
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>创建产品瓶型关联表</title>
    <style>
        body { font-family: 'Segoe UI','Microsoft YaHei',sans-serif; padding: 0; margin: 0; background: #1a1a2e; color: #e0e0e0; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; }
        .container { max-width: 900px; margin: 0 auto; background: linear-gradient(135deg, #2d2d44 0%, #1e1e32 100%); padding: 25px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.06); box-shadow: 0 4px 20px rgba(0,0,0,0.3); }
        h1 { color: #fff; border-bottom: 2px solid rgba(0,188,212,0.3); padding-bottom: 10px; font-size: 20px; display: flex; align-items: center; gap: 10px; }
        h1 i { color: #00bcd4; }
        h3 { color: #e0e0e0; font-size: 14px; margin: 18px 0 8px; }
        .result { padding: 15px; margin: 10px 0; border-radius: 8px; border-left: 4px solid; }
        .success { border-left-color: #4CAF50; background: rgba(76,175,80,0.1); color: #81c784; }
        .exists { border-left-color: #ff9800; background: rgba(255,152,0,0.1); color: #ffb74d; }
        .error { border-left-color: #f44336; background: rgba(244,67,54,0.1); color: #ef9a9a; }
        .info { border-left-color: #2196F3; background: rgba(33,150,243,0.1); color: #90caf9; }
        code { background: rgba(255,255,255,0.08); padding: 2px 8px; border-radius: 4px; font-family: Consolas, 'Courier New', monospace; font-size: 13px; color: #80deea; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 10px 14px; border: 1px solid rgba(255,255,255,0.06); text-align: left; font-size: 13px; }
        th { background: rgba(0,188,212,0.12); color: #80deea; font-weight: 600; }
        td { color: #d0d6e0; }
        .btn-row { margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.06); display: flex; gap: 12px; flex-wrap: wrap; }
        .btn { display: inline-flex; align-items: center; gap: 6px; padding: 10px 18px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: 500; cursor: pointer; border: none; transition: all 0.2s; }
        .btn-primary { background: linear-gradient(135deg, #00bcd4, #00838f); color: #fff; }
        .btn-primary:hover { background: linear-gradient(135deg, #00acc1, #006064); transform: translateY(-1px); }
        .btn-outline { background: transparent; color: #00bcd4; border: 1px solid rgba(0,188,212,0.4); }
        .btn-outline:hover { background: rgba(0,188,212,0.1); }
        @media (max-width: 768px) {
            .main-content { margin-left: 0; padding: 15px; }
            .container { padding: 15px; }
        }
    </style>
</head>
<body data-theme="operation-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-database"></i> 产品瓶型关联表创建工具</h2>
            <div class="breadcrumb">
                <a href="index.asp">技术中心</a> / <span>瓶型表创建工具</span>
            </div>
        </div>
        
<div class="container">
    <h1><i class="fas fa-table"></i> 产品瓶型关联表创建工具</h1>
    
    <div class="info result">
        <strong>说明：</strong>本脚本用于创建 ProductBottleStyles（产品-瓶型关联表）。<br>
        如果表已存在，将自动跳过创建。
    </div>
    
    <h2>执行结果</h2>
    
    <% 
    Dim result, errorMsg
    
    ' ========== 创建 ProductBottleStyles 表 ==========
    Response.Write "<h3>1. ProductBottleStyles 表（产品-瓶型关联表）</h3>"
    
    If TableExists("ProductBottleStyles") Then
        Response.Write "<div class='exists result'><strong>已存在</strong> - ProductBottleStyles 表已存在，跳过创建</div>"
    Else
        Dim sqlProductBottleStyles
        sqlProductBottleStyles = "CREATE TABLE ProductBottleStyles (" & _
                      "ID INT IDENTITY(1,1) PRIMARY KEY, " & _
                      "ProductID INT NOT NULL, " & _
                      "BottleID INT NOT NULL, " & _
                      "CustomPrice MONEY)"
        
        If CreateTable(sqlProductBottleStyles, "ProductBottleStyles") Then
            Response.Write "<div class='success result'><strong>创建成功</strong> - ProductBottleStyles 表创建完成</div>"
            Response.Write "<table>"
            Response.Write "<tr><th>字段名</th><th>类型</th><th>说明</th></tr>"
            Response.Write "<tr><td>ID</td><td>COUNTER</td><td>主键，自增</td></tr>"
            Response.Write "<tr><td>ProductID</td><td>INT</td><td>产品ID，非空</td></tr>"
            Response.Write "<tr><td>BottleID</td><td>INT</td><td>瓶型ID，非空</td></tr>"
            Response.Write "<tr><td>CustomPrice</td><td>CURRENCY</td><td>自定义价格</td></tr>"
            Response.Write "</table>"
        Else
            Response.Write "<div class='error result'><strong>创建失败</strong> - " & Server.HTMLEncode(Session("LastDBError")) & "</div>"
        End If
    End If
    %>
    
    <div class="btn-row">
        <a href="index.asp" class="btn btn-primary"><i class="fas fa-home"></i> 返回技术中心</a>
        <a href="bottle_management.asp" class="btn btn-outline"><i class="fas fa-wine-bottle"></i> 瓶型管理</a>
    </div>
    
</div>
    </div>
    
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="../operation/css/operation-dark.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
</body>
</html>
<%
Call CloseConnection()
%>
