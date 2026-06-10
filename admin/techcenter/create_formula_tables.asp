<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
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
    <title>创建配方相关表</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: #fff; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        h1 { color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; }
        .result { padding: 15px; margin: 10px 0; border-radius: 5px; border-left: 4px solid; }
        .success { border-left-color: #4CAF50; background: #e8f5e9; color: #2e7d32; }
        .exists { border-left-color: #ff9800; background: #fff3e0; color: #ef6c00; }
        .error { border-left-color: #f44336; background: #ffebee; color: #c62828; }
        .info { border-left-color: #2196F3; background: #e3f2fd; color: #1565c0; }
        code { background: #f5f5f5; padding: 2px 6px; border-radius: 3px; font-family: Consolas, monospace; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 10px; border: 1px solid #ddd; text-align: left; }
        th { background: #4CAF50; color: white; }
    </style>
</head>
<body>
<div class="container">
    <h1>配方相关表创建工具</h1>
    
    <div class="info result">
        <strong>说明：</strong>本脚本用于创建 Formulas（配方主表）和 FormulaNotes（配方-香调关联表）。<br>
        如果表已存在，将自动跳过创建。
    </div>
    
    <h2>执行结果</h2>
    
    <% 
    Dim result, errorMsg
    
    ' ========== 创建 Formulas 表 ==========
    Response.Write "<h3>1. Formulas 表（配方主表）</h3>"
    
    If TableExists("Formulas") Then
        Response.Write "<div class='exists result'><strong>已存在</strong> - Formulas 表已存在，跳过创建</div>"
    Else
        Dim sqlFormulas
        sqlFormulas = "CREATE TABLE Formulas (" & _
                      "FormulaID INT IDENTITY(1,1) PRIMARY KEY, " & _
                      "FormulaName NVARCHAR(100) NOT NULL, " & _
                      "Description NVARCHAR(MAX), " & _
                      "IsActive BIT DEFAULT 1, " & _
                      "CreatedAt DATETIME, " & _
                      "UpdatedAt DATETIME)"
        
        If CreateTable(sqlFormulas, "Formulas") Then
            Response.Write "<div class='success result'><strong>创建成功</strong> - Formulas 表创建完成</div>"
            Response.Write "<table>"
            Response.Write "<tr><th>字段名</th><th>类型</th><th>说明</th></tr>"
            Response.Write "<tr><td>FormulaID</td><td>COUNTER</td><td>主键，自增</td></tr>"
            Response.Write "<tr><td>FormulaName</td><td>VARCHAR(100)</td><td>配方名称，非空</td></tr>"
            Response.Write "<tr><td>Description</td><td>MEMO</td><td>配方描述</td></tr>"
            Response.Write "<tr><td>IsActive</td><td>YESNO</td><td>是否启用，默认TRUE</td></tr>"
            Response.Write "<tr><td>CreatedAt</td><td>DATETIME</td><td>创建时间</td></tr>"
            Response.Write "<tr><td>UpdatedAt</td><td>DATETIME</td><td>更新时间</td></tr>"
            Response.Write "</table>"
        Else
            Response.Write "<div class='error result'><strong>创建失败</strong> - " & Server.HTMLEncode(Session("LastDBError")) & "</div>"
        End If
    End If
    
    ' ========== 创建 FormulaNotes 表 ==========
    Response.Write "<h3>2. FormulaNotes 表（配方-香调关联表）</h3>"
    
    If TableExists("FormulaNotes") Then
        Response.Write "<div class='exists result'><strong>已存在</strong> - FormulaNotes 表已存在，跳过创建</div>"
    Else
        Dim sqlFormulaNotes
        sqlFormulaNotes = "CREATE TABLE FormulaNotes (" & _
                          "ID INT IDENTITY(1,1) PRIMARY KEY, " & _
                          "FormulaID INT NOT NULL, " & _
                          "NoteID INT NOT NULL, " & _
                          "Percentage INT DEFAULT 0)"
        
        If CreateTable(sqlFormulaNotes, "FormulaNotes") Then
            Response.Write "<div class='success result'><strong>创建成功</strong> - FormulaNotes 表创建完成</div>"
            Response.Write "<table>"
            Response.Write "<tr><th>字段名</th><th>类型</th><th>说明</th></tr>"
            Response.Write "<tr><td>ID</td><td>COUNTER</td><td>主键，自增</td></tr>"
            Response.Write "<tr><td>FormulaID</td><td>INT</td><td>配方ID，非空</td></tr>"
            Response.Write "<tr><td>NoteID</td><td>INT</td><td>香调ID，非空</td></tr>"
            Response.Write "<tr><td>Percentage</td><td>INT</td><td>占比百分比，默认0</td></tr>"
            Response.Write "</table>"
        Else
            Response.Write "<div class='error result'><strong>创建失败</strong> - " & Server.HTMLEncode(Session("LastDBError")) & "</div>"
        End If
    End If
    %>
    
    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd;">
        <a href="index.asp" class="btn">返回技术中心</a>
        <a href="formula_management.asp" class="btn btn--info">配方管理</a>
    </div>
    
</div>
</body>
</html>
<%
Call CloseConnection()
%>
