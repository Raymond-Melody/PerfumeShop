<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' ============================================
' V14.6 采购订单 - 主调度器
' 原文件 128KB/2707行 → 调度器 + 8个子模块
' ============================================
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
%>
<!--#include file="includes/auth.asp"-->
<!--#include file="../../includes/config.asp"-->
<!--#include file="../../includes/common_utils.asp"-->
<%
Call OpenConnection()
%>
<!--#include file="purchase_schema_migration.asp"-->
<!--#include file="purchase_utility_functions.asp"-->
<!--#include file="purchase_order_post_handler.asp"-->
<!--#include file="purchase_order_data_loader.asp"-->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>采购订单管理 - 采购中心</title>
    <link rel="stylesheet" href="/css/admin.css">
    <link rel="stylesheet" href="/admin/css/admin-theme.css">
    <link rel="stylesheet" href="/css/design-tokens.css">
    <link rel="stylesheet" href="/css/buttons.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
</head>
<body data-theme="purchase-dark">
    <!--#include file="includes/nav.asp"-->
    
    <div class="main-content">
        <div class="page-header">
            <h2 class="page-title"><i class="fas fa-file-invoice"></i> 采购订单管理</h2>
            <div class="breadcrumb">
                <a href="index.asp">采购中心</a> / <span>采购订单</span>
            </div>
        </div>
        
        <% If message <> "" Then %>
        <div class="message <%= messageType %>">
            <i class="fas fa-<%= IIf(messageType="success", "check-circle", "exclamation-circle") %>"></i>
            <%= message %>
        </div>
        <% End If %>
        
        <% 
        ' 调试信息：显示数据库错误详情（仅管理员可见）
        Dim dbgErr
        dbgErr = Session("LastDBError")
        If dbgErr <> "" Then
        %>
        <div class="message error">
            <i class="fas fa-bug"></i>
            <span style="font-size:12px;">调试: <%= Server.HTMLEncode(dbgErr) %></span>
        </div>
        <% 
            Session("LastDBError") = ""
        End If
        %>
        
        <!-- 查看订单详情（子模块） -->
        <!--#include file="purchase_order_view.asp"-->
        
        <!-- 新建/编辑表单（子模块） -->
        <!--#include file="purchase_order_form.asp"-->
        
        <!-- 筛选栏 + 订单列表 + 模态框（子模块） -->
        <!--#include file="purchase_order_list.asp"-->
    </div>

    <!-- 服务端注入变量（必须在外部JS之前） -->
    <script nonce="<%= Session("csp_nonce") %>">
        var baseNotesData = [
        <%
        If Not rsBaseNotes Is Nothing Then
            Dim bnFirst : bnFirst = True
            Do While Not rsBaseNotes.EOF
                If Not bnFirst Then Response.Write ","
                bnFirst = False
                Dim bnDesc : bnDesc = CStr(rsBaseNotes("Description") & "")
                Dim bnName : bnName = CStr(rsBaseNotes("BaseNoteName") & "")
                Dim bnNameJS : bnNameJS = Replace(bnName, Chr(34), "\" & Chr(34))
                Dim bnDescJS : bnDescJS = Replace(bnDesc, Chr(34), "\" & Chr(34))
        %>
            {id:<%=rsBaseNotes("BaseNoteID")%>, name:"<%=bnNameJS%>", desc:"<%=bnDescJS%>", price:<%=SafeNum(rsBaseNotes("UnitPrice"))%>}
        <%
                rsBaseNotes.MoveNext
            Loop
            rsBaseNotes.Close
            Set rsBaseNotes = Nothing
        End If
        %>
        ];
    </script>
    <!-- 公共脚本 -->
    <script src="/admin/js/admin-common.js"></script>
    <!-- 页面脚本 -->
    <script src="/admin/js/purchase-orders.js"></script>
</body>
</html>
<%
Call CloseConnection()
%>
