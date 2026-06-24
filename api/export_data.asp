<%@ Language="VBScript" CodePage="65001" %>
<%
' ============================================
' V16.0 数据导出工具 (Data Export)
' 支持: 订单导出CSV、财务报表导出
' ============================================
Response.Charset = "UTF-8"
%>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/connection.asp"-->
<!--#include file="../includes/audit_utils.asp"-->
<%
' 权限检查
If Session("AdminID") = "" Then
    Response.Write "请先登录管理后台"
    Response.End
End If

Call OpenConnection()

' V16: 确保审计日志表存在
Call EnsureAuditLogTable()

Dim exportType, dateFrom, dateTo
exportType = Trim(Request.QueryString("type"))
dateFrom = Trim(Request.QueryString("from"))
dateTo = Trim(Request.QueryString("to"))

If dateFrom = "" Then dateFrom = DateAdd("d", -30, Date())
If dateTo = "" Then dateTo = Date()

Select Case LCase(exportType)
    Case "orders"
        Call AuditLog(AUDIT_ACTION_EXPORT, AUDIT_TARGET_ORDER, 0, "订单导出", "日期范围: " & dateFrom & " ~ " & dateTo)
        ExportOrdersCSV dateFrom, dateTo
    Case "revenue"
        Call AuditLog(AUDIT_ACTION_EXPORT, AUDIT_TARGET_FINANCE, 0, "营收导出", "日期范围: " & dateFrom & " ~ " & dateTo)
        ExportRevenueCSV dateFrom, dateTo
    Case "customers"
        Call AuditLog(AUDIT_ACTION_EXPORT, AUDIT_TARGET_USER, 0, "客户导出", "日期范围: " & dateFrom & " ~ " & dateTo)
        ExportCustomersCSV dateFrom, dateTo
    Case "products"
        Call AuditLog(AUDIT_ACTION_EXPORT, AUDIT_TARGET_PRODUCT, 0, "产品导出", "全量导出")
        ExportProductsCSV
    Case Else
        Response.Write "无效的导出类型。支持: orders, revenue, customers, products"
End Select

Call CloseConnection()

' ============================================
' 导出订单CSV
' ============================================
Sub ExportOrdersCSV(dateFrom, dateTo)
    Response.ContentType = "text/csv; charset=UTF-8"
    Response.AddHeader "Content-Disposition", "attachment; filename=orders_" & FormatDate(dateFrom) & "_" & FormatDate(dateTo) & ".csv"
    
    ' BOM for Excel UTF-8
    Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)
    
    ' 表头
    Response.Write "订单号,客户,邮箱,金额,状态,支付方式,创建时间,发货时间" & vbCrLf
    
    Dim rs
    Set rs = conn.Execute("SELECT o.OrderNo, u.Username, u.Email, o.TotalAmount, o.Status, " & _
        "o.PaymentMethod, o.CreatedAt, o.ShippedAt FROM Orders o LEFT JOIN Users u ON o.UserID=u.UserID " & _
        "WHERE o.CreatedAt BETWEEN '" & SafeSQL(dateFrom) & "' AND '" & SafeSQL(dateTo) & " 23:59:59' ORDER BY o.CreatedAt DESC")
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            Response.Write """" & rs("OrderNo") & """,""" & rs("Username") & """,""" & rs("Email") & """,""" & rs("TotalAmount") & """,""" & rs("Status") & """,""" & rs("PaymentMethod") & """,""" & rs("CreatedAt") & """,""" & rs("ShippedAt") & """" & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
End Sub

' ============================================
' 导出营收CSV
' ============================================
Sub ExportRevenueCSV(dateFrom, dateTo)
    Response.ContentType = "text/csv; charset=UTF-8"
    Response.AddHeader "Content-Disposition", "attachment; filename=revenue_" & FormatDate(dateFrom) & "_" & FormatDate(dateTo) & ".csv"
    Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)
    
    Response.Write "日期,订单数,总营收,平均客单价" & vbCrLf
    
    Dim rs
    Set rs = conn.Execute("SELECT CAST(o.CreatedAt AS DATE) AS OrderDate, COUNT(*) AS Cnt, " & _
        "ISNULL(SUM(o.TotalAmount),0) AS Amt, ISNULL(AVG(o.TotalAmount),0) AS AvgAmt " & _
        "FROM Orders o WHERE o.CreatedAt BETWEEN '" & SafeSQL(dateFrom) & "' AND '" & SafeSQL(dateTo) & " 23:59:59' " & _
        "AND o.Status<>'Cancelled' GROUP BY CAST(o.CreatedAt AS DATE) ORDER BY OrderDate")
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            Response.Write """" & rs("OrderDate") & """," & rs("Cnt") & "," & rs("Amt") & "," & rs("AvgAmt") & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
End Sub

' ============================================
' 导出客户CSV
' ============================================
Sub ExportCustomersCSV(dateFrom, dateTo)
    Response.ContentType = "text/csv; charset=UTF-8"
    Response.AddHeader "Content-Disposition", "attachment; filename=customers_" & FormatDate(dateFrom) & "_" & FormatDate(dateTo) & ".csv"
    Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)
    
    Response.Write "用户名,邮箱,姓名,手机,注册时间,订单数,累计消费" & vbCrLf
    
    Dim rs
    Set rs = conn.Execute("SELECT u.Username, u.Email, u.FullName, u.Phone, u.CreatedAt, " & _
        "COUNT(o.OrderID) AS OrderCnt, ISNULL(SUM(o.TotalAmount),0) AS TotalSpent " & _
        "FROM Users u LEFT JOIN Orders o ON u.UserID=o.UserID AND o.Status<>'Cancelled' " & _
        "WHERE u.CreatedAt BETWEEN '" & SafeSQL(dateFrom) & "' AND '" & SafeSQL(dateTo) & " 23:59:59' " & _
        "GROUP BY u.Username, u.Email, u.FullName, u.Phone, u.CreatedAt ORDER BY u.CreatedAt DESC")
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            Response.Write """" & rs("Username") & """,""" & rs("Email") & """,""" & rs("FullName") & """,""" & rs("Phone") & """,""" & rs("CreatedAt") & """," & rs("OrderCnt") & "," & rs("TotalSpent") & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
End Sub

' ============================================
' 导出产品CSV
' ============================================
Sub ExportProductsCSV()
    Response.ContentType = "text/csv; charset=UTF-8"
    Response.AddHeader "Content-Disposition", "attachment; filename=products_" & FormatDate(Date()) & ".csv"
    Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)
    
    Response.Write "产品ID,产品名称,类型,基础价格,是否活跃,库存,创建时间" & vbCrLf
    
    Dim rs
    Set rs = conn.Execute("SELECT ProductID, ProductName, ProductType, BasePrice, IsActive, Stock, CreatedAt FROM Products ORDER BY ProductID")
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            Response.Write rs("ProductID") & ",""" & rs("ProductName") & """,""" & rs("ProductType") & """," & rs("BasePrice") & "," & rs("IsActive") & "," & rs("Stock") & ",""" & rs("CreatedAt") & """" & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
End Sub

Function FormatDate(d)
    If IsDate(d) Then FormatDate = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) Else FormatDate = "unknown"
End Function
%>
