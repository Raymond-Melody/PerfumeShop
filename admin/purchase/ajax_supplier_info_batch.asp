<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<% Option Explicit %>
<!--#include file="../../includes/connection.asp"-->
<%
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

Call OpenConnection()

Function SafeNum(val)
    If IsNull(val) Or val = "" Or Not IsNumeric(val) Then SafeNum = 0 Else SafeNum = CDbl(val)
End Function

Function SafeSQL(str)
    If IsNull(str) Or str = "" Then SafeSQL = "" Else SafeSQL = Replace(str, "'", "''")
End Function

Dim itemCodesStr, orderType, results, i
itemCodesStr = Trim(Request.Form("itemcodes"))
If itemCodesStr = "" Then itemCodesStr = Trim(Request.QueryString("itemcodes"))
orderType = Trim(Request.Form("ordertype"))
If orderType = "" Then orderType = Trim(Request.QueryString("ordertype"))
If orderType = "" Then orderType = "RawMaterial"

If itemCodesStr = "" Then
    Response.Write "[]"
    Call CloseConnection()
    Response.End
End If

Dim itemCodes : itemCodes = Split(itemCodesStr, ",")
Dim codeCount : codeCount = UBound(itemCodes) + 1

On Error Resume Next

' 构建 JSON 数组
Dim jsonParts : jsonParts = ""
Dim inClause : inClause = ""

' 构建 IN 子句
For i = 0 To UBound(itemCodes)
    Dim code : code = Trim(itemCodes(i))
    If code <> "" Then
        If inClause <> "" Then inClause = inClause & ","
        inClause = inClause & "'" & SafeSQL(code) & "'"
    End If
Next

If inClause = "" Then
    Response.Write "[]"
    Call CloseConnection()
    Response.End
End If

' 批量查询 SupplierPrices
Dim rs, sql
sql = "SELECT sp.ItemCode, sp.SupplierID, s.SupplierName, sp.UnitPrice " & _
      "FROM SupplierPrices sp " & _
      "LEFT JOIN Suppliers s ON sp.SupplierID = s.SupplierID " & _
      "WHERE sp.ItemCode IN (" & inClause & ") AND sp.IsActive = 1 " & _
      "ORDER BY sp.CreatedAt DESC"

Dim supplierMap : Set supplierMap = Server.CreateObject("Scripting.Dictionary")
supplierMap.CompareMode = 1  ' TextCompare - case insensitive

Set rs = conn.Execute(sql)
If Not rs Is Nothing Then
    Do While Not rs.EOF
        Dim keyItemCode : keyItemCode = rs("ItemCode") & ""
        If Not supplierMap.Exists(keyItemCode) Then
            supplierMap.Add keyItemCode, Array(CLng(SafeNum(rs("SupplierID"))), rs("SupplierName") & "", SafeNum(rs("UnitPrice")))
        End If
        rs.MoveNext
    Loop
    rs.Close
End If
Set rs = Nothing

' 处理结果
For i = 0 To UBound(itemCodes)
    code = Trim(itemCodes(i))
    If code <> "" Then
        Dim supID, supName, supPrice, found
        supID = 0 : supName = "" : supPrice = 0 : found = False
        
        If supplierMap.Exists(code) Then
            Dim info : info = supplierMap(code)
            supID = info(0) : supPrice = info(2) : supName = info(1)
            found = True
        End If
        
        If jsonParts <> "" Then jsonParts = jsonParts & ","
        jsonParts = jsonParts & "{""itemcode"":""" & Replace(Replace(code, "\", "\\"), """", "\""") & """,""supplier_id"":" & supID & ",""supplier_name"":""" & Replace(Replace(supName, "\", "\\"), """", "\""") & """,""unit_price"":" & Replace(FormatNumber(supPrice, 4, -1, 0, 0), ",", "") & "}"
    End If
Next

If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

Response.Write "[" & jsonParts & "]"

Set supplierMap = Nothing
Call CloseConnection()
%>
