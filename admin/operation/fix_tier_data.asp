<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"

Dim conn, rs, sql

' Fix tier names using ChrW to construct correct Unicode
Dim correctNames(3,1)
correctNames(0,0) = "silver"  : correctNames(0,1) = ChrW(&H94F6) & ChrW(&H5361) & ChrW(&H4F1A) & ChrW(&H5458) ' 银卡会员
correctNames(1,0) = "gold"    : correctNames(1,1) = ChrW(&H91D1) & ChrW(&H5361) & ChrW(&H4F1A) & ChrW(&H5458) ' 金卡会员
correctNames(2,0) = "diamond" : correctNames(2,1) = ChrW(&H94BB) & ChrW(&H77F3) & ChrW(&H4F1A) & ChrW(&H5458) ' 钻石会员
correctNames(3,0) = "black"   : correctNames(3,1) = ChrW(&H9ED1) & ChrW(&H91D1) & ChrW(&H4F1A) & ChrW(&H5458) ' 黑金会员

%>
<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="UTF-8"><title>Fix Tier Names</title>
<style>body{font-family:monospace;background:#1a1a2e;color:#e0e0e0;padding:20px;}
h2{color:#00bcd4;}.ok{color:#4caf50;}.err{color:#f44336;}</style>
</head>
<body>
<h2>Fix MemberTiers TierName Data</h2>

<h3>Step 1: Show correct names constructed via ChrW</h3>
<%
Dim i
For i = 0 To 3
    Response.Write "<p>TierCode=" & correctNames(i,0) & " → Correct TierName=[" & correctNames(i,1) & "] Hex: "
    Dim j, ch
    For j = 1 To Len(correctNames(i,1))
        ch = AscW(Mid(correctNames(i,1), j, 1))
        Response.Write Right("0" & Hex(ch), 4) & " "
    Next
    Response.Write "</p>"
Next
%>

<h3>Step 2: Read current DB values</h3>
<%
Set conn = Server.CreateObject("ADODB.Connection")
' Use SQLOLEDB for reliable UPDATE persistence
conn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;Auto Translate=False;"
Set rs = conn.Execute("SELECT TierCode, TierName FROM MemberTiers ORDER BY SortOrder")
Response.Write "<table border=1 cellpadding=5><tr><th>TierCode</th><th>Current TierName</th><th>Hex</th></tr>"
Do While Not rs.EOF
    Dim rawName : rawName = rs("TierName") & ""
    Response.Write "<tr><td>" & rs("TierCode") & "</td><td>" & rawName & "</td><td>"
    For j = 1 To Len(rawName)
        Response.Write Right("0" & Hex(AscW(Mid(rawName, j, 1))), 4) & " "
    Next
    Response.Write "</td></tr>"
    rs.MoveNext
Loop
rs.Close
%>
</table>

<h3>Step 3: UPDATE with parameterized Command + explicit transaction</h3>
<%
Dim updated, errMsg, cmd, param
updated = 0
' Start explicit transaction
conn.BeginTrans
For i = 0 To 3
    On Error Resume Next
    Err.Clear
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = "UPDATE MemberTiers SET TierName = ?, UpdatedAt = GETDATE() WHERE TierCode = ?"
    cmd.CommandType = 1 ' adCmdText
    ' adVarWChar = 202 for Unicode NVARCHAR
    cmd.Parameters.Append cmd.CreateParameter("@TierName", 202, 1, 50, correctNames(i,1))
    ' adVarChar = 200 for VARCHAR
    cmd.Parameters.Append cmd.CreateParameter("@TierCode", 200, 1, 20, correctNames(i,0))
    cmd.Execute
    If Err.Number = 0 Then
        Response.Write "<p class='ok'>✓ Updated " & correctNames(i,0) & " → " & correctNames(i,1) & "</p>"
        updated = updated + 1
    Else
        Response.Write "<p class='err'>✗ Failed " & correctNames(i,0) & ": " & Err.Description & " (0x" & Hex(Err.Number) & ")</p>"
        Err.Clear
    End If
    Set cmd = Nothing
    On Error GoTo 0
Next
' Commit the transaction
If updated = 4 Then
    conn.CommitTrans
    Response.Write "<p class='ok'>✓ Transaction committed - all 4 updates persisted</p>"
Else
    conn.RollbackTrans
    Response.Write "<p class='err'>✗ Transaction rolled back - only " & updated & "/4 succeeded</p>"
End If
%>

<h3>Step 4: Verify after UPDATE</h3>
<%
Set rs = conn.Execute("SELECT TierCode, TierName FROM MemberTiers ORDER BY SortOrder")
Response.Write "<table border=1 cellpadding=5><tr><th>TierCode</th><th>Updated TierName</th><th>Hex</th></tr>"
Do While Not rs.EOF
    rawName = rs("TierName") & ""
    Response.Write "<tr><td>" & rs("TierCode") & "</td><td>" & rawName & "</td><td>"
    For j = 1 To Len(rawName)
        Response.Write Right("0" & Hex(AscW(Mid(rawName, j, 1))), 4) & " "
    Next
    Response.Write "</td></tr>"
    rs.MoveNext
Loop
rs.Close : conn.Close : Set rs = Nothing : Set conn = Nothing
%>
</table>

<p>Expected: 银卡会员 金卡会员 钻石会员 黑金会员</p>
<p>Hardcoded verify: 银卡会员 金卡会员 钻石会员 黑金会员 ✓</p>

</body>
</html>
