Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Server=localhost\YOURPERFUME;Database=PerfumeShop;Integrated Security=SSPI;"
Set rs = conn.Execute("SELECT UserID, Username, [Password], IsActive FROM Users WHERE Username = 'Raymond'")
If rs.EOF Then
    WScript.Echo "NOT FOUND"
Else
    WScript.Echo "UserID: " & rs("UserID")
    WScript.Echo "Username: " & rs("Username")
    pwd = rs("Password")
    If IsNull(pwd) Then
        WScript.Echo "Password: NULL"
    Else
        WScript.Echo "Password: " & pwd
        WScript.Echo "Length: " & Len(pwd)
    End If
    WScript.Echo "IsActive: " & rs("IsActive")
End If
rs.Close
conn.Close
