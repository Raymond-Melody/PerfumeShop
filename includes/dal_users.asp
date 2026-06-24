<%
' ============================================
' V15.0 DAL - 用户数据访问层
' 依赖: dal.asp, connection.asp
' ============================================

' ============================================
' 根据用户名获取用户
' ============================================
Function DAL_Users_GetByUsername(username)
    Dim sql, params(0)
    sql = "SELECT * FROM Users WHERE Username=@Username AND IsActive <> 0"
    params(0) = Array("@Username", DAL_adVarChar, 50, username)
    Set DAL_Users_GetByUsername = DAL_GetRow(sql, params)
End Function

' ============================================
' 根据ID获取用户
' ============================================
Function DAL_Users_GetByID(userId)
    Dim sql, params(0)
    sql = "SELECT * FROM Users WHERE UserID=@UserID"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    Set DAL_Users_GetByID = DAL_GetRow(sql, params)
End Function

' ============================================
' 根据邮箱获取用户
' ============================================
Function DAL_Users_GetByEmail(email)
    Dim sql, params(0)
    sql = "SELECT * FROM Users WHERE Email=@Email AND IsActive <> 0"
    params(0) = Array("@Email", DAL_adVarChar, 100, email)
    Set DAL_Users_GetByEmail = DAL_GetRow(sql, params)
End Function

' ============================================
' 检查用户名/邮箱是否已存在
' ============================================
Function DAL_Users_Exists(username, email)
    Dim sql, params(1), count
    sql = "SELECT COUNT(*) FROM Users WHERE Username=@Username OR Email=@Email"
    params(0) = Array("@Username", DAL_adVarChar, 50, username)
    params(1) = Array("@Email", DAL_adVarChar, 100, email)
    count = CLng(DAL_GetScalar(sql, params, 0))
    DAL_Users_Exists = (count > 0)
End Function

' ============================================
' 用户注册
' ============================================
Function DAL_Users_Register(username, email, passwordHash, fullName, phone)
    Dim sql, fields(4), params(4), newId
    
    fields(0) = "Username"
    fields(1) = "Email"
    fields(2) = "PasswordHash"
    fields(3) = "FullName"
    fields(4) = "Phone"
    
    params(0) = Array("@Username", DAL_adVarChar, 50, username)
    params(1) = Array("@Email", DAL_adVarChar, 100, email)
    params(2) = Array("@PasswordHash", DAL_adVarChar, 255, passwordHash)
    params(3) = Array("@FullName", DAL_adVarChar, 100, fullName)
    params(4) = Array("@Phone", DAL_adVarChar, 20, phone)
    
    newId = DAL_Insert("Users", fields, params)
    DAL_Users_Register = newId
End Function

' ============================================
' 更新用户最后登录时间
' ============================================
Sub DAL_Users_UpdateLastLogin(userId)
    Dim sql, params(0)
    sql = "UPDATE Users SET LastLoginAt=GETDATE() WHERE UserID=@UserID"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    DAL_Execute sql, params
End Sub

' ============================================
' 更新用户密码
' ============================================
Function DAL_Users_UpdatePassword(userId, newPasswordHash)
    Dim sql, params(1)
    sql = "UPDATE Users SET PasswordHash=@PasswordHash WHERE UserID=@UserID"
    params(0) = Array("@PasswordHash", DAL_adVarChar, 255, newPasswordHash)
    params(1) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    DAL_Users_UpdatePassword = (DAL_Execute(sql, params) >= 0)
End Function

' ============================================
' 获取用户积分
' ============================================
Function DAL_Users_GetPoints(userId)
    Dim sql, params(0)
    sql = "SELECT ISNULL(AvailablePoints, ISNULL(Points, 0)) FROM UserPoints WHERE UserID=@UserID"
    params(0) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    DAL_Users_GetPoints = CLng(DAL_GetScalar(sql, params, 0))
End Function

' ============================================
' 添加积分
' ============================================
Sub DAL_Users_AddPoints(userId, points, reason)
    ' 更新或插入 UserPoints
    Dim sql, params(3)
    sql = "UPDATE UserPoints SET AvailablePoints=ISNULL(AvailablePoints,0)+@Points, " & _
          "TotalPoints=ISNULL(TotalPoints,0)+@Points, LastUpdatedAt=GETDATE() " & _
          "WHERE UserID=@UserID; " & _
          "IF @@ROWCOUNT=0 INSERT INTO UserPoints (UserID, AvailablePoints, TotalPoints, LastUpdatedAt) " & _
          "VALUES (@UserID, @Points, @Points, GETDATE())"
    params(0) = Array("@Points", DAL_adInteger, 0, CLng(points))
    params(1) = Array("@UserID", DAL_adInteger, 0, CLng(userId))
    DAL_Execute sql, params
End Sub

' ============================================
' 获取用户列表（分页）
' ============================================
Function DAL_Users_GetList(search, page, pageSize, ByRef pageInfo)
    Dim sql, whereClause
    sql = "SELECT UserID, Username, Email, FullName, Phone, Points, " & _
          "CreatedAt, LastLoginAt, IsActive FROM Users WHERE 1=1"
    If Not IsNull(search) And search <> "" Then
        sql = sql & " AND (Username LIKE '%' + @Search + '%' OR Email LIKE '%' + @Search + '%' OR FullName LIKE '%' + @Search + '%')"
        Set DAL_Users_GetList = DAL_GetListPaged(sql, Array(Array("@Search", DAL_adVarChar, 100, search)), page, pageSize, pageInfo)
    Else
        Set DAL_Users_GetList = DAL_GetListPaged(sql, Null, page, pageSize, pageInfo)
    End If
End Function

' ============================================
' 管理员登录验证 (使用参数化查询)
' ============================================
Function DAL_Users_AdminLogin(username, passwordHash)
    Dim sql, params(1)
    sql = "SELECT AdminID, Username, RoleID, IsActive FROM AdminUsers " & _
          "WHERE Username=@Username AND PasswordHash=@PasswordHash AND IsActive <> 0"
    params(0) = Array("@Username", DAL_adVarChar, 50, username)
    params(1) = Array("@PasswordHash", DAL_adVarChar, 255, passwordHash)
    Set DAL_Users_AdminLogin = DAL_GetRow(sql, params)
End Function
%>